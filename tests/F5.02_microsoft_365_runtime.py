"""FB2-FB8: loopback Graph protocol for mail, watch, calendar, OneDrive and Teams, errors, cancellation and both runtimes.

Every provider exchange targets the loopback double in graph_fixture.py. The
wallet and its obviously fake credentials exist only in disposable test storage.
"""

import base64
from datetime import datetime, timedelta, timezone
import hashlib
import json
from pathlib import Path
import sys
import tempfile
import threading
import time

ROOT = Path(__file__).resolve().parents[3]
sys.path[:0] = [str(ROOT), str(ROOT / "tests"), str(Path(__file__).resolve().parent)]
from blocs.microsoft_365 import files, listener, logic, operations, storage, transport, watch
from blocs.microsoft_365.block import Microsoft365Block
from block_test_packages import install_test_package, prepare_release_run, surface_payload
from graph_fixture import (ACCESS, ATTACHMENT_BYTES, CHANNEL, DOCUMENTS, FOLDERS, FRIEND, GRANT, ME, OTHER_CHANNEL,
                           PARTNER, PDF_BYTES, PHOTO_BYTES, PUBLIC_GRANT, REF, REFRESHED, REPORTS, ROTATED, STRANGER, TEAM,
                           TENANT, error, graph_api, settings, stamp)
from ui_smoke_common import (create_project_api, create_run_api, data_edge, display_node, graph_payload, graph_storage_dir,
                             http_json, isolated_server, stop_run_api, text_node, wait_for_run_predicate,
                             wait_for_run_terminal)

TOKEN = {"access_token": ACCESS}
M1, M2, M3 = "AAMkFixtureMessage0001AAAA=", "AAMkFixtureMessage0002AAAA=", "AAMkFixtureMessage0003AAAA="
QUOTA = {"status": 429, "body": error(429, "TooManyRequests", "Too many requests.")}
SOLO, TEAM_SYNC, PARTNER_REVIEW, INVITED = ("AAMkFixtureEventSolo00001=", "AAMkFixtureEventTeam00002=",
                                            "AAMkFixtureEventPartner003=", "AAMkFixtureEventInvited004=")
SECRETS = (ACCESS, REFRESHED, ROTATED, GRANT["client_secret"], GRANT["refresh_token"])


def config_of(api, **extra):
    return logic.configuration(settings(api, **extra))


def run(config, value, secret=TOKEN, cancel=lambda: False, storage_dir=None, port="request", now=None):
    """Execute one event through the same local checks, session and operations as the block."""
    moment = now or datetime.now(timezone.utc)
    request = logic.parse_request(value, config, moment) if port == "request" else logic.file_request(value, config)
    extra = {"storage_dir": storage_dir, "now": moment}
    upload = None
    if request["operation"] == "mail_send":
        extra["attachments"] = files.read_attachments(request["attachments"], config)
    if request["operation"] == "file_upload":
        upload = extra["upload"] = files.open_upload(request["path"], config)
        request.setdefault("name", logic.drive_name(upload.name, "The local file name"))
    try:
        return operations.run(transport.Session(config, secret, cancel), request, config, **extra)
    finally:
        if upload is not None:
            upload.close()


def clean(data):
    text = json.dumps(data)
    assert not any(secret in text for secret in SECRETS), "A credential leaked into a result"


def failure(call, code):
    try:
        call()
    except logic.GraphError as exc:
        assert exc.code == code, (code, exc.code, str(exc))
        clean(exc.envelope())
        return exc
    raise AssertionError("Expected " + code)


def subjects(values):
    return [item["subject"] for item in values["items"]]


def until(predicate, timeout=10.0, label="condition not reached"):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if predicate():
            return
        time.sleep(0.05)
    raise AssertionError(label)


def mail_protocol():
    """# - FB2 - # - FB7 - listings, exact page links, filters and search, bodies, grants, send, move, mark and attachments."""
    with graph_api() as api, tempfile.TemporaryDirectory() as folder:
        base = Path(folder).resolve()
        node, attach = base / "node", base / "attach"
        node.mkdir()
        attach.mkdir()
        (attach / "report.txt").write_bytes(b"weekly\n")
        reader = config_of(api, page_size=5, allowed_operations=["mail_list", "mail_get", "mail_attachment_download"])
        listed = run(reader, "go")
        assert subjects(listed) == [f"Report {index:02d}" for index in range(1, 13)], subjects(listed)
        assert listed["result"] == {"request_id": None, "operation": "mail_list", "count": 12, "truncated": False,
                                    "stop_reason": "end", "next_page_link": None, "pages": 3, "folder": "inbox", "requests": 3}
        first = api.calls[-3]
        assert first["path"] == "/v1.0/me/mailFolders/inbox/messages" and first["query"]["$top"] == "5"
        assert first["query"]["$select"] == logic.MESSAGE_SELECT and first["query"]["$orderby"] == "receivedDateTime desc"
        assert first["auth"] == "Bearer " + ACCESS and first["accept_encoding"] == "identity"
        assert listed["items"][0]["from"] == {"name": "Sender", "address": PARTNER} and listed["items"][0]["has_attachments"]
        # A limit never cuts a page: the exact next_page_link resumes without gap or duplicate.
        limited = run(reader, {"limit": 7, "request_id": "p-1"})
        assert limited["result"]["count"] == 5 and limited["result"]["stop_reason"] == "limit"
        chained, link = subjects(limited), limited["result"]["next_page_link"]
        while link:
            follow = run(reader, {"limit": 7, "page_link": link})
            chained += subjects(follow)
            link = follow["result"]["next_page_link"]
        assert chained == subjects(listed)
        capped = run(config_of(api, page_size=2, max_pages=2), {"limit": 50})
        assert capped["result"]["stop_reason"] == "max_pages" and capped["result"]["count"] == 4
        unread = run(reader, {"unread_only": True, "received_after": stamp(api.now - timedelta(hours=1))})
        assert subjects(unread) == ["Report 01", "Report 03", "Report 05"]
        assert api.calls[-1]["query"]["$filter"].startswith("receivedDateTime ge ")
        searched = run(reader, {"search": "budget line 7"})
        assert subjects(searched) == ["Report 07"] and "$orderby" not in api.calls[-1]["query"]
        assert api.calls[-1]["query"]["$search"] == '"budget line 7"'
        narrow = run(reader, {"select": ["subject"], "limit": 2})
        assert narrow["items"][0]["subject"] == "Report 01" and narrow["items"][0]["from"] is None
        assert api.calls[-1]["query"]["$select"] == "id,subject"
        assert subjects(run(reader, {"folder": "archive"})) == ["Old archive"]
        missing = failure(lambda: run(reader, {"folder": "AAMkFixtureFolderMissing00="}), "not_found")
        assert missing.provider_code == "ErrorItemNotFound" and missing.provider_request_id == "fixture-request-404"
        raw = run(reader, {"filter": "isRead eq false", "limit": 20})
        assert len(raw["items"]) == 6 and "$orderby" not in api.calls[-1]["query"]
        # Full message: bounded body in the chosen format and attachment metadata, never contents.
        got = run(reader, {"operation": "mail_get", "message_id": M1, "request_id": "g-1"})["result"]
        assert got["request_id"] == "g-1" and got["message"]["body"]["content"].startswith("Body of report 01")
        assert [item["kind"] for item in got["message"]["attachments"]] == ["file", "item"]
        assert api.calls[-1]["prefer"] == 'outlook.body-content-type="text"'
        assert api.calls[-1]["query"]["$expand"] == "attachments($select=id,name,contentType,size,isInline)"
        long_id = api.deliver("Long body", body="x" * 1000)
        html = run(logic.configuration({**reader, "max_body_chars": 256}),
                   {"operation": "mail_get", "message_id": long_id, "body_format": "html"})["result"]["message"]["body"]
        assert html == {"content_type": "html", "content": "<p>" + "x" * 253, "truncated": True, "length": 1007}
        # Refresh-token grants: confidential and public clients, tenant in the path, the rotated token is ignored.
        before = len(api.calls)
        assert run(reader, {"limit": 1}, GRANT)["result"]["requests"] == 2
        token = api.calls[before]
        assert token["kind"] == "token" and token["tenant"] == TENANT and token["has_secret"]
        assert token["body"]["scope"] == logic.DEFAULT_SCOPE and token["body"]["grant_type"] == "refresh_token"
        assert api.calls[before + 1]["auth"] == "Bearer " + REFRESHED
        run(reader, {"limit": 1}, PUBLIC_GRANT)
        assert api.calls[-2]["kind"] == "token" and not api.calls[-2]["has_secret"]
        refused = failure(lambda: run(reader, "go", {**GRANT, "refresh_token": "revoked-fixture"}), "auth_failed")
        assert refused.provider_code == "invalid_grant" and refused.http_status == 400 and "AADSTS70000" in refused.provider_message
        other = failure(lambda: run(config_of(api, tenant_id="other.example"), "go", GRANT), "auth_failed")
        assert other.provider_code == "invalid_request"
        # sendMail: one request, recipients and attachments as documented, accepted with HTTP 202.
        writer = config_of(api, allow_writes=True, allow_send=True, allowed_recipients=[FRIEND, "@partner.example"],
                           attachment_root=str(attach), allowed_operations=["mail_send", "mail_move", "mail_mark_read"])
        sent = run(writer, {"operation": "mail_send", "to": [FRIEND], "cc": [{"address": PARTNER, "name": "Partner"}],
                            "subject": "Weekly", "body": "<b>Done</b>", "body_format": "html", "attachments": ["report.txt"],
                            "request_id": "s-1"})["result"]
        assert sent == {"request_id": "s-1", "operation": "mail_send", "accepted": True, "to": [FRIEND], "cc": [PARTNER],
                        "bcc_count": 0, "attachments": ["report.txt"], "saved_to_sent_items": True, "requests": 1}
        message = api.sent[-1]["message"]
        assert api.sent[-1]["saveToSentItems"] is True and message["subject"] == "Weekly"
        assert message["toRecipients"] == [{"emailAddress": {"address": FRIEND}}]
        assert message["ccRecipients"] == [{"emailAddress": {"address": PARTNER, "name": "Partner"}}]
        assert message["body"] == {"contentType": "html", "content": "<b>Done</b>"}
        assert message["attachments"] == [{"@odata.type": "#microsoft.graph.fileAttachment", "name": "report.txt",
                                           "contentType": "text/plain",
                                           "contentBytes": base64.b64encode(b"weekly\n").decode("ascii")}]
        before = len(api.calls)
        failure(lambda: run(writer, {"operation": "mail_send", "to": [STRANGER]}), "recipient_not_allowed")
        failure(lambda: run(writer, {"operation": "mail_send", "to": [FRIEND], "attachments": ["../node"]}), "path_denied")
        assert len(api.calls) == before, "Refused sends contact nobody"
        moved = run(writer, {"operation": "mail_move", "message_id": M2, "destination": "archive"})["result"]
        assert moved["message"]["id"] != M2 and moved["message"]["parent_folder_id"] == FOLDERS["archive"]
        assert api.calls[-1]["body"] == {"destinationId": "archive"}
        marked = run(writer, {"operation": "mail_mark_read", "message_id": M1, "is_read": True})["result"]
        assert marked["message"]["is_read"] is True and api.messages[M1].is_read is True
        assert failure(lambda: run(writer, {"operation": "mail_move", "message_id": M2, "destination": "archive"}),
                       "not_found").provider_code == "ErrorItemNotFound"
        # Attachments: file attachments stream to node storage; item attachments are refused.
        saved = run(reader, {"operation": "mail_attachment_download", "message_id": M1,
                             "attachment_id": "AAMkFixtureAttachment01AA="}, storage_dir=lambda: node)
        path = Path(saved["file"])
        assert path.read_bytes() == ATTACHMENT_BYTES and path.name == "invoice.txt" and path.parent.parent == node / "downloads"
        assert saved["result"]["sha256"] == hashlib.sha256(ATTACHMENT_BYTES).hexdigest()
        assert api.kinds()[-2:] == ["attachment", "attachment_value"]
        before = len(api.calls)
        failure(lambda: run(reader, {"operation": "mail_attachment_download", "message_id": M1,
                                     "attachment_id": "AAMkFixtureAttachment02AA="}, storage_dir=lambda: node), "not_downloadable")
        assert api.kinds()[before:] == ["attachment"] and not list(node.rglob(".partial-*"))
        clean(listed)
    print("[ok] Microsoft 365 mail listing, page links, filters, bodies, grants, send, move, mark and attachments", flush=True)


def watch_protocol():
    """# - FB6 - delta cursor: no flood on first start, new mail once, bounded polls, resync, One Shot poll and lock."""
    with graph_api() as api, tempfile.TemporaryDirectory() as folder:
        node = Path(folder).resolve()
        config = config_of(api, watch_mail=True, page_size=5, watch_max_items=5,
                           allowed_operations=["mail_list", "mail_watch"])

        def poll(document, now=None, conf=config):
            return watch.poll(transport.Session(conf, TOKEN), conf, document, now or datetime.now(timezone.utc))

        document, found, summary = poll(logic.empty_state(), api.now)
        assert found == [] and summary == {"folder": "inbox", "count": 0, "pages": 1, "skipped": 0, "started": True,
                                           "resynced": False, "caught_up": True}
        call = api.calls[-1]
        assert call["kind"] == "delta" and call["query"]["$filter"] == "receivedDateTime ge " + stamp(api.now)
        assert call["prefer"] == "odata.maxpagesize=5" and call["query"]["$select"] == logic.MESSAGE_SELECT
        fresh = api.deliver("Fresh 1")
        document, found, summary = poll(document)
        assert [item["subject"] for item in found] == ["Fresh 1"] and not summary["started"] and summary["caught_up"]
        assert "$deltatoken" in api.calls[-1]["raw_query"] and api.calls[-1]["path"] == "/v1.0/me/mailfolders('inbox')/messages/delta"
        # Read-flag changes and moves of handed-over mail are not new mail.
        api.touch(fresh)
        document, found, summary = poll(document)
        assert found == [] and summary["skipped"] == 1
        # A burst is handed over in bounded polls, each message exactly once.
        burst = [api.deliver(f"Burst {index:02d}") for index in range(12)]
        batches = []
        for _ in range(3):
            document, found, summary = poll(document)
            batches.append([item["subject"] for item in found])
        assert [len(batch) for batch in batches] == [5, 5, 2] and summary["caught_up"]
        assert sorted(sum(batches, [])) == [f"Burst {index:02d}" for index in range(12)] and len(burst) == 12
        # emit_existing: a fresh cursor hands over the recent mail once, bounded by watch_max_items.
        recent = config_of(api, watch_mail=True, emit_existing=True, existing_since_hours=1, page_size=50, watch_max_items=50)
        expected = sorted(m.subject for m in api.messages.values()
                          if m.folder == "inbox" and m.received >= datetime.now(timezone.utc) - timedelta(hours=1))
        existing, found, _summary = poll(logic.empty_state(), conf=recent)
        assert sorted(item["subject"] for item in found) == expected and "Report 05" in expected and "Report 07" not in expected
        assert poll(existing, conf=recent)[1] == []
        # HTTP 410: the round restarts from the newest handed-over time; nothing is replayed.
        api.script["delta"] = [{"status": 410, "body": error(410, "SyncStateNotFound", "The sync state has expired.")}]
        api.deliver("After reset")
        before = len(api.calls)
        document, found, summary = poll(document)
        assert [item["subject"] for item in found] == ["After reset"] and summary["resynced"] and summary["started"] is False
        assert "$deltatoken" in api.calls[before]["raw_query"] and api.calls[before + 1]["query"]["$filter"].startswith(
            "receivedDateTime ge "), "After HTTP 410 the round restarts from the newest handed-over time"
        # A cursor for another host is never followed: the watch restarts from its last position.
        tampered = json.loads(json.dumps(document))
        entry = tampered["watches"][logic.watch_key(config)]
        entry["link"] = "https://evil.example/v1.0/me/mailFolders/inbox/messages/delta?$deltatoken=x"
        before = len(api.calls)
        tampered, found, summary = poll(tampered)
        assert found == [] and summary["resynced"] and all(call["kind"] == "delta" for call in api.calls[before:])
        api.script["delta"] = [{"status": 200, "body": {"value": []}}]
        failure(lambda: poll(document), "protocol")
        # One Shot or a request: one poll with the cursor persisted in node storage under a lock.
        time.sleep(1.1)  # The cursor starts at whole seconds; keep earlier deliveries out of it.
        first = run(config, {"operation": "mail_watch", "request_id": "w-1"}, storage_dir=lambda: node)
        assert first["items"] == [] and first["result"]["started"] and first["result"]["source"] == "request"
        api.deliver("Polled once")
        second = run(config, {"operation": "mail_watch"}, storage_dir=lambda: node)
        assert subjects(second) == ["Polled once"] and run(config, {"operation": "mail_watch"}, storage_dir=lambda: node)["items"] == []
        state = (node / storage.STATE_NAME).read_text()
        assert not any(secret in state for secret in SECRETS) and REF not in state and "Polled once" not in state
        holder = storage.WatchStore(node)
        assert holder.acquire()
        try:
            failure(lambda: run(config, {"operation": "mail_watch"}, storage_dir=lambda: node), "busy")
        finally:
            holder.close()
        failure(lambda: run(config, {"operation": "mail_watch"}), "storage")
    print("[ok] Microsoft 365 mail watch cursor, bounded polls, resync, One Shot poll and lock", flush=True)


def listener_protocol():
    """# - FB6 - the Active Runtime watcher persists before hand-off, never replays after restart and stops promptly."""
    with graph_api() as api, tempfile.TemporaryDirectory() as folder:
        node = Path(folder).resolve()
        # The 60 s minimum is a settings rule; the watcher object itself is driven faster here.
        fast = {**config_of(api, watch_mail=True), "watch_interval_sec": 0.3}

        def start(emit, secret=lambda: json.dumps(GRANT), directory=lambda: str(node)):
            stop = threading.Event()
            watcher = listener.MailWatcher(fast, resolve_secret=secret, storage_dir=directory, emit=emit, stop=stop.is_set)
            thread = threading.Thread(target=watcher.run, name="microsoft-365-test-watcher")
            thread.start()
            return stop, thread

        def finish(stop, thread):
            started = time.monotonic()
            stop.set()
            thread.join(3)
            assert not thread.is_alive() and time.monotonic() - started < 1.5, "The watcher must stop promptly"

        emitted = []
        stop, thread = start(lambda kind, value, stats: emitted.append((kind, value, stats)))
        try:
            until(lambda: any(kind == "status" for kind, _v, _s in emitted), label="first poll missing")
            assert not any(kind == "items" for kind, _v, _s in emitted), "The first start never floods"
            api.deliver("Live 1")
            until(lambda: any(kind == "items" for kind, _v, _s in emitted), label="new mail not handed over")
            kind, (messages, result), stats = next(item for item in emitted if item[0] == "items")
            assert [m["subject"] for m in messages] == ["Live 1"] and result["source"] == "watch" and result["count"] == 1
            assert stats["state"] == "watching" and stats["received"] == 1
            time.sleep(1)
            assert sum(1 for item in emitted if item[0] == "items") == 1, "Each message is handed over once"
            assert api.kinds().count("token") == 1, "The access token is reused between polls"
        finally:
            finish(stop, thread)
        api.deliver("While stopped")
        emitted.clear()
        stop, thread = start(lambda kind, value, stats: emitted.append((kind, value, stats)))
        try:
            until(lambda: any(kind == "items" for kind, _v, _s in emitted), label="restart lost mail")
            assert [m["subject"] for m in next(v for k, v, _s in emitted if k == "items")[0]] == ["While stopped"]
            # A request poll while the listener holds the cursor is refused, not interleaved.
            failure(lambda: run(logic.configuration({**fast, "watch_interval_sec": 60,
                                                     "allowed_operations": ["mail_watch"]}),
                                {"operation": "mail_watch"}, storage_dir=lambda: node), "busy")
        finally:
            finish(stop, thread)
        # A refused hand-off restores the stored cursor: the next Run hands the same mail over again.
        api.deliver("Refused once")
        refused = []

        def refusing(kind, value, stats):
            refused.append(kind)
            if kind == "items":
                raise RuntimeError("runtime_listeners.queue_full")

        stop, thread = start(refusing)
        try:
            until(lambda: "items" in refused, label="refused hand-off not attempted")
        finally:
            finish(stop, thread)
        emitted.clear()
        stop, thread = start(lambda kind, value, stats: emitted.append((kind, value, stats)))
        try:
            until(lambda: any(kind == "items" for kind, _v, _s in emitted), label="refused mail lost")
            assert [m["subject"] for m in next(v for k, v, _s in emitted if k == "items")[0]] == ["Refused once"]
        finally:
            finish(stop, thread)
        # Failures are reported once per distinct code; Stop interrupts a held poll at once.
        errors = []
        stop, thread = start(lambda kind, value, stats: errors.append((kind, value)), secret=lambda: "{broken")
        try:
            until(lambda: errors, label="secret error missing")
            time.sleep(1)
            assert [value["code"] for kind, value in errors if kind == "error"] == ["secret"] and len(errors) == 1
        finally:
            finish(stop, thread)
        errors.clear()
        stop, thread = start(lambda kind, value, stats: errors.append((kind, value)),
                             directory=lambda: (_ for _ in ()).throw(RuntimeError("block_storage.missing_scope")))
        try:
            until(lambda: errors, label="storage error missing")
            assert errors[0][1]["code"] == "storage" and errors[0][1]["source"] == "watch"
        finally:
            finish(stop, thread)
        api.script["delta"] = ["hold"]
        stop, thread = start(lambda kind, value, stats: None)
        try:
            until(lambda: api.holding == 1, label="held poll not reached")
        finally:
            finish(stop, thread)
            api.release.set()
        time.sleep(0.2)
        assert not [item for item in threading.enumerate() if item.name == transport.THREAD_NAME]
    print("[ok] Microsoft 365 watcher hand-off, restart without replay, refused hand-off, errors and Stop", flush=True)


def calendar_protocol():
    """# - FB3 - calendarView windows and time zones, events with invitations, and gated updates and deletions."""
    with graph_api() as api:
        writer = config_of(api, allow_writes=True, allow_send=True, allow_delete=True, allowed_recipients=[FRIEND, ME],
                           page_size=2, allowed_operations=["events_list", "event_create", "event_update", "event_delete"])
        window = {"operation": "events_list", "start": stamp(api.now), "end": stamp(api.now + timedelta(days=3))}
        listed = run(writer, window)
        assert subjects(listed) == ["Focus time", "Team sync", "Partner review", "Invited call"]
        assert listed["result"]["pages"] == 2 and listed["result"]["time_zone"] == "UTC" and listed["result"]["calendar_id"] is None
        first, second = api.calls[-2:]
        assert first["query"]["startDateTime"] == stamp(api.now) and first["query"]["$orderby"] == "start/dateTime"
        assert first["prefer"] == second["prefer"] == 'outlook.timezone="UTC"' and "$skiptoken" in second["raw_query"]
        team = listed["items"][1]
        assert team["attendees"] == [{"name": "colleague", "address": FRIEND, "type": "required", "response": "none"}]
        assert team["is_organizer"] is True and team["start"]["time_zone"] == "UTC"
        paris = run(logic.configuration({**writer, "time_zone": "Europe/Paris"}), window)
        assert paris["items"][0]["start"]["time_zone"] == "Europe/Paris" and api.calls[-1]["prefer"] == 'outlook.timezone="Europe/Paris"'
        assert run(writer, {"operation": "events_list", "start": "2020-01-01", "end": "2020-01-02"})["items"] == []
        created = run(writer, {"operation": "event_create", "subject": "Planning", "start": "2026-11-02T09:00",
                               "end": "2026-11-02T10:00", "attendees": [FRIEND], "location": "Room A",
                               "request_id": "e-1"})["result"]
        assert created["event"]["subject"] == "Planning" and created["invitations_sent"] and ("invitation", FRIEND) in api.notifications
        body = api.calls[-1]["body"]
        assert body["start"] == {"dateTime": "2026-11-02T09:00:00", "timeZone": "UTC"} and len(body["transactionId"]) == 36
        assert body["location"] == {"displayName": "Room A"} and body["subject"] == "Planning"
        api.script["event_create"] = [{**QUOTA, "headers": {"Retry-After": "1"}}]
        before = len(api.events)
        retried = run(writer, {"operation": "event_create", "subject": "Retry", "start": "2026-11-03T09:00",
                               "end": "2026-11-03T09:30"})["result"]
        assert len(api.events) == before + 1 and retried["requests"] == 2
        assert api.calls[-1]["body"]["transactionId"] == api.calls[-2]["body"]["transactionId"]
        notified = len(api.notifications)
        assert run(writer, {"operation": "event_update", "event_id": SOLO, "subject": "Deep work"})["result"]["updates_sent"] is False
        assert api.events[SOLO].subject == "Deep work" and len(api.notifications) == notified
        updated = run(writer, {"operation": "event_update", "event_id": TEAM_SYNC, "start": "2026-11-04T09:00",
                               "end": "2026-11-04T09:30"})["result"]
        assert updated["updates_sent"] and api.notifications[-1] == ("update", FRIEND)
        before = len(api.calls)
        failure(lambda: run(writer, {"operation": "event_update", "event_id": PARTNER_REVIEW, "subject": "x"}),
                "recipient_not_allowed")
        assert api.kinds()[before:] == ["event_get"], "A meeting with a non-allowed attendee is never changed"
        quiet = logic.configuration({**writer, "allow_send": False})
        failure(lambda: run(quiet, {"operation": "event_update", "event_id": TEAM_SYNC, "subject": "x"}), "send_disabled")
        assert run(quiet, {"operation": "event_update", "event_id": INVITED, "subject": "Mine"})["result"]["updates_sent"] is False
        failure(lambda: run(quiet, {"operation": "event_delete", "event_id": TEAM_SYNC}), "send_disabled")
        assert TEAM_SYNC in api.events
        deleted = run(writer, {"operation": "event_delete", "event_id": SOLO})["result"]
        assert deleted == {"request_id": None, "operation": "event_delete", "event_id": SOLO, "deleted": True,
                           "recoverable": True, "cancellations_sent": False, "requests": 2}
        assert api.deleted_events == [SOLO] and api.calls[-1]["method"] == "DELETE"
        cancelled = run(writer, {"operation": "event_delete", "event_id": TEAM_SYNC})["result"]
        assert cancelled["cancellations_sent"] and api.notifications[-1] == ("cancellation", FRIEND)
        failure(lambda: run(writer, {"operation": "event_delete", "event_id": SOLO}), "not_found")
    print("[ok] Microsoft 365 calendar windows, invitations and gated updates and deletions", flush=True)


def drive_protocol():
    """# - FB4 - # - FB5 - OneDrive listings, verified downloads, simple and session uploads, folders and Teams posts."""
    with graph_api() as api, tempfile.TemporaryDirectory() as folder:
        base = Path(folder).resolve()
        node, outgoing = base / "node", base / "outgoing"
        node.mkdir()
        outgoing.mkdir()
        (outgoing / "report.txt").write_bytes(b"weekly report\n")
        big = bytes(range(256)) * 17579 + b"tail"
        (outgoing / "big.bin").write_bytes(big)
        reader = config_of(api, allowed_operations=["files_list", "file_download"])
        listed = run(reader, {"operation": "files_list"})
        assert [item["name"] for item in listed["items"]] == ["Bulk", "Documents", "photo.png"]
        assert api.calls[-1]["path"] == "/v1.0/me/drive/root/children" and api.calls[-1]["query"]["$select"] == logic.DRIVE_SELECT
        paged = logic.configuration({**reader, "page_size": 10, "max_pages": 2})
        first = run(paged, {"operation": "files_list", "path": "Bulk", "request_id": "f-1"})
        assert first["result"]["count"] == 20 and first["result"]["stop_reason"] == "max_pages"
        rest = run(paged, {"operation": "files_list", "page_link": first["result"]["next_page_link"]})
        assert [item["name"] for item in first["items"] + rest["items"]] == [f"note-{index:02d}.txt" for index in range(1, 26)]
        assert first["result"]["next_page_link"].startswith(api.base_url + "/v1.0/drives/b!fixturedrive/items/")
        assert [item["name"] for item in run(reader, {"operation": "files_list", "item_id": REPORTS})["items"]] == ["Q1 report.pdf"]
        nested = logic.configuration({**reader, "drive_folder": "Documents"})
        assert run(nested, {"operation": "files_list", "path": "Reports"})["items"][0]["name"] == "Q1 report.pdf"
        assert api.calls[-1]["path"] == "/v1.0/me/drive/root:/Documents/Reports:/children"
        assert failure(lambda: run(reader, {"operation": "files_list", "path": "Missing"}), "not_found").provider_code == "itemNotFound"
        got = run(reader, {"operation": "file_download", "path": "Documents/Reports/Q1 report.pdf"}, storage_dir=lambda: node)
        assert Path(got["file"]).read_bytes() == PDF_BYTES and got["result"]["sha256_verified"] is True
        assert api.kinds()[-2:] == ["item", "download"] and api.calls[-1]["auth"] is None
        photo = run(reader, {"operation": "file_download", "item_id": "01FIXTUREPHOTOPNG000000000000000"}, storage_dir=lambda: node)
        assert Path(photo["file"]).read_bytes() == PHOTO_BYTES and photo["result"]["sha256_verified"] is None
        before = len(api.calls)
        failure(lambda: run(reader, {"operation": "file_download", "item_id": DOCUMENTS}, storage_dir=lambda: node),
                "not_downloadable")
        failure(lambda: run(logic.configuration({**reader, "max_download_bytes": 1024}),
                            {"operation": "file_download", "path": "Documents/Reports/Q1 report.pdf"}, storage_dir=lambda: node),
                "file_too_large")
        assert api.kinds()[before:] == ["item", "item"], "Refused downloads read metadata only"
        api.script["body"] = ["corrupt"]
        assert "SHA-256" in str(failure(lambda: run(reader, {"operation": "file_download", "item_id": "01FIXTUREQ1REPORTPDF000000000000"},
                                                    storage_dir=lambda: node), "protocol"))
        api.script["body"] = ["gzip"]
        failure(lambda: run(reader, {"operation": "file_download", "item_id": "01FIXTUREQ1REPORTPDF000000000000"},
                            storage_dir=lambda: node), "protocol")
        api.script["body"] = ["short"]
        before = len(api.calls)
        assert Path(run(reader, {"operation": "file_download", "item_id": "01FIXTUREQ1REPORTPDF000000000000"},
                        storage_dir=lambda: node)["file"]).read_bytes() == PDF_BYTES
        assert api.kinds()[before:] == ["item", "download", "download"], "A truncated read is retried from an empty file"
        assert not list(node.rglob(".partial-*")) and len(list((node / "downloads").iterdir())) == 3
        # Uploads stay inside drive_folder; conflicts never overwrite unless replace is allowed.
        writer = config_of(api, allow_writes=True, upload_root=str(outgoing), drive_folder="Documents", upload_chunk_bytes=327680 * 4,
                           allowed_operations=["files_list", "file_upload", "folder_create"])
        uploaded = run(writer, str(outgoing / "report.txt") + "\n", port="file")["result"]
        assert uploaded["item"]["name"] == "report.txt" and uploaded["method"] == "simple" and uploaded["sha256_verified"]
        call = api.calls[-1]
        assert call["path"] == "/v1.0/me/drive/root:/Documents/report.txt:/content" and call["auth"] == "Bearer " + ACCESS
        assert call["raw_query"] == "@microsoft.graph.conflictBehavior=fail" and call["content_type"] == "text/plain"
        assert failure(lambda: run(writer, {"operation": "file_upload", "path": "report.txt"}), "conflict").provider_code == "nameAlreadyExists"
        renamed = run(logic.configuration({**writer, "conflict_behavior": "rename"}), {"operation": "file_upload", "path": "report.txt"})
        assert renamed["result"]["item"]["name"] == "report 1.txt"
        (outgoing / "report.txt").write_bytes(b"weekly report v2\n")
        replaced = run(logic.configuration({**writer, "conflict_behavior": "replace", "allow_delete": True}),
                       {"operation": "file_upload", "path": "report.txt"})["result"]
        assert replaced["item"]["id"] == uploaded["item"]["id"] and api.items[uploaded["item"]["id"]].content == b"weekly report v2\n"
        session = run(writer, {"operation": "file_upload", "path": "big.bin", "folder": "Reports", "name": "Big data.bin"})["result"]
        assert session["method"] == "session" and session["chunks"] == 4 and session["size"] == len(big)
        assert api.items[session["item"]["id"]].content == big and session["path"] == "Documents/Reports/Big data.bin"
        chunks = [call for call in api.calls if call["kind"] == "chunk"]
        assert [call["content_range"] for call in chunks] == [f"bytes {start}-{min(start + 1310720, len(big)) - 1}/{len(big)}"
                                                             for start in range(0, len(big), 1310720)]
        assert all(call["auth"] is None for call in chunks) and api.calls[-5]["kind"] == "create_session"
        assert api.calls[-5]["body"] == {"item": {"@microsoft.graph.conflictBehavior": "fail", "name": "Big data.bin"}}
        before = len(api.calls)
        failure(lambda: run(writer, {"operation": "file_upload", "path": "big.bin", "folder": "Reports", "name": "Big data.bin"}),
                "conflict")
        assert api.kinds()[before:] == ["create_session"]
        api.script["chunk"] = ["pass", {"status": 503, "body": error(503, "serviceNotAvailable", "Try later.")}]
        middle = failure(lambda: run(writer, {"operation": "file_upload", "path": "big.bin", "name": "b2.bin"}), "provider_unavailable")
        assert not middle.uncertain and api.kinds()[-1] == "session_delete" and not api.sessions
        api.script["commit"] = ["lost"]
        lost = failure(lambda: run(writer, {"operation": "file_upload", "path": "big.bin", "name": "b3.bin"}), "network")
        assert lost.uncertain and api.kinds()[-1] == "session_delete", "A committed final fragment cannot be confirmed"
        api.script["chunk"] = ["pass", "pass", "pass", "lost"]
        open_session = failure(lambda: run(writer, {"operation": "file_upload", "path": "big.bin", "name": "b4.bin"}), "network")
        assert not open_session.uncertain, "The session was still open, so nothing was created"
        api.script["receipt"] = ["badhash"]
        assert failure(lambda: run(writer, {"operation": "file_upload", "path": "report.txt", "name": "hash.txt"}), "protocol").uncertain
        before = len(api.calls)
        failure(lambda: run(logic.configuration({**writer, "max_upload_bytes": 1024}), {"operation": "file_upload", "path": "big.bin"}),
                "file_too_large")
        failure(lambda: run(writer, {"operation": "file_upload", "path": "../node"}), "path_denied")
        assert len(api.calls) == before
        made = run(writer, {"operation": "folder_create", "name": "Invoices"})["result"]
        assert made["created"] and made["folder"]["is_folder"] and made["parent"] == "Documents"
        assert api.calls[-1]["path"] == "/v1.0/me/drive/root:/Documents:/children"
        reused = run(writer, {"operation": "folder_create", "name": "Invoices"})["result"]
        assert not reused["created"] and reused["folder"]["id"] == made["folder"]["id"] and api.kinds()[-2:] == ["folder_create", "item"]
        failure(lambda: run(logic.configuration({**writer, "folder_exists": "fail"}), {"operation": "folder_create", "name": "Invoices"}),
                "conflict")
        assert "file with this name" in str(failure(lambda: run(writer, {"operation": "folder_create", "name": "report.txt"}), "conflict"))
        poster = config_of(api, allow_writes=True, allow_send=True, allowed_operations=["channel_message_send"],
                           allowed_channels=[TEAM + "/" + CHANNEL, TEAM + "/" + OTHER_CHANNEL, TEAM + "/19:fixturegone0003@thread.tacv2"])
        posted = run(poster, {"operation": "channel_message_send", "team_id": TEAM, "channel_id": CHANNEL, "subject": "CI",
                              "content": "<b>Build passed</b>", "content_type": "html"})["result"]
        assert posted["message"]["team_id"] == TEAM and posted["message"]["id"]
        assert api.channel_messages[-1] == (TEAM, CHANNEL, {"body": {"contentType": "html", "content": "<b>Build passed</b>"},
                                                            "subject": "CI"})
        assert api.calls[-1]["path"] == "/v1.0/teams/" + TEAM + "/channels/19%3Afixturegeneral0001%40thread.tacv2/messages"
        before = len(api.calls)
        failure(lambda: run(poster, {"operation": "channel_message_send", "content": "x"}), "invalid_input")
        failure(lambda: run(poster, {"operation": "channel_message_send", "team_id": TEAM, "channel_id": "19:x@thread.tacv2",
                                     "content": "x"}), "channel_not_allowed")
        assert len(api.calls) == before
        failure(lambda: run(poster, {"operation": "channel_message_send", "team_id": TEAM,
                                     "channel_id": "19:fixturegone0003@thread.tacv2", "content": "x"}), "not_found")
    print("[ok] Microsoft 365 OneDrive listings, downloads, uploads, folders and Teams posts", flush=True)


def error_protocol():
    """# - FB7 - 401/403/404/429/5xx, lost connections, uncertain sends never replayed, cancellation, deadline, TLS."""
    with graph_api() as api, tempfile.TemporaryDirectory() as folder:
        node = Path(folder).resolve()
        config = config_of(api, allow_writes=True, allow_send=True, allowed_recipients=[FRIEND], max_retry_wait_sec=2,
                           allowed_operations=["mail_list", "mail_send", "file_download"])
        send = {"operation": "mail_send", "to": [FRIEND], "subject": "x"}
        expired = failure(lambda: run(config, "go", {"access_token": "-".join(("expired", "fixture", "value"))}), "unauthorized")
        assert expired.provider_code == "InvalidAuthenticationToken" and not expired.uncertain
        api.script["mail_list"] = [{"status": 403, "body": error(403, "ErrorAccessDenied", "Access is denied.")}]
        assert failure(lambda: run(config, "go"), "forbidden").provider_code == "ErrorAccessDenied"
        api.script["mail_list"] = [{**QUOTA, "headers": {"Retry-After": "1"}}]
        started = time.monotonic()
        assert run(config, {"limit": 1})["result"]["requests"] == 2 and time.monotonic() - started >= 1
        api.script["send"] = [{**QUOTA, "headers": {"Retry-After": "1"}}]
        before = len(api.sent)
        assert run(config, send)["result"]["accepted"] and len(api.sent) == before + 1 and api.kinds()[-2:] == ["send", "send"]
        api.script["mail_list"] = [{**QUOTA, "headers": {"Retry-After": "30"}}]
        limited = failure(lambda: run(config, "go"), "rate_limited")
        assert limited.retry_after == 30 and limited.attempts == 1
        api.script["mail_list"] = [QUOTA] * 3
        assert failure(lambda: run(logic.configuration({**config, "max_retries": 0}), "go"), "rate_limited").attempts == 1
        api.script["mail_list"] = ["lost", {"status": 503, "body": error(503, "serviceNotAvailable", "Busy."),
                                            "headers": {"Retry-After": "1"}}]
        before = len(api.calls)
        assert run(config, {"limit": 1})["items"] and api.kinds()[before:] == ["mail_list"] * 3
        for scripted, code in (({"status": 503, "body": error(503, "serviceNotAvailable", "Busy.")}, "provider_unavailable"),
                               ("html502", "provider_unavailable"), ("lost", "network"), ("oversized", "protocol"),
                               ("gzip", "protocol")):
            api.script["send"] = [scripted]
            before = len(api.calls)
            outcome = failure(lambda: run(logic.configuration({**config, "max_response_bytes": 65536}), send), code)
            assert outcome.uncertain and api.kinds()[before:].count("send") == 1, scripted
        api.script["mail_list"] = ["oversized"]
        assert not failure(lambda: run(logic.configuration({**config, "max_response_bytes": 65536, "max_retries": 0}), "go"),
                           "protocol").uncertain
        api.script["mail_list"] = [{"status": 400, "body": error(400, "BadRequest", "Invalid filter clause.")}]
        assert failure(lambda: run(config, "go"), "bad_request").provider_message == "Invalid filter clause."
        api.script["send"] = ["hold"]
        started = time.monotonic()
        cancelled = failure(lambda: run(config, send, cancel=lambda: time.monotonic() - started > 0.4), "cancelled")
        assert cancelled.uncertain and time.monotonic() - started < 2.5
        api.release.set()
        time.sleep(0.2)
        api.release.clear()
        api.script["download"] = ["hold"]
        started = time.monotonic()
        quiet = failure(lambda: run(config, {"operation": "file_download", "item_id": "01FIXTUREQ1REPORTPDF000000000000"},
                                    storage_dir=lambda: node, cancel=lambda: time.monotonic() - started > 0.4), "cancelled")
        assert not quiet.uncertain and time.monotonic() - started < 2.5
        assert not list(node.rglob("*.*")) and not list((node / "downloads").iterdir()), "No partial download remains"
        api.release.set()
        time.sleep(0.2)
        api.release.clear()
        api.script["send"] = ["hold"]
        started = time.monotonic()
        timed = failure(lambda: run(logic.configuration({**config, "timeout_sec": 5}), send), "timeout")
        assert timed.uncertain and 4 <= time.monotonic() - started < 8
        api.release.set()
        time.sleep(0.2)
        assert not [thread for thread in threading.enumerate() if thread.name == transport.THREAD_NAME]
        assert not failure(lambda: run(logic.configuration({**config, "api_base_url": "http://127.0.0.1:9", "max_retries": 0}),
                                       "go"), "network").uncertain
        failure(lambda: run(config, "go", cancel=lambda: True), "cancelled")
    with graph_api(tls=True) as api:
        config = config_of(api)
        assert run(config, {"limit": 3})["result"]["count"] == 3
        assert run(config, {"limit": 1}, GRANT)["result"]["requests"] == 2
        before = len(api.calls)
        untrusted = failure(lambda: run(logic.configuration({**config, "ca_file": ""}), "go"), "tls")
        assert not untrusted.uncertain and len(api.calls) == before
    print("[ok] Microsoft 365 errors, retries, uncertainty, cancellation, deadline and TLS", flush=True)


def value(data, port):
    raw = data.get("output_values", {}).get("m365:" + str(port), {}).get("value")
    return (json.loads(raw) if port != 3 else raw) if raw else None


def wallet(server, secret):
    http_json(server.base_url, "/api/application/secrets/init", method="POST", payload={"password": "disposable-m365-fixture"})
    http_json(server.base_url, "/api/application/secrets", method="POST",
              payload={"name": "microsoft_365_fixture", "value": json.dumps(secret)})


def graph_runs():
    """# - FB2 - # - FB4 - # - FB6 - # - FB8 - real graphs in One Shot and Active Runtime for bundled, managed and linked packages."""
    block = Microsoft365Block()
    for mode in ("centralized", "zeromq_active"):
        for origin in (None, "managed", "linked"):
            full = origin is None
            with graph_api() as api, isolated_server() as server:
                uploads = server.root_dir / "fixture-outgoing"
                uploads.mkdir()
                (uploads / "graph.txt").write_text("graph upload " + mode, encoding="utf-8")
                (server.root_dir / "outside.txt").write_text("outside the upload root", encoding="utf-8")
                node = block.build_node_payload(node_id="m365", position={"x": 380, "y": 160}, config_overrides=settings(
                    api, allow_writes=True, upload_root=str(uploads), drive_folder="Documents", timeout_sec=20,
                    watch_mail=True, emit_existing=True, existing_since_hours=1,
                    allowed_operations=["mail_list", "mail_watch", "file_download", "file_upload", "folder_create"]))
                if origin:
                    model = install_test_package(server, "microsoft_365", origin=origin)
                    node["block_version"] = model["version"]
                    for surface in ("modal", "inspector_panel", "node_card"):
                        html = surface_payload(server, model, node, surface=surface)["html"]
                        assert "{{" not in html and ACCESS not in html
                        if surface != "node_card":
                            assert 'data-block-config-field="allowed_recipients"' in html
                wallet(server, GRANT if full else TOKEN)
                sinks = [display_node("listed", "Items", 760, 20), display_node("receipt", "Result", 760, 160),
                         display_node("saved", "File", 760, 300), display_node("problem", "Problem", 760, 440)]
                edges = [data_edge("o1", "m365", 1, "listed", 1), data_edge("o2", "m365", 2, "receipt", 1),
                         data_edge("o3", "m365", 3, "saved", 1), data_edge("o4", "m365", 4, "problem", 1)]

                def graph(request="", file=""):
                    return graph_payload("Microsoft 365", [text_node("req", "Request", request, 60, 100),
                                                           text_node("src", "File", file, 60, 320), node, *sinks],
                                         [data_edge("in", "req", 1, "m365", 1), data_edge("up", "src", 1, "m365", 2), *edges])

                recent = [f"Report {index:02d}" for index in range(1, 6)]
                if mode == "centralized":
                    def once(document, expect="success", project=None):
                        project = project or create_project_api(server, document=document)["project"]
                        created = create_run_api(server, document, project_id=project["graph_id"], runtime_mode=mode)
                        data = wait_for_run_terminal(server, created["run_id"], timeout_sec=40)
                        assert data["status"] == expect, (origin, data.get("status"), data.get("logs"))
                        clean(data)
                        return data, project

                    data, _ = once(graph(json.dumps({"request_id": "q-1", "limit": 2})))
                    assert [item["subject"] for item in value(data, 1)] == ["Report 01", "Report 02"]
                    assert value(data, 2)["request_id"] == "q-1" and value(data, 2)["truncated"] is True
                    data, project = once(graph(json.dumps({"operation": "file_download", "path": "Reports/Q1 report.pdf"})))
                    path = Path(value(data, 3))
                    storage_dir = graph_storage_dir(server, project) / "instances" / "1" / "blocs" / "m365" / "downloads"
                    assert path.is_relative_to(storage_dir) and path.read_bytes() == PDF_BYTES, (path, storage_dir)
                    # One Shot never starts the watcher: mail_watch polls once per run, without replay across runs.
                    document = graph(json.dumps({"operation": "mail_watch"}))
                    data, project = once(document)
                    assert sorted(item["subject"] for item in value(data, 1)) == recent and value(data, 2)["started"]
                    api.deliver("Between runs")
                    data, _ = once(document, project=project)
                    assert [item["subject"] for item in value(data, 1)] == ["Between runs"] and not value(data, 2)["started"]
                    assert "delta" in api.kinds() and api.kinds().count("delta") == 2
                    if full:
                        before = len(api.items)
                        data, _ = once(graph(json.dumps({"operation": "folder_create", "name": "Graph"}), str(uploads / "graph.txt")))
                        assert api.kinds().count("simple_upload") == 1 and api.kinds().count("folder_create") == 1, api.kinds()
                        assert len(api.items) == before + 2
                        before = len(api.calls)
                        data, _ = once(graph(json.dumps({"operation": "file_upload", "path": "../outside.txt"})))
                        assert value(data, 4)["code"] == "path_denied" and len(api.calls) == before
                        assert data["results"]["m365"]["microsoft_365"]["error_codes"] == ["path_denied"]
                        http_json(server.base_url, "/api/application/secrets/lock", method="POST", payload={})
                        data, _ = once(graph("go"))
                        assert value(data, 4)["code"] == "secret" and len(api.calls) == before
                        http_json(server.base_url, "/api/application/secrets/unlock", method="POST",
                                  payload={"password": "disposable-m365-fixture"})
                        # Stop while OneDrive holds a download: the run ends promptly and nothing is left behind.
                        api.script["download"] = ["hold"]
                        document = graph(json.dumps({"operation": "file_download", "path": "Reports/Q1 report.pdf"}))
                        project = create_project_api(server, document=document)["project"]
                        rid = create_run_api(server, document, project_id=project["graph_id"], runtime_mode=mode)["run_id"]
                        until(lambda: api.holding == 1, 20, "held download not reached")
                        held, started = len(api.calls), time.monotonic()
                        stop_run_api(server, rid)
                        stopped = wait_for_run_terminal(server, rid, timeout_sec=20)
                        assert time.monotonic() - started < 10, stopped.get("logs")
                        assert stopped["status"] == "cancelled", (stopped["status"], value(stopped, 4), stopped.get("logs"))
                        assert value(stopped, 3) is None and len(api.calls) == held
                        assert "shutdown_timeout" not in json.dumps(stopped.get("logs"))
                        api.release.set()
                        storage_dir = graph_storage_dir(server, project) / "instances" / "1" / "blocs" / "m365"
                        until(lambda: not list(storage_dir.rglob(".partial-*")), 5, "a stopped download left a partial file")
                else:
                    document = graph()
                    project = create_project_api(server, document=document)["project"]
                    storage_dir = graph_storage_dir(server, project) / "instances" / "1" / "blocs" / "m365" / "downloads"

                    def publish(rid, source, text):
                        http_json(server.base_url, f"/api/runs/{rid}/active/control", method="POST", payload={
                            "action": "publish_output", "node_id": source, "port_id": 1, "value": text,
                            "content_type": "text/plain"})

                    def wait(rid, predicate, label):
                        data = wait_for_run_predicate(server, rid, lambda item: item.get("status") == "failed" or predicate(item),
                                                      label, timeout_sec=30)
                        assert data.get("status") != "failed", data.get("logs")
                        clean(data)
                        return data

                    def stop(rid):
                        started = time.monotonic()
                        stop_run_api(server, rid)
                        stopped = wait_for_run_terminal(server, rid, timeout_sec=20)
                        assert time.monotonic() - started < 10, stopped.get("logs")
                        assert "shutdown_timeout" not in json.dumps(stopped.get("logs"))
                        return stopped

                    def watched(item):
                        result = value(item, 2) or {}
                        return result.get("source") == "watch" and result.get("operation") == "mail_watch"

                    rid = prepare_release_run(server, project["graph_id"], document)["run_id"]
                    http_json(server.base_url, f"/api/runs/{rid}/play", method="POST", payload={})
                    try:
                        data = wait(rid, watched, "watched mail missing")
                        assert sorted(item["subject"] for item in value(data, 1)) == recent
                        publish(rid, "req", json.dumps({"request_id": "live-1", "limit": 3}))
                        data = wait(rid, lambda item: (value(item, 2) or {}).get("request_id") == "live-1", "listing missing")
                        assert [item["subject"] for item in value(data, 1)] == ["Report 01", "Report 02", "Report 03"]
                        publish(rid, "src", str(uploads / "graph.txt"))
                        data = wait(rid, lambda item: ((value(item, 2) or {}).get("item") or {}).get("name") == "graph.txt",
                                    "upload missing")
                        assert api.kinds().count("simple_upload") == 1
                        if full:
                            publish(rid, "req", json.dumps({"request_id": "live-2", "operation": "file_download",
                                                            "path": "Reports/Q1 report.pdf"}))
                            data = wait(rid, lambda item: (value(item, 2) or {}).get("request_id") == "live-2"
                                        and value(item, 3) is not None, "download missing")
                            assert Path(value(data, 3)).is_relative_to(storage_dir)
                            publish(rid, "req", json.dumps({"request_id": "live-3", "operation": "mail_watch"}))
                            data = wait(rid, lambda item: (value(item, 4) or {}).get("request_id") == "live-3", "busy missing")
                            assert value(data, 4)["code"] == "busy", "The listener owns the cursor while the Run is loaded"
                    finally:
                        stop(rid)
                    if full:
                        api.deliver("Between runs")
                        rid = prepare_release_run(server, project["graph_id"], document)["run_id"]
                        http_json(server.base_url, f"/api/runs/{rid}/play", method="POST", payload={})
                        try:
                            data = wait(rid, watched, "restart lost new mail")
                            assert [item["subject"] for item in value(data, 1)] == ["Between runs"], "A restart never replays"
                            api.script["mail_list"] = ["hold"]
                            publish(rid, "req", json.dumps({"request_id": "live-4"}))
                            until(lambda: api.holding == 1, 20, "held listing not reached")
                        finally:
                            stop(rid)
                            api.release.set()
                logs = "\n".join(server.logs)
                assert not any(secret in logs for secret in SECRETS)
            print("[ok] Microsoft 365 graph " + mode + " " + str(origin), flush=True)


def main():
    mail_protocol()
    watch_protocol()
    listener_protocol()
    calendar_protocol()
    drive_protocol()
    error_protocol()
    graph_runs()


if __name__ == "__main__":
    main()
