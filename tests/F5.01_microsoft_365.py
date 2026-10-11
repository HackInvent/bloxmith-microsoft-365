"""FB1-FB9: pure settings, credentials, requests, resources, errors, watch state, local files and block hooks.

No socket is opened and no provider is contacted; the wallet resolver is a local double.
"""

from datetime import datetime, timezone
import json
import os
from pathlib import Path
import sys
import tempfile
from types import MappingProxyType, SimpleNamespace

ROOT = Path(__file__).resolve().parents[3]
sys.path[:0] = [str(ROOT), str(ROOT / "tests"), str(Path(__file__).resolve().parent)]
from blocs.microsoft_365 import files, logic, storage
from blocs.microsoft_365.block import Microsoft365Block
from bloxsmith_app.block_api import BlockRuntimeContext, BlockRuntimePreparationContext
from graph_fixture import ACCESS, CHANNEL, FRIEND, GRANT, PARTNER, PUBLIC_GRANT, REF, STRANGER, TEAM, TENANT

NOW = datetime(2026, 10, 11, 8, 0, 0, tzinfo=timezone.utc)
MESSAGE_ID = "AAMkAGI2TGVzdE1lc3NhZ2UwMDAxAAA="
EVENT_ID = "AAMkAGI2TGVzdEV2ZW50MDAwMDAxAAA="
ITEM_ID = "01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K"
BASE = {"enabled": True, "secret_ref": REF, "tenant_id": TENANT}
WRITER = {**BASE, "allow_writes": True, "allow_send": True, "allow_delete": True,
          "allowed_recipients": [FRIEND, "@partner.example"], "allowed_channels": [TEAM + "/" + CHANNEL],
          "attachment_root": "/srv/attachments", "upload_root": "/srv/outgoing", "drive_folder": "Shared/Reports",
          "allowed_operations": list(logic.OPERATIONS)}


def rejected(call, code=None):
    """A contract violation must raise a coded error, never fall back silently."""
    try:
        call()
    except logic.GraphError as exc:
        assert code is None or exc.code == code, (code, exc.code, str(exc))
        assert ACCESS not in str(exc) and GRANT["client_secret"] not in str(exc) and GRANT["refresh_token"] not in str(exc)
        return exc
    raise AssertionError("Invalid value accepted")


def parse(value, config=None):
    return logic.parse_request(value, logic.configuration(config or WRITER), NOW)


def settings_contract():
    """# - FB1 - remote access, writes, sends and deletions stay off by default and every setting is bounded."""
    defaults = logic.configuration({})
    assert defaults["enabled"] is False and defaults["allow_writes"] is False and defaults["allow_send"] is False
    assert defaults["allow_delete"] is False and defaults["watch_mail"] is False and defaults["emit_existing"] is False
    assert defaults["api_base_url"] == logic.OFFICIAL_API == "https://graph.microsoft.com"
    assert defaults["login_base_url"] == "https://login.microsoftonline.com" and defaults["tenant_id"] == "common"
    assert logic.selectable_operations(defaults) == ["mail_list", "mail_get", "events_list", "files_list", "file_download"]
    assert defaults["watch_interval_sec"] == 300 and logic.BOUNDS["watch_interval_sec"][0] == 60
    assert set(Microsoft365Block().default_config()) == set(logic.DEFAULTS)
    config = logic.configuration({**WRITER, "execution": {"inhibited": False}, "tenant_id": " Fixture.Example ",
                                  "mail_folder": "Inbox", "drive_folder": "/Shared/Reports/",
                                  "allowed_recipients": [" Colleague@Fixture.Example", "@Partner.Example", FRIEND]})
    assert config["tenant_id"] == "fixture.example" and config["mail_folder"] == "inbox"
    assert config["drive_folder"] == "Shared/Reports" and logic.selectable_operations(config) == list(logic.OPERATIONS)
    assert config["allowed_recipients"] == ["colleague@fixture.example", "@partner.example"]
    frozen = logic.configuration(MappingProxyType({**WRITER, "allowed_operations": ("files_list",)}))
    assert frozen["allowed_operations"] == ["files_list"]
    for tenant in ("common", "organizations", "consumers", "72f988bf-86f1-41af-91ab-2d7cd011db47", "contoso.onmicrosoft.com"):
        assert logic.configuration({**BASE, "tenant_id": tenant})["tenant_id"] == tenant
    for raw in ({"enabled": True}, {**BASE, "secret_ref": ""}, {**BASE, "secret_ref": ACCESS},
                {**BASE, "secret_ref": "env://GRAPH"}, {**BASE, "token": ACCESS}, {**BASE, "client_secret": "x"},
                {**BASE, "tenant_id": "contoso"}, {**BASE, "tenant_id": "https://login.microsoftonline.com/x"},
                {**BASE, "tenant_id": 5}, {**BASE, "operation": "mail_delete"},
                {**BASE, "operation": "mail_send"}, {**BASE, "allowed_operations": ["event_create"]},
                {**BASE, "allowed_operations": []}, {**BASE, "allowed_operations": ["mail_list", "mail_list"]},
                {**BASE, "allowed_operations": "mail_list"},
                {**WRITER, "allow_send": False}, {**WRITER, "allow_delete": False},
                {**WRITER, "allowed_recipients": []}, {**WRITER, "allowed_channels": []}, {**WRITER, "upload_root": ""},
                {**WRITER, "allowed_recipients": ["not-an-address"]}, {**WRITER, "allowed_recipients": ["@nodot"]},
                {**WRITER, "allowed_recipients": "a@b.example"}, {**WRITER, "allowed_channels": ["team/19:x@thread"]},
                {**WRITER, "allowed_channels": [TEAM]}, {**WRITER, "upload_root": "relative/dir"},
                {**WRITER, "attachment_root": "/srv/../etc"}, {**BASE, "mail_folder": "Inbox2"},
                {**BASE, "mail_folder": "../inbox"}, {**BASE, "drive_folder": "a/../b"}, {**BASE, "drive_folder": "a:b"},
                {**BASE, "calendar_id": "short"}, {**BASE, "enabled": "yes"}, {**BASE, "allow_writes": 1},
                {**BASE, "body_format": "rtf"}, {**BASE, "conflict_behavior": "overwrite"},
                {**BASE, "conflict_behavior": "replace"}, {**BASE, "folder_exists": "rename"},
                {**BASE, "upload_chunk_bytes": 327681}, {**BASE, "time_zone": "Europe/Paris\""},
                {**BASE, "oauth_scope": ""}, {**BASE, "oauth_scope": "Mail.Read\nFiles.Read"}, {**BASE, "ca_file": "ca.pem"},
                {**BASE, "api_base_url": "http://graph.microsoft.com"}, {**BASE, "api_base_url": "https://graph.microsoft.com/v1.0"},
                {**BASE, "api_base_url": "https://user:pw@graph.microsoft.com"}, {**BASE, "api_base_url": "https://x.test/?a=1"},
                {**BASE, "api_base_url": "ftp://graph.microsoft.com"}, {**BASE, "login_base_url": "http://example.test:8080"}):
        rejected(lambda raw=raw: logic.configuration(raw), "config")
    # Disabled nodes may keep an incomplete write setup: nothing can run while remote access is off.
    assert logic.configuration({**WRITER, "enabled": False, "upload_root": "", "allowed_recipients": []})["upload_root"] == ""
    assert logic.configuration({**WRITER, "conflict_behavior": "replace"})["conflict_behavior"] == "replace"
    for key, (low, high) in logic.BOUNDS.items():
        if key == "upload_chunk_bytes":
            continue
        for bad in (low - 1, high + 1, True, float(low), str(low)):
            rejected(lambda key=key, bad=bad: logic.configuration({**BASE, key: bad}), "config")
        assert logic.configuration({**BASE, key: low})[key] == low and logic.configuration({**BASE, key: high})[key] == high
    assert logic.configuration({**BASE, "upload_chunk_bytes": 327680 * 192})["upload_chunk_bytes"] == 62914560
    origins = {"https://graph.microsoft.com/": "https://graph.microsoft.com", "http://127.0.0.1:8081": "http://127.0.0.1:8081",
               "http://localhost:9": "http://localhost:9", "http://[::1]:7": "http://[::1]:7",
               "https://Graph.Example.test:8443": "https://graph.example.test:8443"}
    for raw, expected in origins.items():
        assert logic.configuration({**BASE, "api_base_url": raw, "login_base_url": raw})["api_base_url"] == expected


def credential_contract():
    """# - FB1 - # - FB7 - wallet values: bare token, access_token JSON, public or confidential refresh grant; never echoed."""
    assert logic.credentials(" " + ACCESS + "\n") == {"access_token": ACCESS}
    assert logic.credentials({"access_token": ACCESS}) == {"access_token": ACCESS}
    assert logic.credentials(json.dumps(GRANT)) == GRANT and logic.credentials(PUBLIC_GRANT) == PUBLIC_GRANT
    for bad in ("", None, 12, {"access_token": ""}, {"access_token": "a b"}, {**GRANT, "extra": "x"},
                {"client_id": "a", "client_secret": "b"}, {"refresh_token": "b"}, '{"access_token": ',
                json.dumps({"access_token": 5}), {"tenant_id": TENANT, **GRANT}):
        error = rejected(lambda bad=bad: logic.credentials(bad), "secret")
        assert GRANT["refresh_token"] not in json.dumps(error.envelope())
    text = logic.redact("Bearer " + ACCESS + " secret=" + GRANT["client_secret"], [ACCESS, GRANT["client_secret"], ""])
    assert ACCESS not in text and GRANT["client_secret"] not in text and text.count("***") == 2


def mail_request_contract():
    """# - FB2 - # - FB6 - mail requests: folders, filters or search, select, page links, recipients and attachments."""
    reader = logic.configuration(BASE)
    trigger = logic.parse_request("go", reader, NOW)
    assert trigger == {"operation": "mail_list", "request_id": None, "limit": 100, "page_link": None, "folder": "inbox",
                       "filter": None, "search": None, "order": "receivedDateTime desc", "select": logic.MESSAGE_SELECT}
    unread = logic.parse_request({"unread_only": True, "received_after": "2026-10-01T10:00:00+02:00", "limit": 5,
                                  "folder": "Archive", "request_id": "r-1"}, reader, NOW)
    assert unread["filter"] == "receivedDateTime ge 2026-10-01T08:00:00Z and isRead eq false"
    assert unread["order"] == "receivedDateTime desc" and unread["folder"] == "archive" and unread["limit"] == 5
    only_unread = logic.parse_request({"unread_only": True}, reader, NOW)
    assert only_unread["filter"] == "receivedDateTime ge 1900-01-01T00:00:00Z and isRead eq false"
    raw = logic.parse_request({"filter": "from/emailAddress/address eq 'a@b.example'"}, reader, NOW)
    assert raw["filter"] == "from/emailAddress/address eq 'a@b.example'" and raw["order"] is None
    found = logic.parse_request({"search": " budget 2026 ", "select": ["subject", "from"]}, reader, NOW)
    assert found["search"] == "budget 2026" and found["order"] is None and found["select"] == "id,subject,from"
    link = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=10&$skip=10"
    assert logic.parse_request({"page_link": link, "limit": 10}, reader, NOW)["page_link"] == link
    got = logic.parse_request({"operation": "mail_get", "message_id": MESSAGE_ID}, reader, NOW)
    assert got == {"operation": "mail_get", "request_id": None, "message_id": MESSAGE_ID, "body_format": "text"}
    writer = logic.configuration(WRITER)
    sent = logic.parse_request({"operation": "mail_send", "to": [FRIEND, {"address": "Ops@Partner.Example", "name": "Ops"}],
                                "bcc": ["audit@partner.example"], "subject": "Report", "body": "Line 1\nLine 2",
                                "body_format": "html", "attachments": ["report.pdf"], "importance": "high"}, writer, NOW)
    assert sent["to"] == [{"address": FRIEND, "name": None}, {"address": "Ops@partner.example", "name": "Ops"}]
    assert sent["bcc"] == [{"address": "audit@partner.example", "name": None}] and sent["cc"] == []
    assert sent["attachments"] == ["report.pdf"] and sent["save_to_sent_items"] is True and sent["importance"] == "high"
    moved = logic.parse_request({"operation": "mail_move", "message_id": MESSAGE_ID, "destination": "DeletedItems"}, writer, NOW)
    assert moved["destination"] == "deleteditems"
    assert logic.parse_request({"operation": "mail_mark_read", "message_id": MESSAGE_ID, "is_read": False}, writer, NOW)[
        "is_read"] is False
    assert logic.parse_request({"operation": "mail_watch", "request_id": "w"}, writer, NOW) == {
        "operation": "mail_watch", "request_id": "w"}
    for config, value, code in (
            (reader, {"operation": "mail_send", "to": [FRIEND]}, "operation_not_allowed"),
            (reader, {"operation": "mail_purge"}, "invalid_input"), (reader, [{"folder": "inbox"}], "invalid_input"),
            (reader, "[1, 2]", "invalid_input"), (reader, {"limit": 101}, "invalid_input"), (reader, {"limit": 0}, "invalid_input"),
            (reader, {"folder": "Inbox2"}, "invalid_input"), (reader, {"filter": "a", "search": "b"}, "invalid_input"),
            (reader, {"filter": "a", "unread_only": True}, "invalid_input"), (reader, {"search": 'say "hi"'}, "invalid_input"),
            (reader, {"search": "a\\b"}, "invalid_input"), (reader, {"filter": "x" * 1025}, "invalid_input"),
            (reader, {"filter": "a\nb"}, "invalid_input"), (reader, {"select": ["subject", "body"]}, "invalid_input"),
            (reader, {"unread_only": "yes"}, "invalid_input"), (reader, {"received_after": "yesterday"}, "invalid_input"),
            (reader, {"received_after": "2026-02-30"}, "invalid_input"),
            (reader, {"page_link": "https://evil.example/v1.0/me/mailFolders/inbox/messages"}, "invalid_input"),
            (reader, {"page_link": "https://graph.microsoft.com/v1.0/me/drive/root/children"}, "invalid_input"),
            (reader, {"page_link": link, "folder": "inbox"}, "invalid_input"),
            (reader, {"operation": "mail_get"}, "invalid_input"),
            (reader, {"operation": "mail_get", "message_id": "short"}, "invalid_input"),
            (reader, {"operation": "mail_get", "message_id": MESSAGE_ID, "body_format": "rtf"}, "invalid_input"),
            (reader, '{"limit": 1, "limit": 2}', "invalid_input"), (reader, "{broken", "invalid_input"),
            (reader, {"request_id": "bad id"}, "invalid_input"), (reader, {"unknown": 1}, "invalid_input"),
            (writer, {"operation": "mail_send", "to": []}, "invalid_input"),
            (writer, {"operation": "mail_send", "to": [STRANGER]}, "recipient_not_allowed"),
            (writer, {"operation": "mail_send", "to": [FRIEND], "cc": ["x@sub.partner.example"]}, "recipient_not_allowed"),
            (writer, {"operation": "mail_send", "to": [FRIEND], "bcc": [STRANGER]}, "recipient_not_allowed"),
            (writer, {"operation": "mail_send", "to": ["Name <a@partner.example>"]}, "invalid_input"),
            (writer, {"operation": "mail_send", "to": [{"address": FRIEND, "type": "cc"}]}, "invalid_input"),
            (writer, {"operation": "mail_send", "to": [FRIEND], "attachments": ["a"] * 11}, "invalid_input"),
            (writer, {"operation": "mail_send", "to": [FRIEND], "subject": "a\nb"}, "invalid_input"),
            (writer, {"operation": "mail_send", "to": [FRIEND], "importance": "urgent"}, "invalid_input"),
            (logic.configuration({**WRITER, "attachment_root": ""}),
             {"operation": "mail_send", "to": [FRIEND], "attachments": ["a.txt"]}, "attachment_root_required"),
            (writer, {"operation": "mail_move", "message_id": MESSAGE_ID}, "invalid_input"),
            (writer, {"operation": "mail_mark_read", "message_id": MESSAGE_ID, "is_read": "true"}, "invalid_input"),
            (writer, {"operation": "mail_watch", "folder": "archive"}, "invalid_input"),
            (logic.configuration({**BASE, "max_request_bytes": 1024}), {"search": "x", "pad": "y" * 2000}, "request_too_large"),
            (logic.configuration({**BASE, "max_request_bytes": 1024}), "y" * 2000, "request_too_large")):
        rejected(lambda config=config, value=value: logic.parse_request(value, config, NOW), code)
    # A tampered configuration (validation bypassed) still cannot change or send anything.
    rejected(lambda: logic.parse_request({"operation": "mail_mark_read", "message_id": MESSAGE_ID},
                                         {**writer, "allow_writes": False}, NOW), "writes_disabled")
    rejected(lambda: logic.parse_request({"operation": "mail_send", "to": [FRIEND]}, {**writer, "allow_send": False}, NOW),
             "send_disabled")


def calendar_drive_teams_contract():
    """# - FB3 - # - FB4 - # - FB5 - event windows and times, attendee gates, OneDrive paths and allowlisted channels."""
    reader = logic.configuration(BASE)
    writer = logic.configuration(WRITER)
    window = logic.parse_request({"operation": "events_list"}, reader, NOW)
    assert window == {"operation": "events_list", "request_id": None, "limit": 100, "page_link": None, "calendar_id": "",
                      "start": "2026-10-11T08:00:00Z", "end": "2026-10-18T08:00:00Z"}
    explicit = logic.parse_request({"operation": "events_list", "start": "2026-10-12", "end": "2026-10-13T12:00:00+02:00"},
                                   reader, NOW)
    assert (explicit["start"], explicit["end"]) == ("2026-10-12T00:00:00Z", "2026-10-13T10:00:00Z")
    created = logic.parse_request({"operation": "event_create", "subject": "Review", "start": "2026-10-20T09:00",
                                   "end": "2026-10-20T10:30", "time_zone": "Europe/Paris", "location": "Room 1",
                                   "attendees": [FRIEND, {"address": "ops@partner.example", "type": "optional"}],
                                   "show_as": "tentative", "is_online_meeting": True}, writer, NOW)
    fields = created["fields"]
    assert fields["start"] == {"dateTime": "2026-10-20T09:00:00", "timeZone": "Europe/Paris"}
    assert fields["end"] == {"dateTime": "2026-10-20T10:30:00", "timeZone": "Europe/Paris"}
    assert fields["attendees"] == [{"emailAddress": {"address": FRIEND}, "type": "required"},
                                   {"emailAddress": {"address": "ops@partner.example"}, "type": "optional"}]
    assert fields["location"] == {"displayName": "Room 1"} and fields["showAs"] == "tentative" and fields["isOnlineMeeting"]
    zoned = logic.parse_request({"operation": "event_create", "subject": "UTC", "start": "2026-10-20T09:00:00+02:00",
                                 "end": "2026-10-20T08:00:00Z"}, writer, NOW)["fields"]
    assert zoned["start"] == {"dateTime": "2026-10-20T07:00:00", "timeZone": "UTC"} and zoned["end"]["dateTime"] == "2026-10-20T08:00:00"
    day = logic.parse_request({"operation": "event_create", "subject": "Off", "all_day": True, "start": "2026-10-20",
                               "end": "2026-10-21"}, writer, NOW)["fields"]
    assert day["isAllDay"] is True and day["start"] == {"dateTime": "2026-10-20T00:00:00", "timeZone": "UTC"}
    update = logic.parse_request({"operation": "event_update", "event_id": EVENT_ID, "subject": "Renamed"}, writer, NOW)
    assert update["fields"] == {"subject": "Renamed"} and "attendees" not in update
    assert logic.parse_request({"operation": "event_delete", "event_id": EVENT_ID}, writer, NOW)["event_id"] == EVENT_ID
    listed = logic.parse_request({"operation": "files_list"}, writer, NOW)
    assert listed["path"] == "Shared/Reports" and listed["item_id"] is None
    assert logic.parse_request({"operation": "files_list", "path": "Q1/Drafts"}, writer, NOW)["path"] == "Shared/Reports/Q1/Drafts"
    assert logic.parse_request({"operation": "file_download", "item_id": ITEM_ID}, writer, NOW)["item_id"] == ITEM_ID
    upload = logic.parse_request({"operation": "file_upload", "path": "a.pdf", "folder": "Q1", "name": "Final.pdf"}, writer, NOW)
    assert upload == {"operation": "file_upload", "request_id": None, "path": "a.pdf", "folder": "Shared/Reports/Q1",
                      "name": "Final.pdf"}
    assert logic.parse_request({"operation": "folder_create", "name": "2026"}, writer, NOW)["folder"] == "Shared/Reports"
    post = logic.parse_request({"operation": "channel_message_send", "content": "<b>Done</b>", "content_type": "html"}, writer, NOW)
    assert (post["team_id"], post["channel_id"], post["content_type"]) == (TEAM, CHANNEL, "html")
    assert logic.file_request(" /srv/outgoing/a.txt\n", writer) == {"operation": "file_upload", "request_id": None,
                                                                    "path": "/srv/outgoing/a.txt", "folder": "Shared/Reports"}
    no_send = logic.configuration({**WRITER, "allow_send": False, "allowed_operations": ["event_create", "event_update"]})
    for config, value, code in (
            (reader, {"operation": "events_list", "start": "2026-10-12", "end": "2026-10-12"}, "invalid_input"),
            (reader, {"operation": "events_list", "start": "2026-01-01", "end": "2027-01-03"}, "invalid_input"),
            (reader, {"operation": "events_list", "start": "2026-10-12T09:00"}, "invalid_input"),
            (reader, {"operation": "events_list", "calendar_id": "x"}, "invalid_input"),
            (reader, {"operation": "event_create", "subject": "x"}, "operation_not_allowed"),
            (writer, {"operation": "event_create", "subject": "x", "start": "2026-10-20T09:00"}, "invalid_input"),
            (writer, {"operation": "event_create", "start": "2026-10-20T09:00", "end": "2026-10-20T10:00"}, "invalid_input"),
            (writer, {"operation": "event_create", "subject": "x", "start": "2026-10-20T10:00", "end": "2026-10-20T09:00"},
             "invalid_input"),
            (writer, {"operation": "event_create", "subject": "x", "start": "2026-10-20T09:00", "end": "2026-10-20T10:00Z"},
             "invalid_input"),
            (writer, {"operation": "event_create", "subject": "x", "start": "2026-10-20T25:00", "end": "2026-10-20T26:00"},
             "invalid_input"),
            (writer, {"operation": "event_create", "subject": "x", "all_day": True, "start": "2026-10-20T09:00",
                      "end": "2026-10-21"}, "invalid_input"),
            (writer, {"operation": "event_create", "subject": "x", "start": "2026-10-20T09:00", "end": "2026-10-20T10:00",
                      "time_zone": "Paris\""}, "invalid_input"),
            (writer, {"operation": "event_create", "subject": "x", "start": "2026-10-20T09:00", "end": "2026-10-20T10:00",
                      "attendees": [STRANGER]}, "recipient_not_allowed"),
            (writer, {"operation": "event_create", "subject": "x", "start": "2026-10-20T09:00", "end": "2026-10-20T10:00",
                      "show_as": "away"}, "invalid_input"),
            (no_send, {"operation": "event_create", "subject": "x", "start": "2026-10-20T09:00", "end": "2026-10-20T10:00",
                       "attendees": [FRIEND]}, "send_disabled"),
            (writer, {"operation": "event_update", "event_id": EVENT_ID}, "invalid_input"),
            (writer, {"operation": "event_update", "event_id": EVENT_ID, "body_format": "html"}, "invalid_input"),
            (writer, {"operation": "event_update", "event_id": EVENT_ID, "all_day": True, "start": "2026-10-20"}, "invalid_input"),
            (logic.configuration({**WRITER, "allow_delete": False, "allowed_operations": ["files_list"]}),
             {"operation": "event_delete", "event_id": EVENT_ID}, "operation_not_allowed"),
            (writer, {"operation": "files_list", "path": "../secret"}, "invalid_input"),
            (writer, {"operation": "files_list", "path": "a:b"}, "invalid_input"),
            (writer, {"operation": "files_list", "path": "a", "item_id": ITEM_ID}, "invalid_input"),
            (writer, {"operation": "files_list", "item_id": "bad id"}, "invalid_input"),
            (writer, {"operation": "file_download"}, "invalid_input"),
            (writer, {"operation": "file_upload"}, "invalid_input"),
            (writer, {"operation": "file_upload", "path": "a", "name": "x/y"}, "invalid_input"),
            (writer, {"operation": "file_upload", "path": "a", "name": " lead"}, "invalid_input"),
            (writer, {"operation": "file_upload", "path": "a", "folder": "../../escape"}, "invalid_input"),
            (writer, {"operation": "folder_create"}, "invalid_input"),
            (writer, {"operation": "folder_create", "name": "end."}, "invalid_input"),
            (writer, {"operation": "channel_message_send", "team_id": TEAM, "channel_id": "19:other@thread.tacv2",
                      "content": "x"}, "channel_not_allowed"),
            (writer, {"operation": "channel_message_send", "team_id": "not-a-guid", "channel_id": CHANNEL, "content": "x"},
             "invalid_input"),
            (writer, {"operation": "channel_message_send", "content": ""}, "invalid_input"),
            (writer, {"operation": "channel_message_send", "content": "x" * 28001}, "invalid_input"),
            (writer, {"operation": "channel_message_send", "content": "x", "content_type": "markdown"}, "invalid_input")):
        rejected(lambda config=config, value=value: logic.parse_request(value, config, NOW), code)
    rejected(lambda: logic.file_request("a.txt", reader), "operation_not_allowed")
    rejected(lambda: logic.file_request("a\x00b", writer), "invalid_input")
    rejected(lambda: logic.file_request("a.txt", {**writer, "allow_writes": False}), "writes_disabled")
    rejected(lambda: logic.parse_request({"operation": "event_delete", "event_id": EVENT_ID},
                                         {**writer, "allow_delete": False}, NOW), "delete_disabled")


def resource_contract():
    """# - FB2 - # - FB3 - # - FB4 - # - FB7 - Graph resources, paths and links are validated and normalized."""
    raw = {"id": MESSAGE_ID, "subject": "Q1", "from": {"emailAddress": {"name": "A", "address": FRIEND}},
           "toRecipients": [{"emailAddress": {"address": PARTNER}}], "receivedDateTime": "2026-10-11T07:00:00Z",
           "isRead": False, "hasAttachments": True, "flag": {"flagStatus": "flagged"}, "categories": ["Blue"]}
    message = logic.normalize_message(raw)
    assert message["from"] == {"name": "A", "address": FRIEND} and message["to"] == [{"name": None, "address": PARTNER}]
    assert message["flag"] == "flagged" and message["cc"] == [] and message["web_link"] is None and "body" not in message
    full = logic.normalize_message({**raw, "body": {"contentType": "HTML", "content": "x" * 300},
                                    "attachments": [{"@odata.type": "#microsoft.graph.fileAttachment", "id": "AAMkAttachment000001=",
                                                     "name": "a.pdf", "contentType": "application/pdf", "size": 10,
                                                     "isInline": False}]}, max_body=256)
    assert full["body"] == {"content_type": "html", "content": "x" * 256, "truncated": True, "length": 300}
    assert full["attachments"][0]["kind"] == "file" and full["bcc"] == []
    for bad in (None, [], {**raw, "id": ""}, {**raw, "id": "a b"}, {**raw, "subject": 5}, {**raw, "isRead": "no"},
                {**raw, "from": {"name": "x"}}, {**raw, "toRecipients": "x"}, {**raw, "categories": [1]}):
        rejected(lambda bad=bad: logic.normalize_message(bad), "protocol")
    event = logic.normalize_event({"id": EVENT_ID, "subject": "S", "start": {"dateTime": "2026-10-20T09:00:00.0000000",
                                                                            "timeZone": "UTC"},
                                   "attendees": [{"emailAddress": {"address": FRIEND}, "type": "required",
                                                  "status": {"response": "accepted"}}], "isOrganizer": True,
                                   "onlineMeeting": {"joinUrl": "https://teams.example.test/join"}})
    assert event["attendees"] == [{"name": None, "address": FRIEND, "type": "required", "response": "accepted"}]
    assert event["online_meeting_url"] == "https://teams.example.test/join" and event["end"] is None
    rejected(lambda: logic.normalize_event({"id": EVENT_ID, "attendees": {}}), "protocol")
    item = logic.normalize_item({"id": ITEM_ID, "name": "a.pdf", "size": 10, "file": {"mimeType": "application/pdf",
                                 "hashes": {"sha256Hash": "AB" * 32, "quickXorHash": "q="}},
                                 "parentReference": {"id": "01PARENT", "path": "/drive/root:/Docs"},
                                 "@microsoft.graph.downloadUrl": "https://download.example.test/secret-link"})
    assert item["sha256"] == "ab" * 32 and item["parent_path"] == "/drive/root:/Docs" and not item["is_folder"]
    assert "secret-link" not in json.dumps(item), "Pre-authenticated download URLs never leave the block"
    for bad in ({"id": ITEM_ID, "size": -1}, {"id": ITEM_ID, "file": "x"}, {"id": ITEM_ID, "file": {"hashes": {"sha256Hash": "x"}}},
                {"id": ITEM_ID, "folder": {"childCount": "2"}}):
        rejected(lambda bad=bad: logic.normalize_item(bad), "protocol")
    assert logic.collection({"value": [1], "@odata.nextLink": "n"}, 2) == ([1], "n", None)
    for bad in ({"value": {}}, {"value": [1, 2, 3]}, {"value": [], "@odata.deltaLink": 5}, []):
        rejected(lambda bad=bad: logic.collection(bad, 2), "protocol")
    assert logic.drive_item_path("", None, "/children") == "/v1.0/me/drive/root/children"
    assert logic.drive_item_path("Docs/Q1 #1", None, "/children") == "/v1.0/me/drive/root:/Docs/Q1%20%231:/children"
    assert logic.drive_item_path("Docs/a.pdf", None) == "/v1.0/me/drive/root:/Docs/a.pdf"
    assert logic.drive_item_path(None, "A!1", "/content") == "/v1.0/me/drive/items/A%211/content"
    assert logic.folder_path("AAMk+/=") == "/v1.0/me/mailFolders/AAMk%2B%2F%3D"
    config = logic.configuration(BASE)
    assert logic.api_link("https://graph.microsoft.com/v1.0/me/x?$skiptoken=1", config) == "/v1.0/me/x?$skiptoken=1"
    for bad in ("https://graph.microsoft.com.evil.example/v1.0/me", "http://graph.microsoft.com/v1.0/me",
                "https://graph.microsoft.com/beta/me", "https://user@graph.microsoft.com/v1.0/me", "/v1.0/me", 5,
                "https://graph.microsoft.com/v1.0/me#frag"):
        assert logic.api_link(bad, config) is None, bad
    for link, family in (("https://graph.microsoft.com/v1.0/me/mailfolders('AQMkADNk')/messages?$skip=10", "mail_list"),
                         ("https://graph.microsoft.com/v1.0/" + "users" + "/0f1e2d3c/mailFolders/inbox/messages?$skip=1", "mail_list"),
                         ("https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=x&$skip=10", "events_list"),
                         ("https://graph.microsoft.com/v1.0/me/calendars('AAMk')/calendarView?$skip=2", "events_list"),
                         ("https://graph.microsoft.com/v1.0/drives/b!abc/items/01ABC/children?$skiptoken=x", "files_list"),
                         ("https://graph.microsoft.com/v1.0/me/drive/root/children?$skiptoken=x", "files_list")):
        assert logic.api_link(link, config, logic.LINK_PATTERNS[family]) is not None, link
        others = [name for name in logic.LINK_PATTERNS if name != family]
        assert all(logic.api_link(link, config, logic.LINK_PATTERNS[name]) is None for name in others), link
    assert logic.api_link("https://graph.microsoft.com/v1.0/me/messages?$skip=1", config, logic.LINK_PATTERNS["mail_list"]) is None
    assert logic.api_link("https://graph.microsoft.com/v1.0/me/mailfolders('x')/messages/delta?$deltatoken=t", config,
                          logic.WATCH_LINK) is not None
    assert logic.api_link("https://graph.microsoft.com/v1.0/me/mailfolders('x')/messages?$skip=1", config, logic.WATCH_LINK) is None
    assert logic.transfer_link("https://contoso.sharepoint.example/_layouts/download?x=1", config) == (
        "https://contoso.sharepoint.example", "/_layouts/download?x=1")
    for bad in ("http://contoso.sharepoint.example/x", "https://203.0.113.5/x", "https://127.0.0.1/x", "https://a:b@h.example/x",
                "file:///etc/passwd", None):
        assert logic.transfer_link(bad, config) is None, bad
    loop = logic.configuration({**BASE, "api_base_url": "http://127.0.0.1:9000"})
    assert logic.transfer_link("http://127.0.0.1:9000/download/x", loop) == ("http://127.0.0.1:9000", "/download/x")
    assert logic.transfer_link("http://127.0.0.1:9001/download/x", loop) is None
    assert logic.local_name("Notes: week/1", ".md") == "Notes_ week_1.md"
    assert logic.local_name("../../etc/passwd") == "_.._etc_passwd" and logic.local_name("..") == "download"
    long = logic.local_name("é" * 300 + ".txt")
    assert long.endswith(".txt") and len(long.encode("utf-8")) <= logic.LOCAL_NAME_BYTES
    assert logic.guessed_type("a.pdf") == "application/pdf" and logic.guessed_type("noext") == "application/octet-stream"


def provider_contract():
    """# - FB7 - Graph error envelopes map to stable codes; changes after 5xx are uncertain; nothing leaks."""
    secrets = [ACCESS]
    for status, provider, code in ((400, "BadRequest", "bad_request"), (401, "InvalidAuthenticationToken", "unauthorized"),
                                   (403, "ErrorAccessDenied", "forbidden"), (404, "ErrorItemNotFound", "not_found"),
                                   (409, "nameAlreadyExists", "conflict"), (410, "SyncStateNotFound", "sync_reset"),
                                   (400, "syncStateNotFound", "sync_reset"), (412, "preconditionFailed", "provider_error"),
                                   (507, "insufficientStorage", "storage_quota"), (403, "ErrorQuotaExceeded", "storage_quota"),
                                   (429, "TooManyRequests", "rate_limited"), (500, "generalException", "provider_unavailable"),
                                   (503, "serviceNotAvailable", "provider_unavailable")):
        body = {"error": {"code": provider, "message": "Detail with " + ACCESS + "\n" + "x" * 400,
                          "innerError": {"request-id": "rid-1", "date": "2026-10-11T08:00:00"}}}
        for mutation in (False, True):
            error = rejected(lambda status=status, body=body, mutation=mutation: logic.provider_failure(
                status, body, secrets, mutation=mutation, retry_after=7), code)
            envelope = error.envelope(operation="mail_send", request_id=None)
            assert envelope["http_status"] == status and envelope["provider_code"] == provider
            assert envelope["provider_request_id"] == "rid-1" and "request_id" not in envelope
            assert len(envelope["provider_message"]) <= 200 and "\n" not in envelope["provider_message"]
            assert ACCESS not in json.dumps(envelope)
            assert error.uncertain is (mutation and code == "provider_unavailable")
            assert envelope.get("retry_after") == (7 if code in ("rate_limited", "provider_unavailable") else None)
    rejected(lambda: logic.provider_failure(502, None, secrets, mutation=False), "provider_unavailable")
    assert rejected(lambda: logic.provider_failure(404, {"error": "x"}, secrets, mutation=False), "not_found").provider_code is None


def watch_state_contract():
    """# - FB6 - the watch document holds links, times and digests only; it is validated, bounded and locked."""
    config = logic.configuration({**BASE, "watch_mail": True})
    key = logic.watch_key(config)
    assert len(key) == 24 and REF not in key and logic.watch_key({**config, "mail_folder": "archive"}) != key
    digest = logic.seen_digest(MESSAGE_ID)
    assert len(digest) == 16 and MESSAGE_ID not in digest
    entry = {"link": "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta?$deltatoken=x",
             "link_kind": "delta", "since": "2026-10-11T08:00:00Z", "last_received": None, "seen": [digest], "updated_at": 1}
    document = logic.with_watch(logic.empty_state(), key, entry)
    assert logic.state_document(document) is document
    for index in range(10):
        document = logic.with_watch(document, f"{index:024x}", {**entry, "updated_at": 10 + index})
    assert len(document["watches"]) == 8 and key not in document["watches"], "The oldest cursors are dropped first"
    for bad in ({"schema_version": 2, "watches": {}}, {"schema_version": 1, "watches": {"bad": entry}},
                {"schema_version": 1, "watches": {key: {**entry, "link_kind": "other"}}},
                {"schema_version": 1, "watches": {key: {**entry, "link": None}}},
                {"schema_version": 1, "watches": {key: {**entry, "seen": ["xyz"]}}},
                {"schema_version": 1, "watches": {key: {**entry, "since": "yesterday"}}},
                {"schema_version": 1, "watches": {key: {**entry, "token": ACCESS}}}, []):
        rejected(lambda bad=bad: logic.state_document(bad), "state_invalid")
    with tempfile.TemporaryDirectory() as folder:
        directory = Path(folder)
        store = storage.WatchStore(directory)
        assert store.load() == logic.empty_state()
        assert store.acquire()
        other = storage.WatchStore(directory)
        assert other.acquire() is False, "A second watcher of the same node storage is refused"
        store.save(logic.with_watch(logic.empty_state(), key, entry))
        assert store.load()["watches"][key]["seen"] == [digest]
        text = (directory / storage.STATE_NAME).read_text()
        assert REF not in text and ACCESS not in text and oct(os.stat(directory / storage.STATE_NAME).st_mode & 0o777) == "0o600"
        store.close()
        assert other.acquire()
        other.close()
        (directory / storage.STATE_NAME).write_text("{broken")
        rejected(lambda: store.load(), "state_invalid")
        rejected(lambda: store.save({"schema_version": 1, "watches": {"x": {}}}), "state_invalid")


def local_files_contract():
    """# - FB2 - # - FB4 - sources stay inside their roots and are bounded; downloads are reserved, atomic and pruned."""
    with tempfile.TemporaryDirectory() as folder:
        base = Path(folder).resolve()
        root, outside = base / "outgoing", base / "private"
        root.mkdir()
        outside.mkdir()
        (root / "report.txt").write_bytes(b"quarterly")
        (root / "sub").mkdir()
        (root / "sub" / "deep.csv").write_bytes(b"a,b\n")
        (outside / "secret.txt").write_bytes(b"do not upload")
        (root / "escape.txt").symlink_to(outside / "secret.txt")
        (root / "big.bin").write_bytes(b"x" * 2048)
        (root / "half.bin").write_bytes(b"h" * 1500000)
        (root / "half2.bin").write_bytes(b"h" * 1500000)
        (root / "third.bin").write_bytes(b"h" * 200000)
        os.mkfifo(root / "pipe")
        config = logic.configuration({**WRITER, "upload_root": str(root), "attachment_root": str(root),
                                      "max_upload_bytes": 1024})
        local = files.open_upload("report.txt", config)
        try:
            assert local.read_all() == b"quarterly" and local.name == "report.txt" and local.size == 9
            assert local.read(2, 3) == b"art"
        finally:
            local.close()
        for value, code in (("../private/secret.txt", "path_denied"), (str(outside / "secret.txt"), "path_denied"),
                            ("escape.txt", "path_denied"), ("missing.txt", "path"), ("sub", "path"), ("pipe", "path"),
                            ("big.bin", "file_too_large")):
            error = rejected(lambda value=value: files.open_upload(value, config), code)
            assert str(outside) not in str(error) and str(root) not in str(error)
        rejected(lambda: files.open_upload("report.txt", {**config, "upload_root": ""}), "upload_root_required")
        rejected(lambda: files.read_attachments(["report.txt"], {**config, "attachment_root": ""}), "attachment_root_required")
        assert files.read_attachments(["report.txt", "sub/deep.csv"], config) == [("report.txt", b"quarterly"),
                                                                                  ("deep.csv", b"a,b\n")]
        roomy = {**config, "max_attachment_bytes": logic.SEND_ATTACHMENTS_LIMIT}
        assert [name for name, _ in files.read_attachments(["half.bin", "third.bin"], roomy)] == ["half.bin", "third.bin"]
        rejected(lambda: files.read_attachments(["half.bin", "half2.bin", "third.bin"], roomy), "file_too_large")
        rejected(lambda: files.read_attachments(["big.bin"], config | {"max_attachment_bytes": 1024}), "file_too_large")
        changing = files.open_upload("report.txt", config)
        try:
            (root / "report.txt").write_bytes(b"quarterly!")
            rejected(lambda: changing.check(), "path")
        finally:
            changing.close()
        storage_dir = base / "node"
        storage_dir.mkdir()
        rejected(lambda: files.storage_directory(None), "storage")
        rejected(lambda: files.storage_directory(lambda: (_ for _ in ()).throw(RuntimeError("block_storage.missing_scope"))),
                 "storage")
        assert files.storage_directory(lambda: str(storage_dir)) == storage_dir
        kept = []
        for index in range(3):
            target = files.reserve(storage_dir)
            os.write(target["fd"], b"content %d" % index)
            kept.append(files.finish(target, f"file-{index}.txt", 2))
        assert [path.read_bytes() for path in kept[1:]] == [b"content 1", b"content 2"] and not kept[0].exists()
        failed = files.reserve(storage_dir)
        os.write(failed["fd"], b"partial")
        files.discard(failed)
        assert not failed["folder"].exists() and not list(storage_dir.rglob(".partial-*"))
        (storage_dir / "downloads" / "user-notes").mkdir()
        files.finish(files.reserve(storage_dir), "x.txt", 1)
        assert (storage_dir / "downloads" / "user-notes").is_dir(), "Folders the block did not create are never pruned"


def block_contract():
    """# - FB1 - # - FB6 - # - FB7 - # - FB8 - # - FB9 - hooks fail before side effects; events run once; surfaces are translated."""
    block = Microsoft365Block()
    ports_in = tuple(SimpleNamespace(id=port["id"], name=port["name"]) for port in block.model["ports"]["inputs"])
    ports_out = tuple(SimpleNamespace(id=port["id"], name=port["name"]) for port in block.model["ports"]["outputs"])
    assert [port.name for port in ports_in] == ["request", "file"]
    assert [port.name for port in ports_out] == ["items", "result", "file", "error"]
    assert all(port["execution_requirement"] == "not_required_for_execution" for port in block.model["ports"]["inputs"])
    assert block.model["runtime"]["centralized_dependency_mode"] == "any_event" and block.model["version"] == "0.1.0"
    assert block.model["description"] == "Work with Outlook mail and calendar, OneDrive files and Teams messages."
    calls = []

    def resolver(ref):
        calls.append(ref)
        return "not json {" + ACCESS

    def run(config, inputs=None, events=(), services=None, mode="zeromq_active"):
        context = BlockRuntimeContext(config={**block.default_config(), **config}, inputs=inputs or {}, input_ports=ports_in,
                                      output_ports=ports_out, input_events=tuple(events), runtime_mode=mode,
                                      services={"resolve_secret": resolver, **(services or {})})
        result = block.execute_runtime(context)
        assert result.status in ("success", "skipped"), result
        return result, [(output.port_name, json.loads(output.value)) for output in result.outputs]

    assert run(BASE)[0].status == "skipped"
    with tempfile.TemporaryDirectory() as folder:
        root = Path(folder).resolve() / "outgoing"
        root.mkdir()
        (root / "a.txt").write_text("a")
        (root.parent / "private.txt").write_text("private")
        (root / "sub").mkdir()
        (root / "sub" / "bad:name.txt").write_text("OneDrive names cannot contain a colon")
        writer = {**WRITER, "upload_root": str(root), "attachment_root": str(root)}
        for config, inputs, code in (({**BASE, "enabled": False}, {"request": "go"}, "disabled"),
                                     (BASE, {"request": {"operation": "mail_send", "to": [FRIEND]}}, "operation_not_allowed"),
                                     (BASE, {"request": '{"limit":'}, "invalid_input"),
                                     (BASE, {"file": str(root / "a.txt")}, "operation_not_allowed"),
                                     (writer, {"file": str(root / "missing.txt")}, "path"),
                                     (writer, {"file": str(root.parent / "private.txt")}, "path_denied"),
                                     (writer, {"file": "sub/bad:name.txt"}, "invalid_input"),
                                     (writer, {"request": {"operation": "mail_send", "to": [FRIEND],
                                                           "attachments": ["../private.txt"]}}, "path_denied"),
                                     (writer, {"request": {"operation": "mail_send", "to": [STRANGER]}}, "recipient_not_allowed"),
                                     (BASE, {"request": {"operation": "file_download", "item_id": ITEM_ID}}, "storage"),
                                     (writer, {"request": {"operation": "mail_watch"}}, "storage"),
                                     ({**BASE, "secret_ref": ACCESS}, {"request": "go"}, "config")):
            result, outputs = run(config, inputs)
            assert [name for name, _ in outputs] == ["error"] and outputs[0][1]["code"] == code, outputs
            assert result.metadata["microsoft_365"] == {"status": "error", "operations": 1, "error_codes": [code]}
            assert ACCESS not in json.dumps(outputs) and ACCESS not in " ".join(result.logs)
        assert "local file name" in run(writer, {"file": "sub/bad:name.txt"})[1][0][1]["message"]
        assert not calls, "The wallet must not be read before settings, request and local file validation"
        result, outputs = run(BASE, {"request": {"request_id": "req-1", "limit": 3}})
        error = outputs[0][1]
        assert error["code"] == "secret" and error["request_id"] == "req-1" and error["operation"] == "mail_list" and calls == [REF]
        assert error["input"] == "request" and ACCESS not in json.dumps(outputs)
        assert result.logs == ["[microsoft-365-error] mail_list: secret"]
        result, outputs = run(BASE, {"request": "go", "file": "x"})
        assert outputs[0][1]["code"] == "invalid_input" and "ambiguous" in outputs[0][1]["message"]
        memory = {}
        services = {"get_node_runtime_value": lambda key, default=None: memory.get(key, default),
                    "set_node_runtime_value": lambda key, value: memory.__setitem__(key, value)}
        event = SimpleNamespace(edge_id="e1", input_port_id=1, input_port_name="request", source_node_id="src",
                                source_port_id=1, value='{"request_id": "once"}', content_type="application/json", sequence=0)
        calls.clear()
        first = run(BASE, {"request": event.value}, [event], services, "centralized")[1]
        again = run(BASE, {"request": event.value}, [event], services, "centralized")
        assert first[0][1]["request_id"] == "once" and again[0].status == "skipped" and calls == [REF]
        other = SimpleNamespace(**{**vars(event), "edge_id": "e2", "input_port_id": 2, "input_port_name": "file",
                                   "value": str(root / "a.txt")})
        mixed = run(writer, {"request": event.value, "file": other.value}, [event, other], services, "centralized")[1]
        assert [item["input"] for _name, item in mixed] == ["file"], "Only the new file event runs"
        flood = [SimpleNamespace(**{**vars(event), "edge_id": f"x{index}"}) for index in range(9)]
        assert run(BASE, {}, flood)[1][0][1]["code"] == "invalid_input"

    def prepared(config, mode="zeromq_active"):
        context = BlockRuntimePreparationContext.from_context(SimpleNamespace(config=config, runtime_mode=mode))
        return block.prepare_runtime(context)

    assert prepared(WRITER).listen_on_run is False
    assert prepared({**BASE, "watch_mail": True}).listen_on_run is True
    assert prepared({**BASE, "watch_mail": True}, "centralized").listen_on_run is False, "One Shot never starts the watcher"
    assert prepared({**BASE, "enabled": False, "watch_mail": True}).listen_on_run is False
    rejected(lambda: prepared({**BASE, "secret_ref": ACCESS}), "config")
    node = block.build_node_payload(node_id="m365", position={"x": 0, "y": 0})
    accepted = block.handle_ui_action(node=node, action="save_settings", values={"node_patch": {"config": WRITER}})
    operation = accepted["graph_operations"][-1]
    assert operation["op"] == "update_node_config" and operation["replace"] is True
    assert operation["config"]["allowed_recipients"] == WRITER["allowed_recipients"]
    refused = block.handle_ui_action(node=node, action="save_settings",
                                     values={"node_patch": {"config": {**BASE, "operation": "mail_send"}}})
    assert "allow_writes" in refused["error"] and "graph_operations" not in refused
    html = block.render_modal(node=node)["html"]
    assert 'data-block-config-field="secret_ref"' in html and 'data-microsoft-365-show="write" hidden' in html
    assert 'data-microsoft-365-show="watch" hidden' in html and 'data-block-value-type="json"' in html
    assert "{{" not in html and "data-path-browser" in html and 'data-i18n="block.microsoft_365.label_attachment_root"' in html
    writing = {**node, "config": {**node["config"], **WRITER, "watch_mail": True}}
    panel = block.render_inspector_panel(node=writing)["html"]
    assert 'data-microsoft-365-show="write" hidden' not in panel and 'data-microsoft-365-show="watch" hidden' not in panel
    card = block.render_node_card(node=node)["html"]
    assert card.count("Remote access off") == 1 and 'data-i18n="block.microsoft_365.card_disabled" data-i18n-params=""' in card
    assert "Read only · mail_list" in block.render_node_card(node={**node, "config": {**node["config"], **BASE}})["html"]
    assert "Read and write · mail_list" in block.render_node_card(node={**node, "config": {**node["config"], **WRITER}})["html"]
    card = block.render_node_card(node=writing)["html"]
    assert ("Watching new mail · mail_list" in card and 'data-i18n="block.microsoft_365.card_watch" '
            'data-i18n-params="{&quot;operation&quot;: &quot;mail_list&quot;}"' in card)
    assert REF not in card and "/srv/outgoing" not in card and TENANT not in card
    assert block.translate("block.microsoft_365.card_write", {"operation": "mail_send"}, lang="fr") == "Lecture et écriture · mail_send"
    english = json.loads((block.directory / "locales" / "en.json").read_text(encoding="utf-8"))
    french = json.loads((block.directory / "locales" / "fr.json").read_text(encoding="utf-8"))
    assert set(english) == set(french) and all(key.startswith("block.microsoft_365.") for key in english)
    fields = json.loads((block.directory / "fields.json").read_text(encoding="utf-8"))
    assert {spec["key"] for spec in fields} == set(logic.DEFAULTS)
    assert all(english["block.microsoft_365.label_" + spec["key"]] == spec["label"] for spec in fields)
    assert {spec.get("group") for spec in fields} == {None, "write", "watch"}


def main():
    settings_contract()
    credential_contract()
    mail_request_contract()
    calendar_drive_teams_contract()
    resource_contract()
    provider_contract()
    watch_state_contract()
    local_files_contract()
    block_contract()
    print("[ok] Microsoft 365 settings, credentials, requests, resources, errors, watch state, local files and hooks")


if __name__ == "__main__":
    main()
