"""Loopback-only Microsoft Graph v1.0 and identity-platform double. No Microsoft account or credential is used.

It reproduces the documented request/response shapes used by the block: the
v2.0 refresh-token grant, messages (list with $filter/$search/$orderby/$top and
@odata.nextLink, get with $expand=attachments and the body-content-type
preference, move, PATCH isRead, attachments and $value, sendMail with 202),
messages/delta with $filter=receivedDateTime ge, Prefer odata.maxpagesize,
nextLink/deltaLink and @removed entries, calendarView, events create (with
transactionId), update and delete, OneDrive children, items with
@microsoft.graph.downloadUrl, simple PUT uploads, upload sessions with
Content-Range fragments, folder creation with conflictBehavior, Teams channel
messages, the {"error": {"code", "message", "innerError"}} envelope,
Retry-After on 429 and 410 sync resets. Every credential below is built at run
time and obviously fake.
"""

from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
import hashlib
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
from pathlib import Path
import re
import secrets
import shutil
import socket
import ssl
import subprocess
import tempfile
import threading
import time
from types import SimpleNamespace
from urllib.parse import parse_qs, unquote, urlsplit
from zoneinfo import ZoneInfo

ACCESS = "-".join(("fixture", "graph", "access", "token", "not", "real"))
REFRESHED = "-".join(("fixture", "graph", "refreshed", "token", "not", "real"))
ROTATED = "-".join(("fixture", "graph", "rotated", "refresh", "not", "real"))
GRANT = {"client_id": "0f1e2d3c-4b5a-4968-8776-655443322110",
         "client_secret": "-".join(("fixture", "client", "secret", "not", "real")),
         "refresh_token": "-".join(("fixture", "refresh", "token", "not", "real"))}
PUBLIC_GRANT = {"client_id": GRANT["client_id"], "refresh_token": GRANT["refresh_token"]}
REF = "secret://workspace/microsoft_365_fixture"
TENANT = "fixture.example"
ME = "owner@fixture.example"
FRIEND = "colleague@fixture.example"
PARTNER = "partner@partner.example"
STRANGER = "stranger@elsewhere.example"
TEAM = "6d1f5b2a-0c3e-4f7a-9b8c-1d2e3f4a5b6c"
CHANNEL = "19:fixturegeneral0001@thread.tacv2"
OTHER_CHANNEL = "19:fixturerandom0002@thread.tacv2"
FOLDERS = {"inbox": "AAMkFixtureFolderInbox0000=", "archive": "AAMkFixtureFolderArchive000=",
           "sentitems": "AAMkFixtureFolderSent00000=", "deleteditems": "AAMkFixtureFolderDeleted00=",
           "drafts": "AAMkFixtureFolderDrafts0000=", "junkemail": "AAMkFixtureFolderJunk00000=",
           "outbox": "AAMkFixtureFolderOutbox0000="}
CUSTOM_FOLDER = "AAMkFixtureFolderProjects0="
ROOT_ID = "01FIXTUREROOT0000000000000000000"
DOCUMENTS = "01FIXTUREDOCUMENTS00000000000000"
REPORTS = "01FIXTUREREPORTS0000000000000000"
BULK = "01FIXTUREBULK0000000000000000000"
PDF_ID = "01FIXTUREQ1REPORTPDF000000000000"
PHOTO_ID = "01FIXTUREPHOTOPNG000000000000000"
PDF_BYTES = b"%PDF-1.7\n% fixture quarterly report\n" + bytes(range(256)) * 8 + b"\n%%EOF\n"
PHOTO_BYTES = b"\x89PNG\r\n\x1a\nfixture-image-bytes"
ATTACHMENT_BYTES = b"invoice line 1\ninvoice line 2\n"
SEED_MESSAGES = 12


def settings(api, **extra):
    """Block settings aimed at this fixture only."""
    return {"enabled": True, "secret_ref": REF, "tenant_id": TENANT, "api_base_url": api.base_url,
            "login_base_url": api.base_url, "ca_file": api.ca_file, "timeout_sec": 15, **extra}


def error(status, code, message):
    return {"error": {"code": code, "message": message,
                      "innerError": {"date": "2026-10-11T08:00:00", "request-id": "fixture-request-" + str(status),
                                     "client-request-id": "fixture-client-request"}}}


def stamp(moment):
    return moment.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def parse(text):
    return datetime.fromisoformat(text.replace("Z", "+00:00"))


class Message:
    def __init__(self, ident, folder, subject, sender, received, *, body="", is_read=False, attachments=()):
        self.id, self.folder, self.subject, self.sender = ident, folder, subject, sender
        self.received, self.body, self.is_read = received, body, is_read
        self.attachments = list(attachments)

    def resource(self, select=None, body_type="text", expand=False):
        full = {"id": self.id, "conversationId": "AAQkFixtureConversation" + self.id[-6:],
                "internetMessageId": "<" + self.id[-8:] + "@fixture.example>", "subject": self.subject,
                "from": {"emailAddress": {"name": "Sender", "address": self.sender}},
                "sender": {"emailAddress": {"name": "Sender", "address": self.sender}},
                "toRecipients": [{"emailAddress": {"name": "Owner", "address": ME}}], "ccRecipients": [],
                "bccRecipients": [], "replyTo": [], "receivedDateTime": stamp(self.received),
                "sentDateTime": stamp(self.received - timedelta(seconds=5)), "isRead": self.is_read, "isDraft": False,
                "importance": "normal", "hasAttachments": bool(self.attachments), "bodyPreview": self.body[:255],
                "webLink": "https://outlook.example.test/owa/?ItemID=" + self.id, "parentFolderId": FOLDERS.get(self.folder, self.folder),
                "categories": [], "flag": {"flagStatus": "notFlagged"},
                "body": {"contentType": body_type, "content": ("<p>" + self.body + "</p>") if body_type == "html" else self.body}}
        result = {key: value for key, value in full.items() if select is None or key in select or key == "id"}
        if expand:
            result["attachments"] = [{"@odata.type": kind, "id": ident, "name": name, "contentType": ctype,
                                      "size": len(data) + 120, "isInline": False}
                                     for ident, kind, name, ctype, data in self.attachments]
        return result


class Event:
    def __init__(self, ident, subject, start, end, *, attendees=(), organizer=True, zone="UTC"):
        self.id, self.subject, self.start, self.end = ident, subject, start, end
        self.attendees, self.organizer, self.zone = list(attendees), organizer, zone
        self.extra = {}

    def resource(self, zone="UTC"):
        def shown(moment):
            target = moment if zone == "UTC" else moment.astimezone(ZoneInfo(zone))
            return {"dateTime": target.strftime("%Y-%m-%dT%H:%M:%S.0000000"), "timeZone": zone}
        return {"id": self.id, "subject": self.subject, "start": shown(self.start), "end": shown(self.end),
                "isAllDay": self.extra.get("isAllDay", False), "location": {"displayName": self.extra.get("location", "")},
                "organizer": {"emailAddress": {"name": "Owner" if self.organizer else "Partner",
                                               "address": ME if self.organizer else PARTNER}},
                "attendees": [{"emailAddress": {"address": address, "name": address.split("@")[0]}, "type": "required",
                               "status": {"response": "none", "time": "0001-01-01T00:00:00Z"}} for address in self.attendees],
                "isOnlineMeeting": self.extra.get("isOnlineMeeting", False), "onlineMeeting": None,
                "webLink": "https://outlook.example.test/calendar/item/" + self.id, "showAs": self.extra.get("showAs", "busy"),
                "isCancelled": False, "isOrganizer": self.organizer, "bodyPreview": "", "categories": [], "type": "singleInstance",
                "seriesMasterId": None}


class Item:
    def __init__(self, ident, name, parent, *, content=None, folder=False, sha256=True):
        self.id, self.name, self.parent, self.content, self.folder, self.sha256 = ident, name, parent, content, folder, sha256


def resource_path(api, item):
    parts, current = [], api.items.get(item.parent) if item.parent else None
    while current is not None and current.id != ROOT_ID:
        parts.append(current.name)
        current = api.items.get(current.parent)
    return "/drive/root:" + ("/" + "/".join(reversed(parts)) if parts else "")


def drive_resource(api, item, download=False):
    body = {"id": item.id, "name": item.name, "createdDateTime": "2026-01-15T09:00:00Z",
            "lastModifiedDateTime": "2026-04-02T12:30:00Z", "webUrl": "https://onedrive.example.test/" + item.id,
            "eTag": "\"{" + item.id[-8:] + "},1\"", "parentReference": {"driveId": "b!fixturedrive", "id": item.parent,
                                                                    "path": resource_path(api, item)} if item.parent else {}}
    if item.folder:
        body["folder"] = {"childCount": len([child for child in api.items.values() if child.parent == item.id])}
        body["size"] = 0
    else:
        hashes = {"quickXorHash": "fixtureQuickXorNotChecked="}
        if item.sha256:
            hashes["sha256Hash"] = hashlib.sha256(item.content).hexdigest().upper()
        body["file"] = {"mimeType": "application/pdf" if item.name.endswith(".pdf") else "application/octet-stream",
                        "hashes": hashes}
        body["size"] = len(item.content)
        if download:
            token = secrets.token_hex(8)
            api.downloads[token] = item.id
            body["@microsoft.graph.downloadUrl"] = api.base_url + "/download/" + token
    return body


def seed(now):
    messages = {}
    for index in range(1, SEED_MESSAGES + 1):
        ident = f"AAMkFixtureMessage{index:04d}AAAA="
        attachments = []
        if index == 1:
            attachments = [("AAMkFixtureAttachment01AA=", "#microsoft.graph.fileAttachment", "invoice.txt", "text/plain",
                            ATTACHMENT_BYTES),
                           ("AAMkFixtureAttachment02AA=", "#microsoft.graph.itemAttachment", "forwarded", "message/rfc822", b"")]
        messages[ident] = Message(ident, "inbox", f"Report {index:02d}", PARTNER if index % 3 else FRIEND,
                                  now - timedelta(minutes=10 * index, seconds=30), body=f"Body of report {index:02d}. Budget line {index}.",
                                  is_read=index % 2 == 0, attachments=attachments)
    messages["AAMkFixtureArchived01AAAA="] = Message("AAMkFixtureArchived01AAAA=", "archive", "Old archive", FRIEND,
                                                     now - timedelta(days=40), body="archived")
    base = now.replace(minute=0, second=0, microsecond=0) + timedelta(days=1)
    events = {}
    for ident, subject, offset, attendees, organizer in (
            ("AAMkFixtureEventSolo00001=", "Focus time", 0, (), True),
            ("AAMkFixtureEventTeam00002=", "Team sync", 2, (FRIEND,), True),
            ("AAMkFixtureEventPartner003=", "Partner review", 4, (STRANGER,), True),
            ("AAMkFixtureEventInvited004=", "Invited call", 6, (ME,), False)):
        events[ident] = Event(ident, subject, base + timedelta(hours=offset), base + timedelta(hours=offset + 1),
                              attendees=attendees, organizer=organizer)
    items = [Item(ROOT_ID, "root", None, folder=True), Item(DOCUMENTS, "Documents", ROOT_ID, folder=True),
             Item(REPORTS, "Reports", DOCUMENTS, folder=True), Item(BULK, "Bulk", ROOT_ID, folder=True),
             Item(PDF_ID, "Q1 report.pdf", REPORTS, content=PDF_BYTES, sha256=True),
             Item(PHOTO_ID, "photo.png", ROOT_ID, content=PHOTO_BYTES, sha256=False)]
    for index in range(1, 26):
        items.append(Item(f"01FIXTUREBULKITEM{index:015d}", f"note-{index:02d}.txt", BULK, content=f"bulk {index}\n".encode()))
    return messages, events, {item.id: item for item in items}


FILTER_CLAUSE = re.compile(r"(receivedDateTime) (ge|gt) ([0-9T:.Z-]+)|(isRead) eq (true|false)")


def parse_filter(text):
    clauses = []
    for part in text.split(" and "):
        match = FILTER_CLAUSE.fullmatch(part.strip())
        if not match:
            raise ValueError(part)
        if match.group(1):
            clauses.append(("received", match.group(2), parse(match.group(3))))
        else:
            clauses.append(("isRead", "eq", match.group(5) == "true"))
    return clauses


def keep(message, clauses):
    for field, operator, value in clauses:
        if field == "received" and not (message.received >= value if operator == "ge" else message.received > value):
            return False
        if field == "isRead" and message.is_read is not value:
            return False
    return True


@contextmanager
def graph_api(tls=False):
    """Run the double on a dynamic loopback port; every owned thread is joined on exit."""
    with tempfile.TemporaryDirectory() as folder:
        directory = Path(folder)
        lock = threading.Lock()
        now = datetime.now(timezone.utc).replace(microsecond=0)
        messages, events, items = seed(now)
        api = SimpleNamespace(calls=[], script={}, release=threading.Event(), closing=threading.Event(), holding=0,
                              tokens={ACCESS, REFRESHED}, messages=messages, events=events, items=items, now=now,
                              changes=[], seq=0, delta_states={}, pages={}, downloads={}, sessions={}, sent=[],
                              notifications=[], deleted_events=[], channel_messages=[], transactions={}, counter=0,
                              channels={(TEAM, CHANNEL), (TEAM, OTHER_CHANNEL)})

        def record_change(message_id, folder_name, removed=False):
            api.seq += 1
            api.changes.append((api.seq, message_id, folder_name, removed))

        def deliver(subject, *, folder_name="inbox", sender=PARTNER, received=None, body="new mail"):
            with lock:
                api.counter += 1
                ident = f"AAMkFixtureDelivered{api.counter:04d}AA="
                api.messages[ident] = Message(ident, folder_name, subject, sender,
                                              received or datetime.now(timezone.utc).replace(microsecond=0), body=body)
                record_change(ident, folder_name)
                return ident

        def touch(ident):
            with lock:
                record_change(ident, api.messages[ident].folder)

        api.deliver, api.touch = deliver, touch

        class Handler(BaseHTTPRequestHandler):
            protocol_version = "HTTP/1.1"

            def log_message(self, *_args):
                pass

            def reply(self, status, body, *, raw=None, content_type="application/json; odata.metadata=minimal",
                      headers=None, length=None):
                data = raw if raw is not None else (b"" if body is None else json.dumps(body).encode("utf-8"))
                try:
                    self.send_response(status)
                    if data or raw is not None:
                        self.send_header("Content-Type", content_type)
                    if length != "omit":
                        self.send_header("Content-Length", str(len(data) if length is None else length))
                    self.send_header("Connection", "close")
                    for key, value in (headers or {}).items():
                        self.send_header(key, value)
                    self.end_headers()
                    self.wfile.write(data)
                except (BrokenPipeError, ConnectionError, ssl.SSLError):
                    pass
                finally:
                    self.close_connection = True

            def do_GET(self):
                self.dispatch("GET")

            def do_POST(self):
                self.dispatch("POST")

            def do_PUT(self):
                self.dispatch("PUT")

            def do_PATCH(self):
                self.dispatch("PATCH")

            def do_DELETE(self):
                self.dispatch("DELETE")

            def scripted(self, kind):
                """Apply one scripted action; return True when the reply has been handled."""
                queue = api.script.get(kind)
                if not queue:
                    return False
                action = queue.pop(0)
                if action == "hold":
                    with lock:
                        api.holding += 1
                    try:
                        api.release.wait(20)
                    finally:
                        with lock:
                            api.holding -= 1
                    return api.closing.is_set()
                if action == "lost":
                    try:
                        self.connection.shutdown(socket.SHUT_RDWR)
                    except OSError:
                        pass
                    self.close_connection = True
                    return True
                if action == "html502":
                    self.reply(502, None, raw=b"<html>Bad Gateway</html>", content_type="text/html")
                    return True
                if action == "oversized":
                    self.reply(200, {"value": [{"id": "x" * 40, "subject": "y" * 2000}] * 60})
                    return True
                if action == "gzip":
                    self.reply(200, None, raw=b"\x1f\x8b", content_type="application/json", headers={"Content-Encoding": "gzip"})
                    return True
                if isinstance(action, dict) and "status" in action:
                    self.reply(action["status"], action.get("body"), headers=action.get("headers"))
                    return True
                return False

            def dispatch(self, method):
                length = int(self.headers.get("Content-Length", "0"))
                assert 0 <= length <= 64 * 1024 * 1024
                raw = self.rfile.read(length)
                parts = urlsplit(self.path)
                query = {key: values[0] for key, values in parse_qs(parts.query, keep_blank_values=True).items()}
                call = {"method": method, "path": parts.path, "raw_query": parts.query, "query": query, "at": time.monotonic(),
                        "auth": self.headers.get("Authorization"), "accept_encoding": self.headers.get("Accept-Encoding"),
                        "content_type": self.headers.get("Content-Type"), "prefer": self.headers.get("Prefer"),
                        "content_range": self.headers.get("Content-Range"), "kind": None, "body": None, "size": len(raw)}
                api.calls.append(call)
                token_path = re.fullmatch(r"/([^/]+)/oauth2/v2\.0/token", parts.path)
                if token_path and method == "POST":
                    call["kind"] = "token"
                    self.token(call, unquote(token_path.group(1)), raw)
                    return
                if parts.path.startswith("/download/"):
                    call["kind"] = "download"
                    if not self.scripted("download"):
                        self.download(call, parts.path[len("/download/"):])
                    return
                if parts.path.startswith("/upload/"):
                    call["kind"] = "chunk" if method == "PUT" else "session_delete"
                    if not self.scripted(call["kind"]):
                        with lock:
                            self.chunk(call, parts.path[len("/upload/"):], raw, method)
                    return
                route = self.route(method, parts.path)
                if route is None:
                    self.reply(404, error(404, "UnknownError", "Resource not found for the segment."))
                    return
                kind, groups = route
                call["kind"] = kind
                token = (self.headers.get("Authorization") or "").removeprefix("Bearer ")
                if token not in api.tokens:
                    self.reply(401, error(401, "InvalidAuthenticationToken", "Access token validation failure. Invalid audience."))
                    return
                if raw and (self.headers.get("Content-Type") or "").startswith("application/json"):
                    call["body"] = json.loads(raw)
                if self.scripted(kind):
                    return
                with lock:
                    getattr(self, "op_" + kind)(call, query, groups, raw)

            def route(self, method, path):
                table = (
                    ("GET", r"/v1\.0/me/mailFolders/([^/]+)/messages", "mail_list"),
                    ("GET", r"/v1\.0/me/mailFolders/([^/]+)/messages/delta", "delta"),
                    ("GET", r"/v1\.0/me/mailfolders\('([^']+)'\)/messages/delta", "delta"),
                    ("GET", r"/v1\.0/me/messages/([^/]+)", "mail_get"),
                    ("PATCH", r"/v1\.0/me/messages/([^/]+)", "mail_patch"),
                    ("POST", r"/v1\.0/me/messages/([^/]+)/move", "mail_move"),
                    ("GET", r"/v1\.0/me/messages/([^/]+)/attachments/([^/]+)", "attachment"),
                    ("GET", r"/v1\.0/me/messages/([^/]+)/attachments/([^/]+)/\$value", "attachment_value"),
                    ("POST", r"/v1\.0/me/sendMail", "send"),
                    ("GET", r"/v1\.0/me(?:/calendars/([^/]+))?/calendarView", "calendar_view"),
                    ("POST", r"/v1\.0/me(?:/calendars/([^/]+))?/events", "event_create"),
                    ("GET", r"/v1\.0/me/events/([^/]+)", "event_get"),
                    ("PATCH", r"/v1\.0/me/events/([^/]+)", "event_patch"),
                    ("DELETE", r"/v1\.0/me/events/([^/]+)", "event_delete"),
                    ("GET", r"/v1\.0/me/drive/(root|items/[^/:]+|root:/.+):?/children", "children"),
                    ("GET", r"/v1\.0/drives/b!fixturedrive/(items/[^/:]+)/children", "children"),
                    ("POST", r"/v1\.0/me/drive/(root|items/[^/:]+|root:/.+):?/children", "folder_create"),
                    ("PUT", r"/v1\.0/me/drive/root:/(.+):/content", "simple_upload"),
                    ("POST", r"/v1\.0/me/drive/root:/(.+):/createUploadSession", "create_session"),
                    ("GET", r"/v1\.0/me/drive/(root|items/[^/:]+|root:/.+)", "item"),
                    ("POST", r"/v1\.0/teams/([^/]+)/channels/([^/]+)/messages", "channel_message"),
                )
                for verb, pattern, kind in table:
                    match = re.fullmatch(pattern, path)
                    if verb == method and match:
                        return kind, [unquote(group) if group else group for group in match.groups()]
                return None

            # -- identity platform ------------------------------------------------------------
            def token(self, call, tenant, raw):
                form = {key: values[0] for key, values in parse_qs(raw.decode("ascii")).items()}
                call["tenant"] = tenant
                call["body"] = {key: ("<redacted>" if key in ("client_secret", "refresh_token") else value)
                                for key, value in form.items()}
                call["has_secret"] = "client_secret" in form
                if self.scripted("token"):
                    return
                if tenant != TENANT:
                    self.reply(400, {"error": "invalid_request", "error_description": "AADSTS90002: Tenant '" + tenant
                                     + "' not found. Trace ID: fixture", "error_codes": [90002]})
                    return
                expected = {"client_id": GRANT["client_id"], "grant_type": "refresh_token",
                            "refresh_token": GRANT["refresh_token"]}
                secret_ok = form.get("client_secret", GRANT["client_secret"]) == GRANT["client_secret"]
                if {key: form.get(key) for key in expected} != expected or not secret_ok or not form.get("scope"):
                    self.reply(400, {"error": "invalid_grant", "error_description": "AADSTS70000: The provided grant is "
                                     "invalid or malformed. Trace ID: fixture", "error_codes": [70000]})
                    return
                self.reply(200, {"token_type": "Bearer", "scope": form["scope"], "expires_in": 3599, "ext_expires_in": 3599,
                                 "access_token": REFRESHED, "refresh_token": ROTATED})

            # -- paging helpers ----------------------------------------------------------------
            def page(self, call, query, path, values, size, link_path=None):
                """Serve values with $skiptoken nextLinks that keep the other query parameters."""
                offset = 0
                path = link_path or path
                if "$skiptoken" in query:
                    state = api.pages.get(query["$skiptoken"])
                    if state is None or state[0] != call["path"]:
                        self.reply(400, error(400, "BadRequest", "Invalid skip token."))
                        return
                    offset = state[1]
                body = {"value": values[offset:offset + size]}
                if offset + size < len(values):
                    token = secrets.token_hex(6)
                    api.pages[token] = (path, offset + size)
                    others = "&".join(part for part in call["raw_query"].split("&") if part and not part.startswith("%24skiptoken")
                                      and not part.startswith("$skiptoken"))
                    body["@odata.nextLink"] = (api.base_url + path + "?" + (others + "&" if others else "")
                                               + "$skiptoken=" + token)
                self.reply(200, body)

            def top(self, query, default=10):
                value = query.get("$top", str(default))
                if not value.isdigit() or not 1 <= int(value) <= 1000:
                    raise ValueError("$top")
                return int(value)

            # -- mail -----------------------------------------------------------------------
            def folder_name(self, value):
                if value in FOLDERS:
                    return value
                reverse = {ident: name for name, ident in FOLDERS.items()}
                if value in reverse:
                    return reverse[value]
                return "projects" if value == CUSTOM_FOLDER else None

            def op_mail_list(self, call, query, groups, _raw):
                name = self.folder_name(groups[0])
                if name is None:
                    self.reply(404, error(404, "ErrorItemNotFound", "The specified object was not found in the store."))
                    return
                found = [message for message in api.messages.values() if message.folder == name]
                if "$search" in query:
                    if "$orderby" in query or "$filter" in query:
                        self.reply(400, error(400, "ErrorInvalidUrlQuery", "The query parameter '$orderBy' is invalid."))
                        return
                    text = query["$search"]
                    if not (text.startswith('"') and text.endswith('"')):
                        self.reply(400, error(400, "BadRequest", "Syntax error in $search."))
                        return
                    needle = text[1:-1].lower()
                    found = [m for m in found if needle in m.subject.lower() or needle in m.body.lower()]
                if "$filter" in query:
                    try:
                        clauses = parse_filter(query["$filter"])
                    except ValueError:
                        self.reply(400, error(400, "BadRequest", "Invalid filter clause."))
                        return
                    if "$orderby" in query and clauses[0][0] != "received":
                        self.reply(400, error(400, "InefficientFilter", "The restriction or sort order is too complex."))
                        return
                    found = [m for m in found if keep(m, clauses)]
                if query.get("$orderby", "receivedDateTime desc") != "receivedDateTime desc":
                    self.reply(400, error(400, "BadRequest", "Unsupported $orderby."))
                    return
                found.sort(key=lambda m: m.received, reverse=True)
                select = query.get("$select", "").split(",") if "$select" in query else None
                try:
                    size = self.top(query)
                except ValueError:
                    self.reply(400, error(400, "BadRequest", "Invalid $top."))
                    return
                self.page(call, query, call["path"], [m.resource(select) for m in found], size)

            def op_delta(self, call, query, groups, _raw):
                name = self.folder_name(groups[0])
                match = re.fullmatch(r"odata\.maxpagesize=([0-9]+)", call["prefer"] or "")
                size = int(match.group(1)) if match else 10
                select = query.get("$select", "").split(",") if "$select" in query else None
                if "$deltatoken" in query or "$skiptoken" in query:
                    token = query.get("$deltatoken") or query.get("$skiptoken")
                    state = api.delta_states.get(token)
                    if state is None or state["folder"] != name:
                        self.reply(410, error(410, "SyncStateNotFound", "The sync state is not valid or has expired."))
                        return
                    if "$deltatoken" in query:
                        # Changes since the round that issued this deltaLink, newest state per message.
                        latest = {}
                        for seq, ident, folder_name, removed in api.changes:
                            if seq > state["seq"]:
                                latest[ident] = (seq, folder_name, removed)
                        values = []
                        for ident, (seq, _folder, removed) in sorted(latest.items(), key=lambda item: item[1][0]):
                            message = api.messages.get(ident)
                            if removed or message is None or message.folder != name:
                                values.append({"id": ident, "@removed": {"reason": "deleted"}})
                            elif message.received >= state["since"]:
                                values.append(message.resource(select))
                        state = {"folder": name, "since": state["since"], "values": values, "offset": 0, "seq": api.seq}
                else:
                    if "$filter" not in query:
                        self.reply(400, error(400, "BadRequest", "This fixture expects the receivedDateTime filter."))
                        return
                    clauses = parse_filter(query["$filter"])
                    assert len(clauses) == 1 and clauses[0][0] == "received", clauses
                    since = clauses[0][2]
                    found = sorted((m for m in api.messages.values() if m.folder == name and m.received >= since),
                                   key=lambda m: m.received)
                    state = {"folder": name, "since": since, "values": [m.resource(select) for m in found], "offset": 0,
                             "seq": api.seq}
                page = state["values"][state["offset"]:state["offset"] + size]
                body = {"value": page}
                token = secrets.token_hex(8)
                # Like Graph, continuation links use the OData key syntax of the folder.
                base = api.base_url + "/v1.0/me/mailfolders('" + name + "')/messages/delta"
                if state["offset"] + size < len(state["values"]):
                    api.delta_states[token] = {**state, "offset": state["offset"] + size}
                    body["@odata.nextLink"] = base + "?$skiptoken=" + token
                else:
                    api.delta_states[token] = {"folder": name, "since": state["since"], "seq": state["seq"]}
                    body["@odata.deltaLink"] = base + "?$deltatoken=" + token
                self.reply(200, body)

            def message(self, ident):
                message = api.messages.get(ident)
                if message is None:
                    self.reply(404, error(404, "ErrorItemNotFound", "The specified object was not found in the store."))
                return message

            def op_mail_get(self, call, query, groups, _raw):
                message = self.message(groups[0])
                if message is None:
                    return
                match = re.fullmatch(r'outlook\.body-content-type="(text|html)"', call["prefer"] or "")
                expand = query.get("$expand", "").startswith("attachments")
                select = query["$select"].split(",") if "$select" in query else None
                self.reply(200, message.resource(select, match.group(1) if match else "html", expand))

            def op_mail_patch(self, call, query, groups, _raw):
                message = self.message(groups[0])
                if message is None:
                    return
                assert set(call["body"]) == {"isRead"}, call["body"]
                message.is_read = call["body"]["isRead"]
                record_change(message.id, message.folder)
                self.reply(200, message.resource())

            def op_mail_move(self, call, query, groups, _raw):
                message = self.message(groups[0])
                if message is None:
                    return
                target = self.folder_name(call["body"].get("destinationId", ""))
                if target is None:
                    self.reply(400, error(400, "ErrorInvalidIdMalformed", "Id is malformed."))
                    return
                old = message.id
                del api.messages[old]
                record_change(old, message.folder, removed=True)
                api.counter += 1
                message.id = f"AAMkFixtureMoved{api.counter:04d}AAAAA="
                message.folder = target
                api.messages[message.id] = message
                record_change(message.id, target)
                self.reply(201, message.resource())

            def attachment(self, groups):
                message = self.message(groups[0])
                if message is None:
                    return None
                for entry in message.attachments:
                    if entry[0] == groups[1]:
                        return entry
                self.reply(404, error(404, "ErrorItemNotFound", "The specified object was not found in the store."))
                return None

            def op_attachment(self, call, query, groups, _raw):
                entry = self.attachment(groups)
                if entry is not None:
                    ident, kind, name, ctype, data = entry
                    assert query.get("$select") == "id,name,contentType,size,isInline", query
                    self.reply(200, {"@odata.type": kind, "id": ident, "name": name, "contentType": ctype,
                                     "size": len(data) + 120, "isInline": False})

            def op_attachment_value(self, call, query, groups, _raw):
                entry = self.attachment(groups)
                if entry is not None:
                    self.reply(200, None, raw=entry[4], content_type=entry[3])

            def op_send(self, call, query, groups, raw):
                body = call["body"]
                assert set(body) == {"message", "saveToSentItems"}, body
                assert set(body["message"]) <= {"subject", "importance", "body", "toRecipients", "ccRecipients",
                                                "bccRecipients", "attachments"}, body["message"]
                api.sent.append(body)
                self.reply(202, None)

            # -- calendar -------------------------------------------------------------------
            def op_calendar_view(self, call, query, groups, _raw):
                if groups[0] not in (None, "AAMkFixtureCalendarWork0="):
                    self.reply(404, error(404, "ErrorItemNotFound", "The specified object was not found in the store."))
                    return
                if "$skiptoken" not in query and ("startDateTime" not in query or "endDateTime" not in query):
                    self.reply(400, error(400, "ErrorInvalidParameter", "This request requires a time window specified "
                                          "by the query string parameters StartDateTime and EndDateTime."))
                    return
                match = re.fullmatch(r'outlook\.timezone="([^"]+)"', call["prefer"] or "")
                zone = match.group(1) if match else "UTC"
                state = api.pages.get(query.get("$skiptoken"))
                start = parse(query["startDateTime"]) if "startDateTime" in query else state[2]
                end = parse(query["endDateTime"]) if "endDateTime" in query else state[3]
                found = sorted((event for event in api.events.values() if event.start < end and event.end > start),
                               key=lambda event: event.start)
                size = self.top(query) if "$top" in query else 10
                values = [event.resource(zone) for event in found]
                offset = state[1] if state else 0
                body = {"value": values[offset:offset + size]}
                if offset + size < len(values):
                    token = secrets.token_hex(6)
                    api.pages[token] = (call["path"], offset + size, start, end)
                    body["@odata.nextLink"] = api.base_url + call["path"] + f"?$top={size}&$skiptoken=" + token
                self.reply(200, body)

            def zoned(self, value):
                moment = datetime.fromisoformat(value["dateTime"])
                zone = timezone.utc if value["timeZone"] == "UTC" else ZoneInfo(value["timeZone"])
                return moment.replace(tzinfo=zone).astimezone(timezone.utc)

            def op_event_create(self, call, query, groups, _raw):
                body = call["body"]
                transaction = body.get("transactionId")
                assert isinstance(transaction, str) and len(transaction) == 36, body
                if transaction in api.transactions:
                    self.reply(201, api.events[api.transactions[transaction]].resource())
                    return
                api.counter += 1
                ident = f"AAMkFixtureEventNew{api.counter:04d}AA="
                attendees = [entry["emailAddress"]["address"] for entry in body.get("attendees", [])]
                event = Event(ident, body["subject"], self.zoned(body["start"]), self.zoned(body["end"]), attendees=attendees)
                event.extra = {key: body[key] for key in ("isAllDay", "showAs", "isOnlineMeeting") if key in body}
                if "location" in body:
                    event.extra["location"] = body["location"]["displayName"]
                api.events[ident] = event
                api.transactions[transaction] = ident
                api.notifications.extend(("invitation", address) for address in attendees)
                match = re.fullmatch(r'outlook\.timezone="([^"]+)"', call["prefer"] or "")
                self.reply(201, event.resource(match.group(1) if match else "UTC"))

            def event(self, ident):
                event = api.events.get(ident)
                if event is None:
                    self.reply(404, error(404, "ErrorItemNotFound", "The specified object was not found in the store."))
                return event

            def op_event_get(self, call, query, groups, _raw):
                event = self.event(groups[0])
                if event is not None:
                    self.reply(200, event.resource())

            def op_event_patch(self, call, query, groups, _raw):
                event = self.event(groups[0])
                if event is None:
                    return
                body = call["body"]
                if "subject" in body:
                    event.subject = body["subject"]
                if "start" in body:
                    event.start = self.zoned(body["start"])
                if "end" in body:
                    event.end = self.zoned(body["end"])
                if "attendees" in body:
                    event.attendees = [entry["emailAddress"]["address"] for entry in body["attendees"]]
                if event.organizer:
                    api.notifications.extend(("update", address) for address in event.attendees)
                self.reply(200, event.resource())

            def op_event_delete(self, call, query, groups, _raw):
                event = self.event(groups[0])
                if event is None:
                    return
                del api.events[event.id]
                api.deleted_events.append(event.id)
                if event.organizer:
                    api.notifications.extend(("cancellation", address) for address in event.attendees)
                self.reply(204, None)

            # -- OneDrive -------------------------------------------------------------------
            def resolve(self, reference):
                """root, items/{id} or root:/a/b -> Item, replying 404 when absent."""
                if reference == "root":
                    return api.items[ROOT_ID]
                if reference.startswith("items/"):
                    item = api.items.get(reference[len("items/"):])
                else:
                    item = self.walk(reference[len("root:/"):].rstrip(":"))
                if item is None:
                    self.reply(404, error(404, "itemNotFound", "The resource could not be found."))
                return item

            def walk(self, path):
                current = api.items[ROOT_ID]
                for part in [segment for segment in path.split("/") if segment]:
                    current = next((item for item in api.items.values() if item.parent == current.id and item.name == part), None)
                    if current is None:
                        return None
                return current

            def op_children(self, call, query, groups, _raw):
                folder = self.resolve(groups[0])
                if folder is None:
                    return
                if not folder.folder:
                    self.reply(400, error(400, "invalidRequest", "Item is not a folder."))
                    return
                children = sorted((item for item in api.items.values() if item.parent == folder.id),
                                  key=lambda item: (not item.folder, item.name))
                # Graph answers /me/drive requests with /drives/{drive-id}/items/{id} continuation links.
                self.page(call, query, call["path"], [drive_resource(api, item) for item in children], self.top(query, 200),
                          link_path="/v1.0/drives/b!fixturedrive/items/" + folder.id + "/children")

            def op_item(self, call, query, groups, _raw):
                item = self.resolve(groups[0])
                if item is not None:
                    self.reply(200, drive_resource(api, item, download="$select" not in query))

            def download(self, call, token):
                assert call["auth"] is None, "The pre-authenticated download URL must not receive the token"
                ident = api.downloads.get(token)
                if ident is None:
                    self.reply(401, None, raw=b"expired", content_type="text/plain")
                    return
                data = api.items[ident].content
                action = (api.script.get("body") or [None]).pop(0) if api.script.get("body") else None
                if action == "corrupt":
                    self.reply(200, None, raw=bytes(reversed(data)), content_type="application/octet-stream")
                elif action == "short":
                    self.reply(200, None, raw=data[:-3], content_type="application/octet-stream", length=len(data))
                elif action == "gzip":
                    self.reply(200, None, raw=data, content_type="application/octet-stream",
                               headers={"Content-Encoding": "gzip"})
                else:
                    self.reply(200, None, raw=data, content_type="application/octet-stream")

            def place(self, path, behavior):
                """Parent folder and final name for an upload, applying conflictBehavior; None after a reply."""
                parent_path, _, name = path.rpartition("/")
                parent = self.walk(parent_path)
                if parent is None or not parent.folder:
                    self.reply(404, error(404, "itemNotFound", "The parent folder could not be found."))
                    return None
                existing = next((item for item in api.items.values() if item.parent == parent.id and item.name == name), None)
                if existing is not None:
                    if behavior == "fail":
                        self.reply(409, error(409, "nameAlreadyExists", "The specified item name already exists."))
                        return None
                    if behavior == "rename":
                        stem, dot, suffix = name.rpartition(".")
                        index = 1
                        while any(item.parent == parent.id and item.name == (f"{stem} {index}.{suffix}" if dot else f"{name} {index}")
                                  for item in api.items.values()):
                            index += 1
                        name = f"{stem} {index}.{suffix}" if dot else f"{name} {index}"
                        existing = None
                return parent, name, existing

            def store(self, parent, name, existing, data):
                if existing is not None:
                    existing.content = data
                    return existing, 200
                api.counter += 1
                item = Item(f"01FIXTURENEWITEM{api.counter:016d}", name, parent.id, content=data, sha256=True)
                api.items[item.id] = item
                return item, 201

            def op_simple_upload(self, call, query, groups, raw):
                behavior = query.get("@microsoft.graph.conflictBehavior", "replace")
                placed = self.place(groups[0], behavior)
                if placed is None:
                    return
                item, status = self.store(*placed, raw)
                override = None
                if (api.script.get("receipt") or [None])[0] == "badhash":
                    api.script["receipt"].pop(0)
                    override = "0" * 64
                body = drive_resource(api, item)
                if override:
                    body["file"]["hashes"]["sha256Hash"] = override
                self.reply(status, body)

            def op_create_session(self, call, query, groups, _raw):
                item = call["body"]["item"]
                behavior = item.get("@microsoft.graph.conflictBehavior", "replace")
                assert item.get("name") == groups[0].rpartition("/")[2], (item, groups)
                if self.place(groups[0], behavior) is None:
                    return
                sid = secrets.token_hex(8)
                api.sessions[sid] = {"path": groups[0], "behavior": behavior, "data": bytearray(), "total": None}
                self.reply(200, {"uploadUrl": api.base_url + "/upload/" + sid, "expirationDateTime": "2026-10-12T00:00:00Z",
                                 "nextExpectedRanges": ["0-"]})

            def chunk(self, call, sid, raw, method):
                assert call["auth"] is None, "The upload session URL must not receive the token"
                session = api.sessions.get(sid)
                if method == "DELETE":
                    if session is None:
                        self.reply(404, error(404, "itemNotFound", "The upload session was not found."))
                    else:
                        del api.sessions[sid]
                        self.reply(204, None)
                    return
                if session is None:
                    self.reply(404, error(404, "itemNotFound", "The upload session was not found."))
                    return
                match = re.fullmatch(r"bytes ([0-9]+)-([0-9]+)/([0-9]+)", call["content_range"] or "")
                assert match, call["content_range"]
                first, last, total = (int(value) for value in match.groups())
                if first != len(session["data"]) or last - first + 1 != len(raw):
                    self.reply(416, error(416, "invalidRange", "The fragment range is not the expected one."))
                    return
                if last + 1 < total:
                    assert len(raw) % 327680 == 0, "Upload fragments must be multiples of 320 KiB"
                session["data"].extend(raw)
                if last + 1 < total:
                    self.reply(202, {"expirationDateTime": "2026-10-12T00:00:00Z", "nextExpectedRanges": [f"{last + 1}-"]})
                    return
                placed = self.place(session["path"], session["behavior"])
                del api.sessions[sid]
                if placed is None:
                    return
                item, status = self.store(*placed, bytes(session["data"]))
                if (api.script.get("commit") or [None])[0] == "lost":
                    # The file is committed but the answer never arrives.
                    api.script["commit"].pop(0)
                    try:
                        self.connection.shutdown(socket.SHUT_RDWR)
                    except OSError:
                        pass
                    self.close_connection = True
                    return
                self.reply(status, drive_resource(api, item))

            def op_folder_create(self, call, query, groups, _raw):
                parent = self.resolve(groups[0])
                if parent is None:
                    return
                body = call["body"]
                assert body.get("folder") == {} and body.get("@microsoft.graph.conflictBehavior") == "fail", body
                if any(item.parent == parent.id and item.name == body["name"] for item in api.items.values()):
                    self.reply(409, error(409, "nameAlreadyExists", "The specified item name already exists."))
                    return
                api.counter += 1
                item = Item(f"01FIXTURENEWFOLDER{api.counter:014d}", body["name"], parent.id, folder=True)
                api.items[item.id] = item
                self.reply(201, drive_resource(api, item))

            # -- Teams ----------------------------------------------------------------------
            def op_channel_message(self, call, query, groups, _raw):
                if (groups[0], groups[1]) not in api.channels:
                    self.reply(404, error(404, "NotFound", "Channel not found."))
                    return
                body = call["body"]
                assert set(body) <= {"body", "subject"} and body["body"]["contentType"] in ("text", "html"), body
                api.counter += 1
                api.channel_messages.append((groups[0], groups[1], body))
                self.reply(201, {"id": str(1700000000000 + api.counter), "createdDateTime": "2026-10-11T08:00:00Z",
                                 "webUrl": "https://teams.example.test/l/message/" + str(api.counter),
                                 "body": body["body"], "from": {"user": {"displayName": "Owner"}}})

        server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        server.daemon_threads = True
        ca = ""
        if tls:
            assert shutil.which("openssl"), "OpenSSL is required for the disposable HTTPS fixture"
            certificate, key = directory / "ca.pem", directory / "key.pem"
            subprocess.run(["openssl", "req", "-x509", "-newkey", "rsa:2048", "-nodes", "-days", "1",
                            "-keyout", str(key), "-out", str(certificate), "-subj", "/CN=localhost",
                            "-addext", "subjectAltName=DNS:localhost,IP:127.0.0.1"], check=True, timeout=30,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
            context.load_cert_chain(certificate, key)
            server.socket = context.wrap_socket(server.socket, server_side=True)
            ca = str(certificate)
        thread = threading.Thread(target=server.serve_forever, kwargs={"poll_interval": 0.05})
        thread.start()
        api.base_url = ("https" if tls else "http") + "://127.0.0.1:" + str(server.server_port)
        api.ca_file = ca
        api.kinds = lambda: [call["kind"] for call in api.calls]
        try:
            yield api
        finally:
            api.closing.set()
            api.release.set()
            server.shutdown()
            server.server_close()
            thread.join(timeout=5)
            assert not thread.is_alive()
