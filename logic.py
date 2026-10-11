"""Pure Microsoft Graph v1.0 contracts: settings, requests, resources, paths and errors.

No network, filesystem, framework or third-party import happens here. Every
function is deterministic (the current time is always passed in), so the policy
can be checked directly by the tests.
"""

from collections.abc import Mapping
from datetime import datetime, timedelta, timezone
import hashlib
import ipaddress
import json
import math
import mimetypes
import re
from urllib.parse import quote, urlsplit

OFFICIAL_API = "https://graph.microsoft.com"
OFFICIAL_LOGIN = "https://login.microsoftonline.com"
VERSION_PREFIX = "/v1.0/"
DEFAULT_SCOPE = "https://graph.microsoft.com/.default offline_access"

MAIL = ("mail_list", "mail_get", "mail_send", "mail_move", "mail_mark_read", "mail_attachment_download", "mail_watch")
CALENDAR = ("events_list", "event_create", "event_update", "event_delete")
DRIVE = ("files_list", "file_download", "file_upload", "folder_create")
TEAMS = ("channel_message_send",)
OPERATIONS = MAIL + CALENDAR + DRIVE + TEAMS
WRITES = frozenset(("mail_send", "mail_move", "mail_mark_read", "event_create", "event_update", "event_delete",
                    "file_upload", "folder_create", "channel_message_send"))
SENDS = frozenset(("mail_send", "channel_message_send"))
DELETES = frozenset(("event_delete",))
# Graph documents simple uploads for files up to 4 MB; larger files use an upload session.
SIMPLE_UPLOAD_LIMIT = 4194304
UPLOAD_FRAGMENT = 327680
# sendMail is limited to 4 MB per request; base64 inflates attachments by a third.
SEND_REQUEST_LIMIT = 4194304
SEND_ATTACHMENTS_LIMIT = 3145728
MAX_ATTACHMENTS = 10
MAX_WINDOW = timedelta(days=366)
OUTPUT_LIMIT = 16777216
LOCAL_NAME_BYTES = 180
SEEN_LIMIT = 4096
TEAMS_TEXT_LIMIT = 28000
WELL_KNOWN_FOLDERS = ("inbox", "drafts", "sentitems", "deleteditems", "archive", "junkemail", "outbox")
FILE_ATTACHMENT = "#microsoft.graph.fileAttachment"
MESSAGE_FIELDS = ("id", "conversationId", "internetMessageId", "subject", "from", "sender", "toRecipients",
                  "ccRecipients", "bccRecipients", "replyTo", "receivedDateTime", "sentDateTime", "isRead", "isDraft",
                  "importance", "hasAttachments", "bodyPreview", "webLink", "parentFolderId", "categories", "flag")
MESSAGE_SELECT = ",".join(field for field in MESSAGE_FIELDS if field not in ("sender", "bccRecipients", "replyTo"))
EVENT_SELECT = ("id,subject,start,end,isAllDay,location,organizer,attendees,isOnlineMeeting,onlineMeeting,webLink,"
                "showAs,isCancelled,isOrganizer,bodyPreview,categories,type,seriesMasterId")
DRIVE_SELECT = "id,name,size,file,folder,createdDateTime,lastModifiedDateTime,webUrl,parentReference,eTag"
SHOW_AS = ("free", "tentative", "busy", "oof", "workingElsewhere")
ATTENDEE_TYPES = ("required", "optional", "resource")
IMPORTANCE = ("low", "normal", "high")
SYNC_RESET = frozenset(("syncStateNotFound", "SyncStateNotFound", "resyncRequired", "SyncStateInvalid",
                        "ErrorInvalidSyncStateData"))

DEFAULTS = {
    "enabled": False, "secret_ref": "", "tenant_id": "common", "operation": "mail_list",
    "allowed_operations": ["mail_list", "mail_get", "events_list", "files_list", "file_download"],
    "mail_folder": "inbox", "drive_folder": "", "calendar_id": "",
    "allow_writes": False, "allow_send": False, "allowed_recipients": [], "allowed_channels": [],
    "attachment_root": "", "upload_root": "", "allow_delete": False,
    "watch_mail": False, "watch_interval_sec": 300, "emit_existing": False, "existing_since_hours": 24,
    "watch_max_items": 50,
    "body_format": "text", "max_body_chars": 65536, "list_limit": 100, "page_size": 50, "max_pages": 10,
    "window_days": 7, "time_zone": "UTC", "conflict_behavior": "fail", "folder_exists": "reuse",
    "max_attachment_bytes": SEND_ATTACHMENTS_LIMIT, "max_upload_bytes": 52428800, "upload_chunk_bytes": 3276800,
    "max_download_bytes": 52428800, "max_kept_downloads": 20, "timeout_sec": 120, "max_request_bytes": 65536,
    "max_response_bytes": 4194304, "max_retries": 2, "max_retry_wait_sec": 30, "oauth_scope": DEFAULT_SCOPE,
    "api_base_url": OFFICIAL_API, "login_base_url": OFFICIAL_LOGIN, "ca_file": "",
}
FRAMEWORK_KEYS = {"execution", "runtime_path", "runtime_path_label"}
BOUNDS = {
    "watch_interval_sec": (60, 86400), "existing_since_hours": (1, 720), "watch_max_items": (1, 200),
    "max_body_chars": (256, 1048576), "list_limit": (1, 1000), "page_size": (1, 200), "max_pages": (1, 50),
    "window_days": (1, 366), "max_attachment_bytes": (1024, SEND_ATTACHMENTS_LIMIT),
    "max_upload_bytes": (1024, 262144000), "upload_chunk_bytes": (UPLOAD_FRAGMENT, 62914560),
    "max_download_bytes": (1024, 1073741824), "max_kept_downloads": (1, 1000), "timeout_sec": (5, 900),
    "max_request_bytes": (1024, 1048576), "max_response_bytes": (65536, 33554432), "max_retries": (0, 5),
    "max_retry_wait_sec": (1, 120),
}
CHOICES = {"body_format": ("text", "html"), "conflict_behavior": ("fail", "rename", "replace"),
           "folder_exists": ("reuse", "fail")}
BOOLEANS = ("enabled", "allow_writes", "allow_send", "allow_delete", "watch_mail", "emit_existing")

SECRET_REF = re.compile(r"secret://(?:workspace/[A-Za-z0-9_.-]{1,80}|project/[A-Za-z0-9_.-]{1,80}/[A-Za-z0-9_.-]{1,80})\Z")
GUID = re.compile(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\Z")
GRAPH_ID = re.compile(r"[A-Za-z0-9=_+/-]{16,512}\Z")
DRIVE_ID = re.compile(r"[A-Za-z0-9!._,-]{4,256}\Z")
CHANNEL_ID = re.compile(r"[A-Za-z0-9:@._=-]{3,256}\Z")
IDENTIFIER = re.compile(r"[A-Za-z0-9][A-Za-z0-9._:-]{0,127}\Z")
TOKEN_TEXT = re.compile(r"[\x21-\x7e]{1,16384}\Z")
SCOPE_ITEM = re.compile(r"[A-Za-z0-9._:/-]{1,200}\Z")
PROPERTY = re.compile(r"[A-Za-z][A-Za-z0-9]{0,63}\Z")
HOST_LABEL = re.compile(r"[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\Z")
LOCAL_PART = re.compile(r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]{1,64}\Z")
TIME_ZONE = re.compile(r"[A-Za-z0-9_+/.:()-][A-Za-z0-9 _+/.:()-]{0,99}\Z")
DATE_ONLY = re.compile(r"[0-9]{4}-[0-9]{2}-[0-9]{2}\Z")
LOCAL_TIME = re.compile(r"[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}(?::[0-9]{2}(?:\.[0-9]{1,7})?)?\Z")
ZONED_TIME = re.compile(r"[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}(?::[0-9]{2}(?:\.[0-9]{1,7})?)?(?:Z|[+-][0-9]{2}:[0-9]{2})\Z")
SHA256_HEX = re.compile(r"[0-9A-Fa-f]{64}\Z")
FORBIDDEN_NAME = set('"*:<>?/\\|')

REQUEST_KEYS = {
    "mail_list": {"operation", "request_id", "folder", "filter", "search", "unread_only", "received_after", "select",
                  "limit", "page_link"},
    "mail_get": {"operation", "request_id", "message_id", "body_format"},
    "mail_send": {"operation", "request_id", "to", "cc", "bcc", "subject", "body", "body_format", "attachments",
                  "importance", "save_to_sent_items"},
    "mail_move": {"operation", "request_id", "message_id", "destination"},
    "mail_mark_read": {"operation", "request_id", "message_id", "is_read"},
    "mail_attachment_download": {"operation", "request_id", "message_id", "attachment_id"},
    "mail_watch": {"operation", "request_id"},
    "events_list": {"operation", "request_id", "calendar_id", "start", "end", "limit", "page_link"},
    "event_create": {"operation", "request_id", "calendar_id", "subject", "body", "body_format", "start", "end",
                     "time_zone", "all_day", "location", "attendees", "show_as", "is_online_meeting"},
    "event_update": {"operation", "request_id", "event_id", "subject", "body", "body_format", "start", "end",
                     "time_zone", "all_day", "location", "attendees", "show_as"},
    "event_delete": {"operation", "request_id", "event_id"},
    "files_list": {"operation", "request_id", "path", "item_id", "limit", "page_link"},
    "file_download": {"operation", "request_id", "path", "item_id"},
    "file_upload": {"operation", "request_id", "path", "name", "folder"},
    "folder_create": {"operation", "request_id", "name", "folder"},
    "channel_message_send": {"operation", "request_id", "team_id", "channel_id", "content", "content_type", "subject"},
}
# Graph continuation links may use other spellings than the request (mailfolders('id'), /drives/{id}/,
# /users/{id}/); each family accepts only its own resources, so a page link cannot reach another family.
LINK_PATTERNS = {
    "mail_list": re.compile(r"/v1\.0/(?:me|users/[^/?]+)/mailfolders[/(]", re.I),
    "events_list": re.compile(r"/v1\.0/(?:me|users/[^/?]+)/(?:calendarview|calendars?[/(])", re.I),
    "files_list": re.compile(r"/v1\.0/(?:me/drive|drives/[^/?]+|users/[^/?]+/drive)/", re.I),
}
WATCH_LINK = re.compile(r"/v1\.0/(?:me|users/[^/?]+)/mailfolders[/(][^?]*/messages/delta(?:\?|\Z)", re.I)

MESSAGES = {
    "config": "Invalid Microsoft 365 settings.",
    "disabled": "Remote access is disabled; enable the block explicitly before contacting Microsoft 365.",
    "secret": "Unlock the wallet and store a Microsoft 365 grant (client_id and refresh_token JSON, or an access token) under secret_ref.",
    "invalid_input": "The request does not follow the documented contract; nothing was sent.",
    "request_too_large": "The request exceeds its size limit; nothing was sent.",
    "operation_not_allowed": "This operation is not listed in allowed_operations; nothing was sent.",
    "writes_disabled": "Writes are disabled; enable allow_writes in the properties; nothing was changed.",
    "send_disabled": "Sending is disabled; enable allow_send to send mail, invitations or channel messages; nothing was sent.",
    "delete_disabled": "Deletion is disabled; enable allow_delete in the properties; nothing was deleted.",
    "recipient_not_allowed": "A recipient or attendee is not in allowed_recipients; nothing was sent.",
    "channel_not_allowed": "This team channel is not in allowed_channels; nothing was posted.",
    "attachment_root_required": "Set attachment_root, the local folder attachments may be read from; nothing was sent.",
    "upload_root_required": "Set upload_root, the local folder files may be uploaded from; nothing was sent.",
    "path": "The local file cannot be read as a regular file; nothing was sent.",
    "path_denied": "The local file is outside its configured root folder; nothing was sent.",
    "file_too_large": "The file is larger than the configured size limit; nothing was transferred.",
    "storage": "Block storage is unavailable; this operation needs a project-scoped run.",
    "busy": "Another mail watch of this node holds the watch state; nothing was polled.",
    "state_invalid": "The stored mail watch state is unreadable; remove it from the node storage to start again.",
    "not_downloadable": "This item has no downloadable file content (folder, item or reference attachment).",
    "auth_failed": "Microsoft refused to issue an access token from the wallet grant; check the refresh token, tenant and scopes.",
    "bad_request": "Microsoft Graph rejected the request (HTTP 400); nothing was changed.",
    "unauthorized": "Microsoft Graph rejected the access token (HTTP 401); refresh or replace the wallet secret.",
    "forbidden": "The account lacks the delegated permission or access to this item (HTTP 403).",
    "not_found": "The item was not found (HTTP 404); check the ID or path.",
    "conflict": "An item with this name already exists (HTTP 409); nothing was overwritten.",
    "sync_reset": "Microsoft Graph asked for a new synchronization (HTTP 410); the watch restarts from its last position.",
    "rate_limited": "Microsoft Graph throttled the request (HTTP 429); retry after the indicated delay.",
    "storage_quota": "The mailbox or OneDrive storage quota is exhausted; nothing was stored.",
    "provider_error": "Microsoft Graph returned an unexpected error status; nothing is assumed changed.",
    "provider_unavailable": "Microsoft Graph is temporarily unavailable (HTTP 5xx); nothing was changed.",
    "provider_unavailable_uncertain": "Microsoft Graph answered HTTP 5xx after the change was sent; the change is uncertain.",
    "network": "The connection to Microsoft failed before the request was complete; nothing was changed.",
    "network_uncertain": "The connection failed after the change was sent; the change is uncertain.",
    "tls": "TLS verification failed for the Microsoft origin; nothing was sent.",
    "timeout": "The operation exceeded timeout_sec before any change was complete; nothing was changed.",
    "timeout_uncertain": "The operation exceeded timeout_sec after the change was sent; the change is uncertain.",
    "cancelled": "The operation was cancelled after the change was sent; the change is uncertain.",
    "cancelled_before_send": "The operation was cancelled before any change was sent; nothing was changed.",
    "protocol": "Microsoft returned an invalid, compressed or oversized response.",
    "internal": "The Microsoft 365 operation failed without a complete result.",
}


class GraphError(ValueError):
    """A content-free diagnostic with a stable machine code and an English message."""

    def __init__(self, code, detail=None, *, uncertain=False, http_status=None, retry_after=None, provider_code=None,
                 provider_message=None, provider_request_id=None, retryable=False):
        super().__init__(detail or MESSAGES.get(code, MESSAGES["internal"]))
        self.code, self.uncertain, self.retryable = code, bool(uncertain), bool(retryable)
        self.http_status, self.retry_after = http_status, retry_after
        self.provider_code, self.provider_message, self.provider_request_id = provider_code, provider_message, provider_request_id
        self.attempts = 0

    def envelope(self, **context):
        """Return the public error document; callers add identifiers only, never content or tokens."""
        result = {"code": self.code, "message": str(self), "uncertain": self.uncertain}
        for key, value in (("http_status", self.http_status), ("retry_after", self.retry_after),
                           ("provider_code", self.provider_code), ("provider_message", self.provider_message),
                           ("provider_request_id", self.provider_request_id)):
            if value is not None:
                result[key] = value
        if self.attempts:
            result["attempts"] = self.attempts
        for key, value in context.items():
            if value is not None:
                result[key] = value
        return result


def fail(code, detail=None, **extra):
    """Raise a structured failure."""
    raise GraphError(code, detail, **extra)


def thaw(value):
    """Turn immutable preparation snapshots back into plain JSON containers."""
    if isinstance(value, Mapping):
        return {key: thaw(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [thaw(item) for item in value]
    return value


def encoded(value, limit, code="invalid_input"):
    """Serialize strict finite JSON with bounded depth, element count and bytes."""
    pending, count = [(value, 0)], 0
    while pending:
        item, depth = pending.pop()
        count += 1
        if depth > 24 or count > 200000:
            fail(code, "JSON exceeds its depth or element limit.")
        if isinstance(item, dict):
            if any(type(key) is not str for key in item):
                fail(code, "JSON object keys must be strings.")
            pending.extend((child, depth + 1) for child in item.values())
        elif isinstance(item, (list, tuple)):
            pending.extend((child, depth + 1) for child in item)
        elif item is not None and type(item) not in (str, int, float, bool):
            fail(code, "A strict JSON value is required.")
        elif type(item) is float and not math.isfinite(item):
            fail(code, "Non-finite numbers are not accepted.")
    try:
        data = json.dumps(value, ensure_ascii=False, allow_nan=False, separators=(",", ":")).encode("utf-8")
    except (TypeError, ValueError, UnicodeError, RecursionError):
        fail(code, "Finite UTF-8 JSON is required.")
    if len(data) > limit:
        fail("request_too_large" if code == "invalid_input" else code, "JSON exceeds its byte limit.")
    return data


def strict_json(text, limit, code="invalid_input"):
    """Parse JSON text, rejecting duplicate keys, non-finite numbers and oversized input."""
    if isinstance(text, bytes):
        try:
            text = text.decode("utf-8", errors="strict")
        except UnicodeError:
            fail(code, "UTF-8 JSON is required.")
    if not isinstance(text, str):
        fail(code, "JSON text is required.")
    if len(text.encode("utf-8", errors="replace")) > limit:
        fail("request_too_large" if code == "invalid_input" else code, "JSON exceeds its byte limit.")

    def pairs(items):
        result = {}
        for key, item in items:
            if key in result:
                fail(code, "Duplicate JSON property.")
            result[key] = item
        return result

    def constant(_name):
        fail(code, "Non-finite numbers are not accepted.")

    try:
        value = json.loads(text, object_pairs_hook=pairs, parse_constant=constant)
    except GraphError:
        raise
    except (ValueError, RecursionError):
        fail(code, "Malformed JSON.")
    encoded(value, limit, code)
    return value


def boolean(value, name, code="config"):
    if type(value) is not bool:
        fail(code, f"{name} must be a boolean.")
    return value


def integer(value, name, low, high, code="config"):
    if type(value) is not int or not low <= value <= high:
        fail(code, f"{name} must be an integer between {low} and {high}.")
    return value


def text_value(value, name, limit, code="invalid_input", *, empty=False, multiline=False):
    """Bounded text without control characters (newlines and tabs only when multiline)."""
    if not isinstance(value, str) or len(value) > limit or (not empty and not value.strip()):
        fail(code, f"{name} must be {'' if empty else 'non-empty '}text of at most {limit} characters.")
    allowed = "\n\r\t" if multiline else ""
    if any((ord(char) < 32 and char not in allowed) or ord(char) == 127 for char in value):
        fail(code, f"{name} cannot contain control characters.")
    return value


def path_text(value, name):
    if not isinstance(value, str) or len(value) > 4096 or any(ord(char) < 32 or ord(char) == 127 for char in value):
        fail("config", f"{name} must be a bounded path without control characters.")
    value = value.strip()
    if value and (not value.startswith("/") or ".." in value.split("/")):
        fail("config", f"{name} must be empty or an absolute path without '..' components.")
    return value


def is_loopback(host):
    if host == "localhost":
        return True
    try:
        return ipaddress.ip_address(host).is_loopback
    except ValueError:
        return False


def host_name(host):
    """True for a syntactically valid DNS host name (not an IP literal)."""
    try:
        ipaddress.ip_address(host)
        return False
    except ValueError:
        pass
    return (0 < len(host) <= 253 and not host.endswith(".")
            and all(HOST_LABEL.fullmatch(label) for label in host.split(".")))


def api_origin(value, name):
    """Accept an HTTPS origin; plain HTTP only for a loopback host such as a local test double."""
    if not isinstance(value, str) or not 1 <= len(value) <= 2048 or any(ord(char) <= 32 or ord(char) >= 127 for char in value):
        fail("config", f"{name} must be an ASCII URL without spaces.")
    try:
        parsed = urlsplit(value)
        port = parsed.port
    except ValueError:
        fail("config", f"{name} has an invalid host or port.")
    host = (parsed.hostname or "").lower()
    if (parsed.scheme not in ("https", "http") or not host or parsed.username is not None
            or parsed.password is not None or parsed.query or parsed.fragment or "?" in value or "#" in value
            or parsed.path not in ("", "/") or "\\" in value or "%" in host):
        fail("config", f"{name} must be a bare https:// origin without credentials, path, query or fragment.")
    try:
        ipaddress.ip_address(host)
    except ValueError:
        if not host_name(host):
            fail("config", f"{name} has an invalid host name.")
    if port is not None and not 1 <= port <= 65535:
        fail("config", f"{name} has an invalid port.")
    if parsed.scheme == "http" and not is_loopback(host):
        fail("config", f"Plain HTTP is accepted for {name} only with a loopback test host; use HTTPS.")
    netloc = "[" + host + "]" if ":" in host else host
    return parsed.scheme + "://" + netloc + (":" + str(port) if port is not None else "")


def split_url(value):
    """Return (origin, path?query) of an absolute http(s) URL, or None when it is not one."""
    if not isinstance(value, str) or not 1 <= len(value) <= 16384 or any(ord(c) <= 32 or ord(c) >= 127 for c in value):
        return None
    try:
        parsed = urlsplit(value)
        port = parsed.port
    except ValueError:
        return None
    host = (parsed.hostname or "").lower()
    if (parsed.scheme not in ("https", "http") or not host or parsed.username is not None or parsed.password is not None
            or parsed.fragment or "#" in value or "\\" in value or not parsed.path.startswith("/")):
        return None
    netloc = "[" + host + "]" if ":" in host else host
    origin = parsed.scheme + "://" + netloc + (":" + str(port) if port is not None else "")
    return origin, parsed.path + ("?" + parsed.query if parsed.query else ""), host


def api_link(value, config, pattern=None):
    """A Graph continuation link (nextLink/deltaLink): same origin as api_base_url, under /v1.0/, of one family."""
    parts = split_url(value)
    if parts is None or parts[0] != config["api_base_url"] or not parts[1].startswith(VERSION_PREFIX):
        return None
    if pattern is not None and not pattern.match(parts[1]):
        return None
    return parts[1]


def transfer_link(value, config):
    """A pre-authenticated download or upload URL: HTTPS with a DNS host, or the configured test origin.

    These URLs never receive the access token.
    """
    parts = split_url(value)
    if parts is None:
        return None
    origin, target, host = parts
    if origin == config["api_base_url"] or (origin.startswith("https://") and host_name(host)):
        return origin, target
    return None


def secret_reference(value):
    if not isinstance(value, str) or len(value) > 256:
        fail("config", "secret_ref must be a wallet reference such as secret://workspace/microsoft_365.")
    value = value.strip()
    if value and not SECRET_REF.fullmatch(value):
        fail("config", "secret_ref must be a wallet reference such as secret://workspace/microsoft_365, never a token.")
    return value


def tenant(value):
    """common, organizations, consumers, a directory GUID or a verified domain name."""
    if not isinstance(value, str):
        fail("config", "tenant_id must be text.")
    value = value.strip().lower()
    if value in ("common", "organizations", "consumers") or GUID.fullmatch(value):
        return value
    if host_name(value) and "." in value:
        return value
    fail("config", "tenant_id must be common, organizations, consumers, a tenant GUID or a domain such as contoso.com.")


def scope_text(value):
    if (not isinstance(value, str) or not value.strip() or len(value) > 1024
            or any(ord(char) < 32 or ord(char) == 127 for char in value)):
        fail("config", "oauth_scope must be space-separated delegated scopes.")
    items = value.split()
    if len(items) > 30 or any(not SCOPE_ITEM.fullmatch(item) for item in items):
        fail("config", "oauth_scope must be space-separated delegated scopes such as https://graph.microsoft.com/.default.")
    return " ".join(items)


def time_zone(value, name="time_zone", code="config"):
    if not isinstance(value, str) or not TIME_ZONE.fullmatch(value.strip()):
        fail(code, f"{name} must be a time zone name such as UTC, Europe/Paris or W. Europe Standard Time.")
    return value.strip()


def address(value, name, code="invalid_input"):
    """One e-mail address, lowercased domain; display names travel separately."""
    if not isinstance(value, str) or len(value) > 320:
        fail(code, f"{name} must be an e-mail address.")
    value = value.strip()
    local, at, domain = value.rpartition("@")
    domain = domain.lower()
    if not at or not LOCAL_PART.fullmatch(local) or ".." in local or not host_name(domain) or "." not in domain:
        fail(code, f"{name} must be an e-mail address such as someone@example.com.")
    return local + "@" + domain


def recipient_rule(value):
    """An allowlist entry: an exact address, or @domain for every address of exactly that domain."""
    if isinstance(value, str) and value.strip().startswith("@"):
        domain = value.strip()[1:].lower()
        if not host_name(domain) or "." not in domain:
            fail("config", "allowed_recipients entries are addresses or @domain.tld.")
        return "@" + domain
    return address(value, "allowed_recipients entry", "config").lower()


def recipient_allowed(value, rules):
    lowered = value.lower()
    return lowered in rules or "@" + lowered.rpartition("@")[2] in rules


def channel_rule(value):
    """An allowed Teams channel: '<team GUID>/<channel ID>'."""
    if not isinstance(value, str):
        fail("config", "allowed_channels entries are '<team id>/<channel id>' text.")
    team, slash, channel = value.strip().partition("/")
    if not slash or not GUID.fullmatch(team.lower()) or not CHANNEL_ID.fullmatch(channel):
        fail("config", "allowed_channels entries are '<team id>/<channel id>', for example "
             "00000000-0000-0000-0000-000000000000/19:example@thread.tacv2.")
    return team.lower() + "/" + channel


def mail_folder(value, name="mail_folder", code="config"):
    """A well-known folder name (inbox, archive...) or a mail folder ID."""
    if not isinstance(value, str):
        fail(code, f"{name} must be a well-known folder name or a folder ID.")
    text = value.strip()
    if text.lower() in WELL_KNOWN_FOLDERS:
        return text.lower()
    if GRAPH_ID.fullmatch(text):
        return text
    fail(code, f"{name} must be one of {', '.join(WELL_KNOWN_FOLDERS)} or a mail folder ID.")


def graph_id(value, name, code="invalid_input"):
    if not isinstance(value, str) or not GRAPH_ID.fullmatch(value.strip()):
        fail(code, f"{name} must be a Microsoft Graph ID (16-512 letters, digits, =, _, +, / or -).")
    return value.strip()


def drive_item_id(value, name="item_id"):
    if not isinstance(value, str) or not DRIVE_ID.fullmatch(value.strip()):
        fail("invalid_input", f"{name} must be a OneDrive item ID (letters, digits, !, ., _, , or -).")
    return value.strip()


def drive_name(value, name, code="invalid_input"):
    """One OneDrive name: no control characters, no \" * : < > ? / \\ |, no surrounding spaces or dot-only."""
    if (not isinstance(value, str) or not value or len(value) > 255 or value != value.strip()
            or value in (".", "..") or value.endswith(".")):
        fail(code, f"{name} must be 1-255 characters without surrounding spaces or a trailing dot.")
    if any(ord(char) < 32 or ord(char) == 127 or char in FORBIDDEN_NAME for char in value):
        fail(code, f"{name} cannot contain control characters or any of \" * : < > ? / \\ |.")
    return value


def drive_path(value, name, code="invalid_input"):
    """A slash-separated OneDrive path; returns its validated segments (empty for the drive root)."""
    if not isinstance(value, str) or len(value) > 1024:
        fail(code, f"{name} must be a OneDrive path such as Reports/2026.")
    segments = [segment for segment in value.strip().strip("/").split("/")] if value.strip().strip("/") else []
    if len(segments) > 32:
        fail(code, f"{name} is nested too deeply.")
    return [drive_name(segment, name + " segment", code) for segment in segments]


def configuration(raw):
    """Pure settings validation; no wallet, filesystem or network access."""
    raw = thaw(raw)
    if not isinstance(raw, Mapping):
        fail("config", "Settings must be a JSON object.")
    unknown = set(raw) - set(DEFAULTS) - FRAMEWORK_KEYS
    if unknown:
        fail("config", "Unknown setting: " + sorted(unknown)[0][:64] + ". Credentials belong in the wallet.")
    config = {**{key: thaw(value) for key, value in DEFAULTS.items()},
              **{key: value for key, value in raw.items() if key in DEFAULTS}}
    for key in BOOLEANS:
        boolean(config[key], key)
    for key, (low, high) in BOUNDS.items():
        integer(config[key], key, low, high)
    if config["upload_chunk_bytes"] % UPLOAD_FRAGMENT:
        fail("config", f"upload_chunk_bytes must be a multiple of {UPLOAD_FRAGMENT} (320 KiB).")
    for key, allowed in CHOICES.items():
        if config[key] not in allowed:
            fail("config", f"{key} must be one of: " + ", ".join(allowed) + ".")
    config["secret_ref"] = secret_reference(config["secret_ref"])
    config["tenant_id"] = tenant(config["tenant_id"])
    if config["operation"] not in OPERATIONS:
        fail("config", "operation must be one of: " + ", ".join(OPERATIONS) + ".")
    allowed = config["allowed_operations"]
    if (not isinstance(allowed, list) or not 1 <= len(allowed) <= len(OPERATIONS)
            or any(item not in OPERATIONS for item in allowed) or len(set(allowed)) != len(allowed)):
        fail("config", "allowed_operations must list distinct operations among " + ", ".join(OPERATIONS) + ".")
    config["mail_folder"] = mail_folder(config["mail_folder"])
    if not isinstance(config["drive_folder"], str):
        fail("config", "drive_folder must be a OneDrive path.")
    config["drive_folder"] = "/".join(drive_path(config["drive_folder"], "drive_folder", "config"))
    if not isinstance(config["calendar_id"], str):
        fail("config", "calendar_id must be empty or a calendar ID.")
    config["calendar_id"] = config["calendar_id"].strip()
    if config["calendar_id"]:
        config["calendar_id"] = graph_id(config["calendar_id"], "calendar_id", "config")
    for key, rule, limit in (("allowed_recipients", recipient_rule, 200), ("allowed_channels", channel_rule, 50)):
        items = config[key]
        if not isinstance(items, list) or len(items) > limit:
            fail("config", f"{key} must be a JSON array of at most {limit} entries.")
        config[key] = list(dict.fromkeys(rule(item) for item in items))
    config["attachment_root"] = path_text(config["attachment_root"], "attachment_root")
    config["upload_root"] = path_text(config["upload_root"], "upload_root")
    config["ca_file"] = path_text(config["ca_file"], "ca_file")
    config["time_zone"] = time_zone(config["time_zone"])
    config["oauth_scope"] = scope_text(config["oauth_scope"])
    config["api_base_url"] = api_origin(config["api_base_url"], "api_base_url")
    config["login_base_url"] = api_origin(config["login_base_url"], "login_base_url")
    selectable = set(selectable_operations(config))
    if selectable & WRITES and not config["allow_writes"]:
        fail("config", "Write operations (" + ", ".join(sorted(selectable & WRITES)) + ") need allow_writes; enable it or remove them.")
    if selectable & SENDS and not config["allow_send"]:
        fail("config", "mail_send and channel_message_send need allow_send; enable it or remove them.")
    if selectable & DELETES and not config["allow_delete"]:
        fail("config", "event_delete needs allow_delete; enable it or remove it.")
    if config["conflict_behavior"] == "replace" and not config["allow_delete"]:
        fail("config", "conflict_behavior replace supersedes existing files and needs allow_delete.")
    if config["enabled"]:
        if not config["secret_ref"]:
            fail("config", "Remote access requires secret_ref, a wallet reference holding the Microsoft 365 grant.")
        if "mail_send" in selectable and not config["allowed_recipients"]:
            fail("config", "mail_send needs allowed_recipients, the addresses or @domains mail may be sent to.")
        if "channel_message_send" in selectable and not config["allowed_channels"]:
            fail("config", "channel_message_send needs allowed_channels, the team channels it may post to.")
        if "file_upload" in selectable and not config["upload_root"]:
            fail("config", "file_upload needs upload_root, the local folder files may be uploaded from.")
    return config


def selectable_operations(config):
    return [name for name in OPERATIONS if name == config["operation"] or name in config["allowed_operations"]]


def credentials(value):
    """Validate the wallet value: a bare access token, {access_token} or a refresh-token grant."""
    secret = value
    if isinstance(secret, str):
        text = secret.strip()
        if text.startswith("{"):
            try:
                secret = strict_json(text, 65536, "secret")
            except GraphError:
                fail("secret")
        else:
            secret = {"access_token": text}
    if not isinstance(secret, Mapping):
        fail("secret")
    keys = set(secret)
    if keys not in ({"access_token"}, {"client_id", "refresh_token"}, {"client_id", "client_secret", "refresh_token"}):
        fail("secret")
    for item in secret.values():
        if not isinstance(item, str) or not TOKEN_TEXT.fullmatch(item):
            fail("secret")
    return dict(secret)


def redact(text, secrets):
    """Mask every credential value wherever it could appear."""
    text = "" if text is None else str(text)
    for secret in secrets:
        if secret and len(secret) >= 4:
            text = text.replace(secret, "***")
    return text


def clean(value, limit):
    """Bounded single-line provider text with control characters removed."""
    if not isinstance(value, str):
        return None
    value = "".join(" " if ord(char) < 32 or ord(char) == 127 else char for char in value).strip()
    return value[:limit] if value else None


# ---------------------------------------------------------------------------
# Time


def utc_text(moment):
    return moment.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def instant(value, name):
    """RFC 3339 date-time with an offset, or a date taken as midnight UTC; returns an aware UTC datetime."""
    if not isinstance(value, str) or len(value) > 40:
        fail("invalid_input", f"{name} must be a date such as 2026-01-31 or an RFC 3339 date-time with an offset.")
    text = value.strip()
    if DATE_ONLY.fullmatch(text):
        text += "T00:00:00Z"
    if not ZONED_TIME.fullmatch(text):
        fail("invalid_input", f"{name} must be a date such as 2026-01-31 or an RFC 3339 date-time with an offset.")
    try:
        return parse_time(text).astimezone(timezone.utc)
    except ValueError:
        fail("invalid_input", f"{name} is not a valid calendar date-time.")


def parse_time(text):
    """datetime.fromisoformat with Z and up to seven fractional digits (Graph uses seven)."""
    text = text.replace("Z", "+00:00")
    match = re.fullmatch(r"(.*T[0-9:]+)\.([0-9]+)(.*)", text)
    if match:
        text = match.group(1) + "." + (match.group(2) + "000000")[:6] + match.group(3)
    return datetime.fromisoformat(text)


def event_time(value, name, zone, all_day):
    """Event start/end: (dateTime, timeZone, comparable). Offsets convert to UTC; local times use zone."""
    if not isinstance(value, str) or len(value) > 40:
        fail("invalid_input", f"{name} must be a date-time such as 2026-10-12T09:00.")
    text = value.strip()
    try:
        if all_day:
            if not DATE_ONLY.fullmatch(text):
                fail("invalid_input", f"All-day events take dates only: {name} must look like 2026-10-12.")
            moment = datetime.fromisoformat(text + "T00:00:00")
            return moment.strftime("%Y-%m-%dT%H:%M:%S"), zone, ("local", moment)
        if ZONED_TIME.fullmatch(text):
            moment = parse_time(text).astimezone(timezone.utc)
            return moment.strftime("%Y-%m-%dT%H:%M:%S"), "UTC", ("utc", moment)
        if LOCAL_TIME.fullmatch(text):
            moment = parse_time(text)
            return moment.strftime("%Y-%m-%dT%H:%M:%S"), zone, ("local", moment)
    except ValueError:
        fail("invalid_input", f"{name} is not a valid calendar date-time.")
    fail("invalid_input", f"{name} must be a local date-time such as 2026-10-12T09:00 or include an offset.")


# ---------------------------------------------------------------------------
# Requests


def identifier(value):
    if value is None:
        return None
    if not isinstance(value, str) or not IDENTIFIER.fullmatch(value):
        fail("invalid_input", "request_id needs 1-128 ASCII letters, digits or ._:- characters.")
    return value


def document(value, config):
    """Read a JSON object; other text is a bare trigger for the configured operation."""
    if isinstance(value, Mapping):
        result = thaw(value)
        encoded(result, config["max_request_bytes"])
        return result
    if isinstance(value, (list, tuple)):
        fail("invalid_input", "The request must be one JSON object, not an array.")
    if isinstance(value, bytes):
        try:
            value = value.decode("utf-8", errors="strict")
        except UnicodeError:
            fail("invalid_input", "The request must be UTF-8 text or JSON.")
    if not isinstance(value, str):
        fail("invalid_input", "The request must be UTF-8 text or JSON.")
    text = value.strip()
    if text.startswith("{") or text.startswith("["):
        parsed = strict_json(text, config["max_request_bytes"])
        if not isinstance(parsed, dict):
            fail("invalid_input", "The request must be one JSON object, not an array.")
        return parsed
    if len(value.encode("utf-8", errors="replace")) > config["max_request_bytes"]:
        fail("request_too_large")
    return {}


def local_path(value, name="path"):
    """A local path; surrounding whitespace such as a trailing newline from a text source is ignored."""
    value = value.strip() if isinstance(value, str) else value
    if (not isinstance(value, str) or not value or len(value.encode("utf-8", errors="replace")) > 4096
            or any(ord(char) < 32 or ord(char) == 127 for char in value)):
        fail("invalid_input", f"{name} must be a local file path without control characters.")
    return value


def limit_value(data, config):
    return integer(data.get("limit", config["list_limit"]), "limit", 1, config["list_limit"], "invalid_input")


def page_link(data, config, operation):
    link = data.get("page_link")
    if link is None:
        return None
    if api_link(link, config, LINK_PATTERNS[operation]) is None:
        fail("invalid_input", "page_link must be the next_page_link of a previous " + operation + " result.")
    return link


def people(value, name, rules, *, attendees=False):
    """Recipients or attendees: addresses or {address, name[, type]}; every address must be allowlisted."""
    if not isinstance(value, list) or len(value) > 100:
        fail("invalid_input", f"{name} must be a JSON array of at most 100 recipients.")
    result = []
    for item in value:
        entry = {"address": None, "name": None}
        if isinstance(item, str):
            entry["address"] = address(item, name)
        elif isinstance(item, Mapping) and set(item) <= ({"address", "name", "type"} if attendees else {"address", "name"}):
            entry["address"] = address(item.get("address"), name)
            if item.get("name") is not None:
                entry["name"] = text_value(item["name"], name + " name", 256)
            if attendees:
                entry["type"] = item.get("type", "required")
                if entry["type"] not in ATTENDEE_TYPES:
                    fail("invalid_input", "Attendee type must be required, optional or resource.")
        else:
            fail("invalid_input", f"{name} entries are addresses or objects with address and name.")
        if attendees and "type" not in entry:
            entry["type"] = "required"
        if not recipient_allowed(entry["address"], rules):
            fail("recipient_not_allowed")
        result.append(entry)
    return result


def body_type(value, default):
    choice = default if value is None else value
    if choice not in ("text", "html"):
        fail("invalid_input", "body_format must be text or html.")
    return choice


def parse_request(value, config, now):
    """Return one normalized operation request; enforce the allowlist and gates before any wallet access."""
    data = document(value, config)
    operation = data.get("operation", config["operation"])
    if operation not in OPERATIONS:
        fail("invalid_input", "operation must be one of: " + ", ".join(OPERATIONS) + ".")
    request = {"operation": operation, "request_id": identifier(data.get("request_id"))}
    if operation not in selectable_operations(config):
        fail("operation_not_allowed")
    if operation in WRITES and not config["allow_writes"]:
        fail("writes_disabled")
    if operation in SENDS and not config["allow_send"]:
        fail("send_disabled")
    if operation in DELETES and not config["allow_delete"]:
        fail("delete_disabled")
    unknown = set(data) - REQUEST_KEYS[operation]
    if unknown:
        fail("invalid_input", "Unknown field for " + operation + ": " + sorted(unknown)[0][:64] + ".")
    builder = BUILDERS.get(operation)
    if builder is not None:
        builder(request, data, config, now)
    return request


def build_mail_list(request, data, config, _now):
    request["limit"] = limit_value(data, config)
    request["page_link"] = page_link(data, config, "mail_list")
    if request["page_link"] is not None:
        if set(data) - {"operation", "request_id", "limit", "page_link"}:
            fail("invalid_input", "page_link continues a previous listing; send it without other query fields.")
        return
    request["folder"] = mail_folder(data.get("folder", config["mail_folder"]), "folder", "invalid_input")
    raw_filter, search = data.get("filter"), data.get("search")
    helpers = data.get("unread_only") is not None or data.get("received_after") is not None
    if sum(item is not None and item is not False for item in (raw_filter, search)) + bool(helpers) > 1:
        fail("invalid_input", "Use one of filter, search, or unread_only/received_after in a request.")
    request["filter"], request["search"], request["order"] = None, None, None
    if raw_filter is not None:
        request["filter"] = text_value(raw_filter, "filter", 1024)
    elif search is not None:
        text = text_value(search, "search", 255)
        if '"' in text or "\\" in text:
            fail("invalid_input", "search cannot contain double quotes or backslashes.")
        request["search"] = text.strip()
    else:
        # Graph needs the ordered property first in $filter when both are used.
        clauses = ["receivedDateTime ge " + (utc_text(instant(data["received_after"], "received_after"))
                                             if data.get("received_after") is not None else "1900-01-01T00:00:00Z")]
        if boolean(data.get("unread_only", False), "unread_only", "invalid_input"):
            clauses.append("isRead eq false")
        request["filter"] = " and ".join(clauses) if helpers else None
        request["order"] = "receivedDateTime desc"
    select = data.get("select")
    if select is None:
        request["select"] = MESSAGE_SELECT
    else:
        if (not isinstance(select, list) or not 1 <= len(select) <= len(MESSAGE_FIELDS)
                or any(item not in MESSAGE_FIELDS for item in select)):
            fail("invalid_input", "select must list message properties among: " + ", ".join(MESSAGE_FIELDS) + ".")
        request["select"] = ",".join(dict.fromkeys(["id", *select]))


def build_mail_get(request, data, config, _now):
    request["message_id"] = graph_id(data.get("message_id"), "message_id")
    request["body_format"] = body_type(data.get("body_format"), config["body_format"])


def build_mail_send(request, data, config, _now):
    rules = config["allowed_recipients"]
    for key in ("to", "cc", "bcc"):
        request[key] = people(data.get(key, []), key, rules)
    total = len(request["to"]) + len(request["cc"]) + len(request["bcc"])
    if not request["to"] or total > 100:
        fail("invalid_input", "mail_send needs at least one to recipient and at most 100 recipients in total.")
    request["subject"] = text_value(data.get("subject", ""), "subject", 255, empty=True)
    request["body"] = text_value(data.get("body", ""), "body", 1048576, empty=True, multiline=True)
    request["body_format"] = body_type(data.get("body_format"), "text")
    importance = data.get("importance", "normal")
    if importance not in IMPORTANCE:
        fail("invalid_input", "importance must be low, normal or high.")
    request["importance"] = importance
    request["save_to_sent_items"] = boolean(data.get("save_to_sent_items", True), "save_to_sent_items", "invalid_input")
    attachments = data.get("attachments", [])
    if not isinstance(attachments, list) or len(attachments) > MAX_ATTACHMENTS:
        fail("invalid_input", f"attachments must be a JSON array of at most {MAX_ATTACHMENTS} local paths.")
    request["attachments"] = [local_path(item, "attachment path") for item in attachments]
    if request["attachments"] and not config["attachment_root"]:
        fail("attachment_root_required")


def build_mail_move(request, data, _config, _now):
    request["message_id"] = graph_id(data.get("message_id"), "message_id")
    if data.get("destination") is None:
        fail("invalid_input", "mail_move needs destination, a well-known folder name or a folder ID.")
    request["destination"] = mail_folder(data["destination"], "destination", "invalid_input")


def build_mail_mark_read(request, data, _config, _now):
    request["message_id"] = graph_id(data.get("message_id"), "message_id")
    request["is_read"] = boolean(data.get("is_read", True), "is_read", "invalid_input")


def build_attachment(request, data, _config, _now):
    request["message_id"] = graph_id(data.get("message_id"), "message_id")
    request["attachment_id"] = graph_id(data.get("attachment_id"), "attachment_id")


def calendar_choice(data, config):
    value = data.get("calendar_id")
    if value is None or value == "":
        return config["calendar_id"]
    return graph_id(value, "calendar_id")


def build_events_list(request, data, config, now):
    request["limit"] = limit_value(data, config)
    request["page_link"] = page_link(data, config, "events_list")
    if request["page_link"] is not None:
        if set(data) - {"operation", "request_id", "limit", "page_link"}:
            fail("invalid_input", "page_link continues a previous listing; send it without other query fields.")
        return
    request["calendar_id"] = calendar_choice(data, config)
    start = instant(data["start"], "start") if data.get("start") is not None else now.astimezone(timezone.utc)
    end = instant(data["end"], "end") if data.get("end") is not None else start + timedelta(days=config["window_days"])
    if end <= start or end - start > MAX_WINDOW:
        fail("invalid_input", "end must be after start, and the window at most 366 days.")
    request["start"], request["end"] = utc_text(start), utc_text(end)


def event_fields(request, data, config, *, creating):
    """Shared event_create / event_update fields; times are validated against each other."""
    zone = time_zone(data["time_zone"], "time_zone", "invalid_input") if data.get("time_zone") is not None else config["time_zone"]
    all_day = boolean(data.get("all_day", False), "all_day", "invalid_input")
    fields = {}
    if data.get("subject") is not None or creating:
        fields["subject"] = text_value(data.get("subject", ""), "subject", 255, empty=not creating)
    if data.get("body") is not None:
        fields["body"] = {"contentType": body_type(data.get("body_format"), "text"),
                          "content": text_value(data["body"], "body", 262144, empty=True, multiline=True)}
    elif data.get("body_format") is not None:
        fail("invalid_input", "body_format needs body.")
    stamps = {}
    for key in ("start", "end"):
        if data.get(key) is not None:
            text, used, comparable = event_time(data[key], key, zone, all_day)
            fields[key] = {"dateTime": text, "timeZone": used}
            stamps[key] = comparable
        elif creating:
            fail("invalid_input", "event_create needs start and end.")
    if len(stamps) == 2:
        if stamps["start"][0] != stamps["end"][0]:
            fail("invalid_input", "start and end must both include an offset or both be local times.")
        if stamps["end"][1] <= stamps["start"][1]:
            fail("invalid_input", "end must be after start.")
        if all_day and (stamps["end"][1] - stamps["start"][1]).days < 1:
            fail("invalid_input", "An all-day event ends on a later date than it starts.")
    if data.get("all_day") is not None:
        if not creating and all_day and len(stamps) != 2:
            fail("invalid_input", "Changing to an all-day event needs both start and end dates.")
        fields["isAllDay"] = all_day
    if data.get("location") is not None:
        fields["location"] = {"displayName": text_value(data["location"], "location", 255, empty=True)}
    if data.get("show_as") is not None:
        if data["show_as"] not in SHOW_AS:
            fail("invalid_input", "show_as must be one of: " + ", ".join(SHOW_AS) + ".")
        fields["showAs"] = data["show_as"]
    if data.get("is_online_meeting") is not None:
        fields["isOnlineMeeting"] = boolean(data["is_online_meeting"], "is_online_meeting", "invalid_input")
    if data.get("attendees") is not None:
        attendees = people(data["attendees"], "attendees", config["allowed_recipients"], attendees=True)
        if attendees and not config["allow_send"]:
            fail("send_disabled", "Attendees receive invitations: enable allow_send to invite people; nothing was sent.")
        request["attendees"] = attendees
        fields["attendees"] = [{"emailAddress": {"address": item["address"], **({"name": item["name"]} if item["name"] else {})},
                                "type": item["type"]} for item in attendees]
    request["fields"] = fields


def build_event_create(request, data, config, _now):
    request["calendar_id"] = calendar_choice(data, config)
    event_fields(request, data, config, creating=True)


def build_event_update(request, data, config, _now):
    request["event_id"] = graph_id(data.get("event_id"), "event_id")
    event_fields(request, data, config, creating=False)
    if not request["fields"]:
        fail("invalid_input", "event_update needs at least one field to change.")


def build_event_delete(request, data, _config, _now):
    request["event_id"] = graph_id(data.get("event_id"), "event_id")


def item_target(request, data, config, *, allow_root=True):
    """files_list / file_download: an item ID, or a path relative to drive_folder."""
    if data.get("item_id") is not None and data.get("path") is not None:
        fail("invalid_input", "Give either item_id or path, not both.")
    if data.get("item_id") is not None:
        request["item_id"], request["path"] = drive_item_id(data["item_id"]), None
        return
    base = drive_path(config["drive_folder"], "drive_folder", "config")
    extra = drive_path(data.get("path", ""), "path")
    if not allow_root and not extra and not base:
        fail("invalid_input", "file_download needs path or item_id.")
    request["item_id"], request["path"] = None, "/".join(base + extra)


def build_files_list(request, data, config, _now):
    request["limit"] = limit_value(data, config)
    request["page_link"] = page_link(data, config, "files_list")
    if request["page_link"] is not None:
        if set(data) - {"operation", "request_id", "limit", "page_link"}:
            fail("invalid_input", "page_link continues a previous listing; send it without other query fields.")
        return
    item_target(request, data, config)


def build_file_download(request, data, config, _now):
    if data.get("item_id") is None and data.get("path") is None:
        fail("invalid_input", "file_download needs path or item_id.")
    item_target(request, data, config, allow_root=False)


def build_file_upload(request, data, config, _now):
    if data.get("path") is None:
        fail("invalid_input", "file_upload needs path, a local file inside upload_root (or send it on the file input).")
    request["path"] = local_path(data["path"])
    request["folder"] = "/".join(drive_path(config["drive_folder"], "drive_folder", "config")
                                 + drive_path(data.get("folder", ""), "folder"))
    if data.get("name") is not None:
        request["name"] = drive_name(data["name"], "name")


def build_folder_create(request, data, config, _now):
    if data.get("name") is None:
        fail("invalid_input", "folder_create needs name.")
    request["name"] = drive_name(data["name"], "name")
    request["folder"] = "/".join(drive_path(config["drive_folder"], "drive_folder", "config")
                                 + drive_path(data.get("folder", ""), "folder"))


def build_channel_message(request, data, config, _now):
    team, channel = data.get("team_id"), data.get("channel_id")
    if team is None and channel is None and len(config["allowed_channels"]) == 1:
        team, channel = config["allowed_channels"][0].split("/", 1)
    if not isinstance(team, str) or not isinstance(channel, str):
        fail("invalid_input", "channel_message_send needs team_id and channel_id (optional with one allowed channel).")
    team, channel = team.strip().lower(), channel.strip()
    if not GUID.fullmatch(team) or not CHANNEL_ID.fullmatch(channel):
        fail("invalid_input", "team_id must be a team GUID and channel_id a channel ID such as 19:...@thread.tacv2.")
    if team + "/" + channel not in config["allowed_channels"]:
        fail("channel_not_allowed")
    request["team_id"], request["channel_id"] = team, channel
    content_type = data.get("content_type", "text")
    if content_type not in ("text", "html"):
        fail("invalid_input", "content_type must be text or html.")
    request["content_type"] = content_type
    request["content"] = text_value(data.get("content"), "content", TEAMS_TEXT_LIMIT, multiline=True)
    request["subject"] = text_value(data["subject"], "subject", 255) if data.get("subject") is not None else None


BUILDERS = {
    "mail_list": build_mail_list, "mail_get": build_mail_get, "mail_send": build_mail_send, "mail_move": build_mail_move,
    "mail_mark_read": build_mail_mark_read, "mail_attachment_download": build_attachment,
    "events_list": build_events_list, "event_create": build_event_create, "event_update": build_event_update,
    "event_delete": build_event_delete, "files_list": build_files_list, "file_download": build_file_download,
    "file_upload": build_file_upload, "folder_create": build_folder_create, "channel_message_send": build_channel_message,
}


def file_request(value, config):
    """A file/path event uploads that file into drive_folder under its local name."""
    if isinstance(value, bytes):
        try:
            value = value.decode("utf-8", errors="strict")
        except UnicodeError:
            fail("invalid_input", "The file input must be a UTF-8 path.")
    if "file_upload" not in selectable_operations(config):
        fail("operation_not_allowed", "The file input uploads to OneDrive; add file_upload to allowed_operations. Nothing was sent.")
    if not config["allow_writes"]:
        fail("writes_disabled")
    return {"operation": "file_upload", "request_id": None, "path": local_path(value), "folder": config["drive_folder"]}


def guessed_type(name):
    guess, _ = mimetypes.guess_type(name, strict=False)
    return guess if isinstance(guess, str) and re.fullmatch(r"[A-Za-z0-9.+-]+/[A-Za-z0-9.+-]+", guess) \
        else "application/octet-stream"


# ---------------------------------------------------------------------------
# Graph paths


def segment(value):
    return quote(value, safe="")


def drive_item_path(path, item_id, suffix=""):
    """Address a drive item by ID or by root-relative path; suffix is /children, /content..."""
    if item_id is not None:
        return "/v1.0/me/drive/items/" + segment(item_id) + suffix
    if not path:
        return "/v1.0/me/drive/root" + suffix
    encoded_path = "/".join(segment(part) for part in path.split("/"))
    return "/v1.0/me/drive/root:/" + encoded_path + (":" + suffix if suffix else "")


def folder_path(folder):
    return "/v1.0/me/mailFolders/" + segment(folder)


def calendar_base(calendar_id):
    return "/v1.0/me/calendars/" + segment(calendar_id) if calendar_id else "/v1.0/me"


# ---------------------------------------------------------------------------
# Resources


def optional_text(item, key, limit):
    value = item.get(key) if isinstance(item, dict) else None
    if value is None:
        return None
    if not isinstance(value, str) or len(value) > limit:
        fail("protocol")
    return value


def optional_bool(item, key):
    value = item.get(key)
    if value is not None and type(value) is not bool:
        fail("protocol")
    return value


def mailbox(value):
    """Graph recipient {emailAddress: {name, address}} -> {name, address}."""
    if value is None:
        return None
    if not isinstance(value, dict) or not isinstance(value.get("emailAddress"), dict):
        fail("protocol")
    inner = value["emailAddress"]
    return {"name": optional_text(inner, "name", 1024), "address": optional_text(inner, "address", 1024)}


def mailboxes(value):
    if value is None:
        return []
    if not isinstance(value, list) or len(value) > 1000:
        fail("protocol")
    return [mailbox(item) for item in value]


def resource_id(item):
    if not isinstance(item, dict):
        fail("protocol")
    ident = item.get("id")
    if not isinstance(ident, str) or not 1 <= len(ident) <= 1024 or any(ord(c) <= 32 or ord(c) >= 127 for c in ident):
        fail("protocol")
    return ident


def normalize_message(item, max_body=None):
    """Validate one Graph message and return the public JSON shape."""
    ident = resource_id(item)
    categories = item.get("categories", [])
    if not isinstance(categories, list) or any(not isinstance(value, str) for value in categories):
        fail("protocol")
    flag = item.get("flag")
    result = {"id": ident, "conversation_id": optional_text(item, "conversationId", 1024),
              "internet_message_id": optional_text(item, "internetMessageId", 2048),
              "subject": optional_text(item, "subject", 65536), "from": mailbox(item.get("from")),
              "to": mailboxes(item.get("toRecipients")), "cc": mailboxes(item.get("ccRecipients")),
              "received_at": optional_text(item, "receivedDateTime", 64), "sent_at": optional_text(item, "sentDateTime", 64),
              "is_read": optional_bool(item, "isRead"), "is_draft": optional_bool(item, "isDraft"),
              "importance": optional_text(item, "importance", 32), "has_attachments": optional_bool(item, "hasAttachments"),
              "body_preview": optional_text(item, "bodyPreview", 4096), "web_link": optional_text(item, "webLink", 4096),
              "parent_folder_id": optional_text(item, "parentFolderId", 1024), "categories": categories[:100],
              "flag": optional_text(flag, "flagStatus", 32) if isinstance(flag, dict) else None}
    if max_body is not None:
        result["bcc"] = mailboxes(item.get("bccRecipients"))
        result["reply_to"] = mailboxes(item.get("replyTo"))
        body = item.get("body") or {}
        if not isinstance(body, dict):
            fail("protocol")
        content = optional_text(body, "content", 67108864) or ""
        result["body"] = {"content_type": (optional_text(body, "contentType", 16) or "text").lower(),
                          "content": content[:max_body], "truncated": len(content) > max_body, "length": len(content)}
        attachments = item.get("attachments", [])
        if not isinstance(attachments, list) or len(attachments) > 500:
            fail("protocol")
        result["attachments"] = [normalize_attachment(entry) for entry in attachments]
    return result


def normalize_attachment(item):
    ident = resource_id(item)
    size = item.get("size")
    if size is not None and (type(size) is not int or size < 0):
        fail("protocol")
    kind = optional_text(item, "@odata.type", 128) or ""
    return {"id": ident, "name": optional_text(item, "name", 1024), "content_type": optional_text(item, "contentType", 255),
            "size": size, "is_inline": optional_bool(item, "isInline"),
            "kind": {"#microsoft.graph.fileAttachment": "file", "#microsoft.graph.itemAttachment": "item",
                     "#microsoft.graph.referenceAttachment": "reference"}.get(kind, "unknown")}


def zoned(value):
    if value is None:
        return None
    if not isinstance(value, dict):
        fail("protocol")
    return {"date_time": optional_text(value, "dateTime", 64), "time_zone": optional_text(value, "timeZone", 128)}


def normalize_event(item):
    """Validate one Graph event and return the public JSON shape."""
    ident = resource_id(item)
    attendees = item.get("attendees", [])
    if not isinstance(attendees, list) or len(attendees) > 1000:
        fail("protocol")
    people_list = []
    for entry in attendees:
        person = mailbox(entry)
        status = entry.get("status") if isinstance(entry, dict) else None
        people_list.append({**person, "type": optional_text(entry, "type", 32),
                            "response": optional_text(status, "response", 64) if isinstance(status, dict) else None})
    location, meeting = item.get("location"), item.get("onlineMeeting")
    categories = item.get("categories", [])
    if not isinstance(categories, list) or any(not isinstance(value, str) for value in categories):
        fail("protocol")
    return {"id": ident, "subject": optional_text(item, "subject", 65536), "start": zoned(item.get("start")),
            "end": zoned(item.get("end")), "all_day": optional_bool(item, "isAllDay"),
            "location": optional_text(location, "displayName", 4096) if isinstance(location, dict) else None,
            "organizer": mailbox(item.get("organizer")), "attendees": people_list,
            "is_online_meeting": optional_bool(item, "isOnlineMeeting"),
            "online_meeting_url": optional_text(meeting, "joinUrl", 4096) if isinstance(meeting, dict) else None,
            "web_link": optional_text(item, "webLink", 4096), "show_as": optional_text(item, "showAs", 32),
            "is_cancelled": optional_bool(item, "isCancelled"), "is_organizer": optional_bool(item, "isOrganizer"),
            "body_preview": optional_text(item, "bodyPreview", 4096), "categories": categories[:100],
            "type": optional_text(item, "type", 32), "series_master_id": optional_text(item, "seriesMasterId", 1024)}


def normalize_item(item):
    """Validate one driveItem and return the public JSON shape (never the pre-authenticated download URL)."""
    ident = resource_id(item)
    name = optional_text(item, "name", 1024)
    size = item.get("size")
    if size is not None and (type(size) is not int or size < 0):
        fail("protocol")
    file_facet, folder_facet = item.get("file"), item.get("folder")
    if (file_facet is not None and not isinstance(file_facet, dict)) or (folder_facet is not None and not isinstance(folder_facet, dict)):
        fail("protocol")
    hashes = (file_facet or {}).get("hashes") or {}
    if not isinstance(hashes, dict):
        fail("protocol")
    sha256 = optional_text(hashes, "sha256Hash", 64)
    if sha256 is not None and not SHA256_HEX.fullmatch(sha256):
        fail("protocol")
    parent = item.get("parentReference") or {}
    if not isinstance(parent, dict):
        fail("protocol")
    count = (folder_facet or {}).get("childCount")
    if count is not None and (type(count) is not int or count < 0):
        fail("protocol")
    return {"id": ident, "name": name, "is_folder": folder_facet is not None, "size": size,
            "mime_type": optional_text(file_facet, "mimeType", 255) if file_facet else None, "child_count": count,
            "created_at": optional_text(item, "createdDateTime", 64), "modified_at": optional_text(item, "lastModifiedDateTime", 64),
            "web_url": optional_text(item, "webUrl", 4096), "parent_id": optional_text(parent, "id", 1024),
            "parent_path": optional_text(parent, "path", 4096), "drive_id": optional_text(parent, "driveId", 1024),
            "sha256": sha256.lower() if sha256 else None, "quick_xor_hash": optional_text(hashes, "quickXorHash", 64)}


def normalize_channel_message(item, team, channel):
    ident = resource_id(item)
    return {"id": ident, "team_id": team, "channel_id": channel, "created_at": optional_text(item, "createdDateTime", 64),
            "web_url": optional_text(item, "webUrl", 4096)}


def collection(body, page_size):
    """Validate one OData collection page: (values, nextLink, deltaLink)."""
    if not isinstance(body, dict):
        fail("protocol")
    values = body.get("value", [])
    following, delta = body.get("@odata.nextLink"), body.get("@odata.deltaLink")
    if (not isinstance(values, list) or len(values) > page_size
            or any(item is not None and not isinstance(item, str) for item in (following, delta))):
        fail("protocol")
    return values, following, delta


def local_name(name, extension=""):
    """Sanitize a provider name into one safe local file name, keeping its extension."""
    text = "".join("_" if ord(char) < 32 or ord(char) == 127 or char in '/\\:*?"<>|' else char for char in name or "")
    text = text.strip().lstrip(".").strip() or "download"
    if extension and not text.lower().endswith(extension):
        text += extension
    stem, dot, suffix = text.rpartition(".")
    if not dot or len(suffix) > 16 or not stem:
        stem, suffix = text, ""
    suffix = "." + suffix if suffix else ""
    while len((stem + suffix).encode("utf-8")) > LOCAL_NAME_BYTES and stem:
        stem = stem[:-1]
    return (stem.rstrip() or "download") + suffix


# ---------------------------------------------------------------------------
# Mail watch state


def watch_key(config):
    """One cursor per tenant, wallet reference and folder; the document never holds the reference itself."""
    material = "|".join((config["tenant_id"], config["secret_ref"], config["mail_folder"]))
    return hashlib.sha256(material.encode("utf-8")).hexdigest()[:24]


def seen_digest(message_id):
    return hashlib.sha256(message_id.encode("utf-8", errors="surrogatepass")).hexdigest()[:16]


def state_document(value):
    """Validate a stored watch document: links, seen digests and timestamps only."""
    if not isinstance(value, dict) or value.get("schema_version") != 1 or set(value) != {"schema_version", "watches"}:
        fail("state_invalid")
    watches = value["watches"]
    if not isinstance(watches, dict) or len(watches) > 8:
        fail("state_invalid")
    for key, entry in watches.items():
        if (not isinstance(key, str) or not re.fullmatch(r"[0-9a-f]{24}", key) or not isinstance(entry, dict)
                or set(entry) != {"link", "link_kind", "since", "last_received", "seen", "updated_at"}):
            fail("state_invalid")
        if entry["link_kind"] not in ("next", "delta", "none") or (entry["link"] is None) != (entry["link_kind"] == "none"):
            fail("state_invalid")
        if entry["link"] is not None and (not isinstance(entry["link"], str) or len(entry["link"]) > 16384):
            fail("state_invalid")
        for stamp in ("since", "last_received"):
            if entry[stamp] is not None and (not isinstance(entry[stamp], str) or not ZONED_TIME.fullmatch(entry[stamp])):
                fail("state_invalid")
        seen = entry["seen"]
        if (not isinstance(seen, list) or len(seen) > SEEN_LIMIT
                or any(not isinstance(item, str) or not re.fullmatch(r"[0-9a-f]{16}", item) for item in seen)):
            fail("state_invalid")
        if type(entry["updated_at"]) is not int or entry["updated_at"] < 0:
            fail("state_invalid")
    return value


def empty_state():
    return {"schema_version": 1, "watches": {}}


def with_watch(document, key, entry):
    """Return a new document holding this watch; at most eight watches are retained."""
    watches = dict(document["watches"])
    watches[key] = entry
    if len(watches) > 8:
        for old in sorted(watches, key=lambda item: watches[item]["updated_at"])[:len(watches) - 8]:
            if old != key:
                watches.pop(old)
    return {"schema_version": 1, "watches": watches}


# ---------------------------------------------------------------------------
# Provider responses


def map_status(status):
    return {400: "bad_request", 401: "unauthorized", 403: "forbidden", 404: "not_found", 409: "conflict",
            410: "sync_reset", 429: "rate_limited"}.get(status, "provider_unavailable" if status >= 500 else "provider_error")


QUOTA_CODES = frozenset(("quotaLimitReached", "ErrorQuotaExceeded", "insufficientStorage", "ErrorMailboxStoreUnavailable"))


def provider_failure(status, body, secrets, *, mutation, retry_after=None):
    """Raise the coded error for a non-2xx Graph response ({"error": {code, message, innerError}})."""
    detail = body.get("error") if isinstance(body, dict) else None
    detail = detail if isinstance(detail, dict) else {}
    provider_code = clean(detail.get("code"), 64)
    message = clean(redact(detail.get("message"), secrets), 200)
    inner = detail.get("innerError") or detail.get("innererror")
    request_id = clean(inner.get("request-id"), 64) if isinstance(inner, dict) else None
    code = map_status(status)
    if provider_code in QUOTA_CODES or status == 507:
        code = "storage_quota"
    if provider_code in SYNC_RESET:
        code = "sync_reset"
    uncertain = mutation and code == "provider_unavailable"
    fail(code, MESSAGES["provider_unavailable_uncertain"] if uncertain else None, uncertain=uncertain, http_status=status,
         retry_after=retry_after if code in ("rate_limited", "provider_unavailable") else None,
         provider_code=provider_code, provider_message=message, provider_request_id=request_id)
