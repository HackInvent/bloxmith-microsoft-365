"""The Microsoft 365 operations, built from Session calls, local file helpers and pure logic.

Every operation validates what it can before its first request and sends at
most one change. Listings follow Graph nextLinks bounded by page_size,
max_pages, the request limit and the output size; a page is never cut, so
next_page_link always resumes exactly where the result stopped.
"""

import base64
from datetime import datetime, timezone
import json
import uuid

from . import files, logic, storage, transport, watch


def identity(request):
    return {"request_id": request["request_id"], "operation": request["operation"]}


def listing(session, request, config, target, query, normalize, pattern, headers=None):
    """Pages of normalized resources, never more than the request limit and never a cut page."""
    limit = request["limit"]
    top = min(config["page_size"], limit)
    user_link = request.get("page_link") is not None
    bound = 1000 if user_link else top
    found, pages, weight_total = [], 0, 2
    pending = request.get("page_link")  # absolute URL of the next page; None before the first query
    stop = "end"
    while True:
        if pages >= config["max_pages"]:
            stop = "max_pages"
            break
        if pending is None:
            body = session.call("GET", target, {**query, "$top": top}, mutation=False, headers=headers)
        else:
            body = session.follow(logic.api_link(pending, config, pattern), headers=headers)
        pages += 1
        values, following, _delta = logic.collection(body, bound)
        if len(found) + len(values) > limit:
            logic.fail("invalid_input", "This page_link returns pages larger than limit; send the limit of the original listing.")
        page = [normalize(item) for item in values]
        weight = sum(len(json.dumps(item, ensure_ascii=False, separators=(",", ":")).encode("utf-8")) + 1 for item in page)
        if weight_total + weight > logic.OUTPUT_LIMIT:
            if not found:
                logic.fail("protocol", "One Graph page exceeds the 16 MiB output limit; lower page_size.")
            # The page that does not fit is dropped whole; its own link resumes it.
            stop = "output_limit"
            break
        found.extend(page)
        weight_total += weight
        if following is None:
            pending, stop = None, "end"
            break
        if logic.api_link(following, config, pattern) is None:
            logic.fail("protocol")
        pending = following
        if len(found) + max(top, len(values)) > limit:
            stop = "limit"
            break
    return found, {**identity(request), "count": len(found), "truncated": stop != "end", "stop_reason": stop,
                   "next_page_link": pending if stop != "end" else None, "pages": pages}


def zone_header(config):
    return {"Prefer": 'outlook.timezone="' + config["time_zone"] + '"'}


def mail_list(session, request, config, **_):
    query = {}
    if request.get("page_link") is None:
        query = {"$select": request["select"]}
        if request["filter"]:
            query["$filter"] = request["filter"]
        if request["search"]:
            query["$search"] = '"' + request["search"] + '"'
        if request["order"]:
            query["$orderby"] = request["order"]
    target = logic.folder_path(request["folder"]) + "/messages" if request.get("page_link") is None else None
    items, result = listing(session, request, config, target, query, logic.normalize_message,
                            logic.LINK_PATTERNS["mail_list"])
    if request.get("page_link") is None:
        result["folder"] = request["folder"]
    result["requests"] = session.requests
    return {"items": items, "result": result}


def mail_get(session, request, config, **_):
    body_type = "html" if request["body_format"] == "html" else "text"
    body = session.call("GET", "/v1.0/me/messages/" + logic.segment(request["message_id"]),
                        {"$select": logic.MESSAGE_SELECT + ",sender,bccRecipients,replyTo,body",
                         "$expand": "attachments($select=id,name,contentType,size,isInline)"},
                        mutation=False, headers={"Prefer": 'outlook.body-content-type="' + body_type + '"'})
    message = logic.normalize_message(body, max_body=config["max_body_chars"])
    return {"result": {**identity(request), "message": message, "requests": session.requests}}


def recipients(entries):
    return [{"emailAddress": {"address": item["address"], **({"name": item["name"]} if item["name"] else {})}}
            for item in entries]


def mail_send(session, request, config, attachments=(), **_):
    """sendMail: one request, accepted with HTTP 202; never replayed once it may have been processed."""
    message = {"subject": request["subject"], "importance": request["importance"],
               "body": {"contentType": request["body_format"], "content": request["body"]},
               "toRecipients": recipients(request["to"]), "ccRecipients": recipients(request["cc"]),
               "bccRecipients": recipients(request["bcc"])}
    if attachments:
        message["attachments"] = [{"@odata.type": logic.FILE_ATTACHMENT, "name": name,
                                   "contentType": logic.guessed_type(name),
                                   "contentBytes": base64.b64encode(data).decode("ascii")} for name, data in attachments]
    payload = {"message": message, "saveToSentItems": request["save_to_sent_items"]}
    logic.encoded(payload, logic.SEND_REQUEST_LIMIT, "request_too_large")
    session.call("POST", "/v1.0/me/sendMail", payload=payload, mutation=True, empty=True,
                 payload_limit=logic.SEND_REQUEST_LIMIT)
    return {"result": {**identity(request), "accepted": True, "to": [item["address"] for item in request["to"]],
                       "cc": [item["address"] for item in request["cc"]], "bcc_count": len(request["bcc"]),
                       "attachments": [name for name, _data in attachments],
                       "saved_to_sent_items": request["save_to_sent_items"], "requests": session.requests}}


def mail_move(session, request, config, **_):
    body = session.call("POST", "/v1.0/me/messages/" + logic.segment(request["message_id"]) + "/move",
                        payload={"destinationId": request["destination"]}, mutation=True)
    moved = logic.normalize_message(body)
    return {"result": {**identity(request), "message_id": request["message_id"], "destination": request["destination"],
                       "message": moved, "requests": session.requests}}


def mail_mark_read(session, request, config, **_):
    body = session.call("PATCH", "/v1.0/me/messages/" + logic.segment(request["message_id"]),
                        payload={"isRead": request["is_read"]}, mutation=True)
    message = logic.normalize_message(body)
    if message["is_read"] is not None and message["is_read"] != request["is_read"]:
        logic.fail("protocol", "Graph answered with a different read state; the change is uncertain.", uncertain=True)
    return {"result": {**identity(request), "message": message, "is_read": request["is_read"],
                       "requests": session.requests}}


def mail_attachment_download(session, request, config, storage_dir=None, **_):
    directory = files.storage_directory(storage_dir)
    path = ("/v1.0/me/messages/" + logic.segment(request["message_id"]) + "/attachments/"
            + logic.segment(request["attachment_id"]))
    meta = session.call("GET", path, {"$select": "id,name,contentType,size,isInline"}, mutation=False)
    attachment = logic.normalize_attachment(meta)
    if attachment["kind"] != "file":
        logic.fail("not_downloadable")
    if attachment["size"] is not None and attachment["size"] > config["max_download_bytes"] + 65536:
        logic.fail("file_too_large", "The attachment is larger than max_download_bytes; nothing was downloaded.")
    target = files.reserve(directory)
    try:
        sink = transport.Sink(target["fd"], config["max_download_bytes"])
        received = session.call("GET", path + "/$value", mutation=False, sink=sink)
        final = files.finish(target, logic.local_name(attachment["name"]), config["max_kept_downloads"])
    except BaseException:
        files.discard(target)
        raise
    result = {**identity(request), "message_id": request["message_id"], "attachment_id": attachment["id"],
              "name": attachment["name"], "content_type": attachment["content_type"], "path": str(final),
              "size": received["size"], "sha256": received["sha256"], "requests": session.requests}
    return {"file": str(final), "result": result}


def mail_watch(session, request, config, storage_dir=None, now=None, **_):
    """One Shot (or a request) runs one bounded poll; the listener runs the same poll on its interval."""
    store = storage.WatchStore(files.storage_directory(storage_dir))
    if not store.acquire():
        logic.fail("busy")
    try:
        document = store.load()
        updated, messages, summary = watch.poll(session, config, document, now or datetime.now(timezone.utc))
        store.save(updated)
    finally:
        store.close()
    return {"items": messages, "result": {**identity(request), **summary, "source": "request",
                                          "requests": session.requests}}


def events_list(session, request, config, **_):
    query, target = {}, None
    if request.get("page_link") is None:
        target = logic.calendar_base(request["calendar_id"]) + "/calendarView"
        query = {"startDateTime": request["start"], "endDateTime": request["end"], "$select": logic.EVENT_SELECT,
                 "$orderby": "start/dateTime"}
    items, result = listing(session, request, config, target, query, logic.normalize_event,
                            logic.LINK_PATTERNS["events_list"], headers=zone_header(config))
    if request.get("page_link") is None:
        result.update(start=request["start"], end=request["end"], calendar_id=request["calendar_id"] or None)
    result["time_zone"] = config["time_zone"]
    result["requests"] = session.requests
    return {"items": items, "result": result}


def meeting_gate(session, request, config, action):
    """Changing or deleting a meeting you organize notifies its attendees: that counts as sending."""
    body = session.call("GET", "/v1.0/me/events/" + logic.segment(request["event_id"]),
                        {"$select": "id,isOrganizer,attendees"}, mutation=False)
    event = logic.normalize_event(body)
    invited = [person["address"] for person in event["attendees"] if person.get("address")]
    notifies = bool(event["is_organizer"] and invited)
    if notifies:
        if not config["allow_send"]:
            logic.fail("send_disabled", f"This meeting has attendees who would be notified of the {action}; "
                       "enable allow_send to allow it. Nothing was changed.")
        if any(not logic.recipient_allowed(address, config["allowed_recipients"]) for address in invited):
            logic.fail("recipient_not_allowed", f"An attendee of this meeting is not in allowed_recipients; the {action} "
                       "would notify them. Nothing was changed.")
    return notifies


def event_create(session, request, config, **_):
    payload = {**request["fields"], "transactionId": str(uuid.uuid4())}
    body = session.call("POST", logic.calendar_base(request["calendar_id"]) + "/events", payload=payload, mutation=True,
                        headers=zone_header(config))
    event = logic.normalize_event(body)
    return {"result": {**identity(request), "event": event, "invitations_sent": bool(request.get("attendees")),
                       "requests": session.requests}}


def event_update(session, request, config, **_):
    notifies = meeting_gate(session, request, config, "update")
    body = session.call("PATCH", "/v1.0/me/events/" + logic.segment(request["event_id"]), payload=request["fields"],
                        mutation=True, headers=zone_header(config))
    event = logic.normalize_event(body)
    return {"result": {**identity(request), "event": event,
                       "updates_sent": bool(notifies or request.get("attendees")), "requests": session.requests}}


def event_delete(session, request, config, **_):
    """DELETE moves the event to Deleted Items (recoverable); permanentDelete is never used."""
    notifies = meeting_gate(session, request, config, "cancellation")
    session.call("DELETE", "/v1.0/me/events/" + logic.segment(request["event_id"]), mutation=True, empty=True)
    return {"result": {**identity(request), "event_id": request["event_id"], "deleted": True, "recoverable": True,
                       "cancellations_sent": notifies, "requests": session.requests}}


def files_list(session, request, config, **_):
    target, query = None, {}
    if request.get("page_link") is None:
        target = logic.drive_item_path(request["path"], request["item_id"], "/children")
        query = {"$select": logic.DRIVE_SELECT}
    items, result = listing(session, request, config, target, query, logic.normalize_item,
                            logic.LINK_PATTERNS["files_list"])
    if request.get("page_link") is None:
        result.update(path=request["path"], item_id=request["item_id"])
    result["requests"] = session.requests
    return {"items": items, "result": result}


def file_download(session, request, config, storage_dir=None, **_):
    """Read the item, then stream its pre-authenticated download URL (no token) into node storage."""
    directory = files.storage_directory(storage_dir)
    meta = session.call("GET", logic.drive_item_path(request["path"], request["item_id"]), mutation=False)
    item = logic.normalize_item(meta)
    if item["is_folder"] or meta.get("file") is None:
        logic.fail("not_downloadable")
    if item["size"] is not None and item["size"] > config["max_download_bytes"]:
        logic.fail("file_too_large", "The OneDrive file is larger than max_download_bytes; nothing was downloaded.")
    link = logic.transfer_link(meta.get("@microsoft.graph.downloadUrl"), config)
    if link is None:
        logic.fail("protocol", "Graph returned no usable download URL for this file.")
    target = files.reserve(directory)
    try:
        sink = transport.Sink(target["fd"], config["max_download_bytes"])
        received = session.transfer(link[0], "GET", link[1], mutation=False, sink=sink)
        if item["size"] is not None and received["size"] != item["size"]:
            logic.fail("protocol", "The downloaded size differs from the OneDrive metadata; the partial file was removed.")
        verified = None
        if item["sha256"]:
            if received["sha256"] != item["sha256"]:
                logic.fail("protocol", "The downloaded SHA-256 differs from the OneDrive metadata; the file was removed.")
            verified = True
        final = files.finish(target, logic.local_name(item["name"]), config["max_kept_downloads"])
    except BaseException:
        files.discard(target)
        raise
    result = {**identity(request), "item_id": item["id"], "name": item["name"], "mime_type": item["mime_type"],
              "path": str(final), "size": received["size"], "sha256": received["sha256"], "sha256_verified": verified,
              "modified_at": item["modified_at"], "requests": session.requests}
    return {"file": str(final), "result": result}


def certain(error):
    """The same failure without uncertainty: nothing was committed."""
    base = {"network": "network", "timeout": "timeout", "cancelled": "cancelled_before_send",
            "provider_unavailable": "provider_unavailable", "protocol": "protocol"}.get(error.code)
    fresh = logic.GraphError(error.code, logic.MESSAGES[base] if base else str(error), http_status=error.http_status,
                             retry_after=error.retry_after, provider_code=error.provider_code,
                             provider_message=error.provider_message, provider_request_id=error.provider_request_id)
    fresh.attempts = error.attempts
    return fresh


def file_upload(session, request, config, upload=None, **_):
    """Simple PUT up to 4 MB, otherwise an upload session in 320 KiB-aligned chunks without the token."""
    local = upload
    name = request.get("name") or logic.drive_name(local.name, "The local file name")
    remote = (request["folder"] + "/" + name) if request["folder"] else name
    behavior = config["conflict_behavior"]
    digest = local.sha256()
    if local.size <= logic.SIMPLE_UPLOAD_LIMIT:
        data = local.read_all()
        target = logic.drive_item_path(remote, None, "/content") + "?@microsoft.graph.conflictBehavior=" + behavior
        created = session.call("PUT", target, raw=data, content_type=logic.guessed_type(name), mutation=True)
        method, chunks = "simple", 1
    else:
        opened = session.call("POST", logic.drive_item_path(remote, None, "/createUploadSession"),
                              payload={"item": {"@microsoft.graph.conflictBehavior": behavior, "name": name}},
                              mutation=False)
        link = logic.transfer_link(opened.get("uploadUrl") if isinstance(opened, dict) else None, config)
        if link is None:
            logic.fail("protocol", "Graph returned no usable upload session URL; nothing was uploaded.")
        offset, chunks, final_sent, created = 0, 0, False, None
        try:
            while offset < local.size:
                size = min(config["upload_chunk_bytes"], local.size - offset)
                data = local.read(offset, size)
                last = offset + size == local.size
                if last:
                    local.check()
                    final_sent = True
                body = session.transfer(link[0], "PUT", link[1], body=data, mutation=True, headers={
                    "Content-Range": f"bytes {offset}-{offset + size - 1}/{local.size}",
                    "Content-Type": "application/octet-stream"})
                chunks += 1
                offset += size
                if not last:
                    ranges = body.get("nextExpectedRanges")
                    if (not isinstance(ranges, list) or not ranges or not isinstance(ranges[0], str)
                            or ranges[0].split("-", 1)[0] != str(offset)):
                        logic.fail("protocol", "The upload session expects another byte range; nothing was committed.")
                else:
                    created = body
        except logic.GraphError as exc:
            still_open = session.discard_session(*link)
            if not final_sent or still_open:
                raise certain(exc) from None
            raise
        method = "session"
    item = logic.normalize_item(created)
    if item["sha256"] and item["sha256"] != digest:
        logic.fail("protocol", "OneDrive stored the upload with a different SHA-256; check the new item " + item["id"] + ".",
                   uncertain=True)
    return {"result": {**identity(request), "item": item, "path": remote, "size": local.size, "method": method,
                       "chunks": chunks, "conflict_behavior": behavior,
                       "sha256_verified": True if item["sha256"] else None, "requests": session.requests}}


def folder_create(session, request, config, **_):
    """Create a folder; with folder_exists=reuse an existing folder of that name is returned instead."""
    parent = request["folder"]
    try:
        body = session.call("POST", logic.drive_item_path(parent, None, "/children"), mutation=True,
                            payload={"name": request["name"], "folder": {}, "@microsoft.graph.conflictBehavior": "fail"})
        return {"result": {**identity(request), "folder": logic.normalize_item(body), "created": True, "parent": parent,
                           "requests": session.requests}}
    except logic.GraphError as exc:
        if exc.code != "conflict" or config["folder_exists"] != "reuse":
            raise
    path = (parent + "/" + request["name"]) if parent else request["name"]
    existing = logic.normalize_item(session.call("GET", logic.drive_item_path(path, None), {"$select": logic.DRIVE_SELECT},
                                                 mutation=False))
    if not existing["is_folder"]:
        logic.fail("conflict", "A file with this name already exists in the parent folder; nothing was created.")
    return {"result": {**identity(request), "folder": existing, "created": False, "parent": parent,
                       "requests": session.requests}}


def channel_message_send(session, request, config, **_):
    payload = {"body": {"contentType": request["content_type"], "content": request["content"]}}
    if request["subject"]:
        payload["subject"] = request["subject"]
    body = session.call("POST", "/v1.0/teams/" + logic.segment(request["team_id"]) + "/channels/"
                        + logic.segment(request["channel_id"]) + "/messages", payload=payload, mutation=True)
    message = logic.normalize_channel_message(body, request["team_id"], request["channel_id"])
    return {"result": {**identity(request), "message": message, "requests": session.requests}}


RUNNERS = {name: globals()[name] for name in logic.OPERATIONS}


def run(session, request, config, **context):
    return RUNNERS[request["operation"]](session, request, config, **context)
