"""New-mail polling through the Graph messages delta query, bounded per poll.

A round starts with `$filter=receivedDateTime ge <since>`: the current time on
first start (nothing older is replayed), or `existing_since_hours` earlier when
emit_existing is set. Graph then returns nextLink pages until a deltaLink, which
later polls follow to receive only changes. Messages already handed over are
recognized by a digest of their ID, so read-flag updates and moves are not new
mail. Nothing is persisted here: the caller stores the returned document before
it hands the messages over, so a restart never replays them.
"""

from datetime import timedelta

from . import logic


def poll(session, config, document, now):
    """One bounded poll; returns (new document, new messages, summary)."""
    key = logic.watch_key(config)
    entry = document["watches"].get(key)
    page_size = min(config["page_size"], config["watch_max_items"])
    headers = {"Prefer": "odata.maxpagesize=" + str(page_size)}
    started = entry is None
    if entry is None:
        since = now - timedelta(hours=config["existing_since_hours"]) if config["emit_existing"] else now
        entry = {"link": None, "link_kind": "none", "since": logic.utc_text(since), "last_received": None, "seen": [],
                 "updated_at": int(now.timestamp())}
    since, last = entry["since"], entry["last_received"]
    link, kind = entry["link"], entry["link_kind"]
    seen = list(entry["seen"])
    known = set(seen)
    found, pages, skipped, resynced = [], 0, 0, False
    while True:
        target = logic.api_link(link, config, logic.WATCH_LINK) if kind != "none" else None
        if kind != "none" and target is None:
            # A cursor for another API origin (settings changed) restarts from the last position.
            kind, since, resynced = "none", last or since, True
        try:
            if kind == "none":
                body = session.call("GET", logic.folder_path(config["mail_folder"]) + "/messages/delta",
                                    {"$select": logic.MESSAGE_SELECT, "$filter": "receivedDateTime ge " + since},
                                    mutation=False, headers=headers)
            else:
                body = session.follow(target, headers=headers)
        except logic.GraphError as exc:
            if exc.code == "sync_reset" and not resynced:
                kind, since, resynced = "none", last or since, True
                continue
            raise
        pages += 1
        values, following, delta = logic.collection(body, page_size)
        for item in values:
            if not isinstance(item, dict):
                logic.fail("protocol")
            if "@removed" in item:
                skipped += 1
                continue
            message = logic.normalize_message(item)
            digest = logic.seen_digest(message["id"])
            if digest in known:
                skipped += 1
                continue
            known.add(digest)
            seen.append(digest)
            found.append(message)
            received = message["received_at"]
            if received:
                try:
                    stamp = logic.utc_text(logic.parse_time(received))
                except ValueError:
                    logic.fail("protocol")
                last = stamp if last is None or stamp > last else last
        if delta is not None:
            if logic.api_link(delta, config, logic.WATCH_LINK) is None:
                logic.fail("protocol")
            link, kind = delta, "delta"
            break
        if following is None or logic.api_link(following, config, logic.WATCH_LINK) is None:
            logic.fail("protocol")
        link, kind = following, "next"
        if pages >= config["max_pages"] or len(found) + page_size > config["watch_max_items"]:
            break
    found.sort(key=lambda message: (message["received_at"] or "", message["id"]))
    updated = {"link": link, "link_kind": kind, "since": since, "last_received": last, "seen": seen[-logic.SEEN_LIMIT:],
               "updated_at": int(now.timestamp())}
    summary = {"folder": config["mail_folder"], "count": len(found), "pages": pages, "skipped": skipped,
               "started": started, "resynced": resynced, "caught_up": kind == "delta"}
    return logic.with_watch(document, key, updated), found, summary
