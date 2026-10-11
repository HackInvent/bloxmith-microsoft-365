"""Bounded, cancellable HTTPS exchanges with Microsoft Graph and the Microsoft identity platform.

Standard library only. Each exchange runs in a helper thread so cancellation,
Stop and the activation deadline can shut the socket down at once. Reads are
retried after HTTP 429, 5xx or a lost connection; changes only after HTTP 429,
which Graph did not process. Ambiguous changes are reported, never replayed.
Pre-authenticated download and upload-session URLs never receive the token.
"""

import hashlib
import http.client
import json
import os
import socket
import ssl
import threading
import time
from urllib.parse import quote, urlencode, urlsplit

from . import logic

CHUNK = 65536
USER_AGENT = "BloxSmith-Microsoft365/0.1.0"
THREAD_NAME = "microsoft-365-http"


class Failure(Exception):
    """A content-free transport outcome: code and the phase it happened in."""

    def __init__(self, code, phase):
        super().__init__(code)
        self.code, self.phase = code, phase


class TooLarge(Exception):
    """A streamed body passed its byte limit."""


class Sink:
    """One exclusive download file: written bytes, SHA-256 and the byte limit, reset before each attempt."""

    def __init__(self, fd, limit):
        self.fd, self.limit = fd, limit
        self.reset()

    def reset(self):
        os.ftruncate(self.fd, 0)
        os.lseek(self.fd, 0, os.SEEK_SET)
        self.size, self.digest = 0, hashlib.sha256()

    def write(self, data):
        if self.size + len(data) > self.limit:
            raise TooLarge()
        view = memoryview(data)
        while view:
            written = os.write(self.fd, view)
            view = view[written:]
        self.size += len(data)
        self.digest.update(data)

    def sha256(self):
        return self.digest.hexdigest()


def reject_constant(_value):
    raise ValueError("Non-finite JSON number.")


def remaining(deadline, phase):
    left = deadline - time.monotonic()
    if left <= 0:
        raise Failure("timeout", phase)
    return max(0.001, left)


def read_json(response, conn, deadline, limit, found):
    """Read a bounded JSON body and report why it cannot be used, never raising on content."""
    chunks, total = [], 0
    while total <= limit:
        if conn.sock is not None:
            conn.sock.settimeout(remaining(deadline, "receive"))
        piece = response.read1(min(CHUNK, limit + 1 - total))
        if not piece:
            break
        total += len(piece)
        chunks.append(piece)
    media = found.get("content-type", "").split(";", 1)[0].strip().lower()
    if total > limit:
        return None, "too_large"
    if found.get("content-encoding", "identity").strip().lower() not in ("identity", ""):
        return None, "encoded"
    if total == 0:
        return None, "empty"
    if media != "application/json":
        return None, "not_json"
    try:
        return json.loads(b"".join(chunks).decode("utf-8", errors="strict"), parse_constant=reject_constant), None
    except (ValueError, UnicodeError, RecursionError):
        return None, "invalid_json"


def stream(response, conn, deadline, sink, found):
    """Copy a successful media body into the sink; the declared length is checked before reading."""
    if found.get("content-encoding", "identity").strip().lower() not in ("identity", ""):
        return "encoded"
    declared = found.get("content-length", "").strip()
    if declared.isdigit() and int(declared) > sink.limit:
        return "too_large"
    sink.reset()
    try:
        while True:
            if conn.sock is not None:
                conn.sock.settimeout(remaining(deadline, "receive"))
            piece = response.read1(CHUNK)
            if not piece:
                break
            sink.write(piece)
    except TooLarge:
        return "too_large"
    if declared.isdigit() and int(declared) != sink.size:
        raise Failure("network", "receive")
    return None


def retry_after(found):
    """Retry-After in seconds (Graph sends delta-seconds); dates and huge values are ignored."""
    value = found.get("retry-after", "").strip()
    return int(value) if value.isdigit() and len(value) <= 6 else None


def perform(state, origin, ca_file, method, target, headers, body, deadline, limit, sink):
    """One request/response without redirects, cookies, proxies or decompression."""
    parsed = urlsplit(origin)
    host, port = parsed.hostname, parsed.port
    try:
        if parsed.scheme == "https":
            context = ssl.create_default_context(cafile=ca_file or None)
            conn = http.client.HTTPSConnection(host, port or 443, timeout=remaining(deadline, "connect"), context=context)
        else:
            conn = http.client.HTTPConnection(host, port or 80, timeout=remaining(deadline, "connect"))
        state["conn"] = conn
        conn.connect()
        if state.get("aborted"):
            raise Failure("network", "connect")
        state["phase"] = "send"
        conn.putrequest(method, target, skip_host=False, skip_accept_encoding=True)
        for key, value in headers.items():
            conn.putheader(key, value)
        conn.putheader("Content-Length", str(len(body or b"")))
        conn.sock.settimeout(remaining(deadline, "send"))
        conn.endheaders(message_body=body or None)
        state["phase"] = "receive"
        conn.sock.settimeout(remaining(deadline, "receive"))
        response = conn.getresponse()
        found = {key.lower(): value for key, value in response.getheaders()}
        result = {"status": response.status, "retry_after": retry_after(found), "body": None, "body_error": None,
                  "streamed": False}
        if sink is not None and 200 <= response.status < 300:
            result["streamed"] = True
            result["body_error"] = stream(response, conn, deadline, sink, found)
        else:
            result["body"], result["body_error"] = read_json(response, conn, deadline, limit, found)
        return result
    except Failure:
        raise
    except ssl.SSLCertVerificationError:
        raise Failure("tls", state["phase"]) from None
    except ssl.SSLError:
        raise Failure("tls" if state["phase"] == "connect" else "network", state["phase"]) from None
    except (socket.timeout, TimeoutError):
        raise Failure("timeout", state["phase"]) from None
    except (OSError, http.client.HTTPException, ValueError, AttributeError):
        raise Failure("network", state["phase"]) from None
    finally:
        conn = state.get("conn")
        if conn is not None:
            conn.close()


def abort(state):
    """Shut the socket down from the supervising thread; the helper thread then fails fast."""
    state["aborted"] = True
    conn = state.get("conn")
    sock = getattr(conn, "sock", None) if conn is not None else None
    if sock is not None:
        try:
            sock.shutdown(socket.SHUT_RDWR)
        except OSError:
            pass


def exchange(origin, ca_file, method, target, headers, body, *, deadline, cancel, limit, sink=None):
    """Run one request in a helper thread and supervise cancellation and the deadline."""
    state, box = {"phase": "connect", "conn": None}, {}

    def worker_target():
        try:
            box["result"] = perform(state, origin, ca_file, method, target, headers, body, deadline, limit, sink)
        except Failure as failure:
            box["failure"] = failure
        except Exception:  # Defensive: an unexpected library error is a transport failure.
            box["failure"] = Failure("network", state["phase"])

    worker = threading.Thread(target=worker_target, name=THREAD_NAME, daemon=True)
    worker.start()
    while worker.is_alive():
        worker.join(0.05)
        if not worker.is_alive():
            break
        if cancel():
            phase = state["phase"]
            abort(state)
            worker.join(2)
            raise Failure("cancelled", phase)
        if time.monotonic() >= deadline + 0.5:
            phase = state["phase"]
            abort(state)
            worker.join(2)
            raise Failure("timeout", phase)
    if "failure" in box:
        raise box["failure"]
    return box["result"]


def secret_fingerprint(secret):
    return hashlib.sha256(json.dumps(secret, sort_keys=True).encode("utf-8")).hexdigest()


class Session:
    """One activation (or one watch poll): one deadline, one credential and bounded requests.

    token_cache, when given, is a caller-owned dict that keeps the access token in
    memory between polls of one listener; it is never written anywhere.
    """

    def __init__(self, config, secret, cancel=lambda: False, token_cache=None):
        self.config, self.secret = config, dict(secret)
        self.cancel = cancel if callable(cancel) else (lambda: False)
        self.deadline = time.monotonic() + config["timeout_sec"]
        self.secrets = [value for value in self.secret.values() if isinstance(value, str)]
        self.cache = token_cache if isinstance(token_cache, dict) else {}
        self.token = None
        self.requests = 0
        cached = self.cache.get("token")
        if (cached and self.cache.get("fingerprint") == secret_fingerprint(self.secret)
                and self.cache.get("expires_at", 0) > time.monotonic() + 60):
            self.token = cached
            self.secrets.append(cached)

    def pause(self, seconds):
        """Cancellation-aware wait between a retryable response and its retry."""
        end = time.monotonic() + seconds
        while time.monotonic() < end:
            if self.cancel():
                logic.fail("cancelled", logic.MESSAGES["cancelled_before_send"])
            time.sleep(min(0.05, max(0.0, end - time.monotonic())))

    def send(self, origin, method, target, headers, body, *, mutation, sink=None):
        """One exchange mapped to a structured outcome; never raises a raw library error."""
        if self.cancel():
            logic.fail("cancelled", logic.MESSAGES["cancelled_before_send"])
        if self.deadline - time.monotonic() <= 0.2:
            logic.fail("timeout")
        self.requests += 1
        try:
            return exchange(origin, self.config["ca_file"], method, target, headers, body, deadline=self.deadline,
                            cancel=self.cancel, limit=self.config["max_response_bytes"], sink=sink)
        except Failure as failure:
            uncertain = mutation and failure.phase == "receive"
            if failure.code == "tls":
                logic.fail("tls")
            if failure.code == "timeout":
                logic.fail("timeout", logic.MESSAGES["timeout_uncertain" if uncertain else "timeout"], uncertain=uncertain)
            if failure.code == "cancelled":
                logic.fail("cancelled", logic.MESSAGES["cancelled" if uncertain else "cancelled_before_send"],
                           uncertain=uncertain)
            logic.fail("network", logic.MESSAGES["network_uncertain" if uncertain else "network"], uncertain=uncertain,
                       retryable=not mutation)

    def retrying(self, attempt, *, mutation):
        """Run attempt() with bounded retries: throttling always; 5xx and lost connections for reads only."""
        attempts = 0
        while True:
            attempts += 1
            try:
                return attempt()
            except logic.GraphError as exc:
                exc.attempts = attempts
                retryable = exc.code == "rate_limited" or (not mutation and (
                    exc.code == "provider_unavailable" or getattr(exc, "retryable", False)))
                wait = exc.retry_after if exc.retry_after is not None else min(2 ** (attempts - 1), 8)
                if (retryable and attempts <= self.config["max_retries"] and wait <= self.config["max_retry_wait_sec"]
                        and time.monotonic() + wait + 1 < self.deadline):
                    self.pause(wait)
                    continue
                raise

    def authorize(self):
        """Use the wallet access token, or exchange the refresh-token grant once (kept in memory only)."""
        if self.token is not None:
            return self.token
        if "access_token" in self.secret:
            self.token = self.secret["access_token"]
            return self.token
        fields = {"client_id": self.secret["client_id"], "grant_type": "refresh_token",
                  "refresh_token": self.secret["refresh_token"], "scope": self.config["oauth_scope"]}
        if "client_secret" in self.secret:
            fields["client_secret"] = self.secret["client_secret"]
        form = urlencode(fields).encode("ascii")
        headers = {"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json",
                   "Accept-Encoding": "identity", "User-Agent": USER_AGENT, "Connection": "close"}
        target = "/" + quote(self.config["tenant_id"], safe="") + "/oauth2/v2.0/token"

        def attempt():
            outcome = self.send(self.config["login_base_url"], "POST", target, headers, form, mutation=False)
            status, body = outcome["status"], outcome["body"]
            if status == 429:
                logic.fail("rate_limited", http_status=429, retry_after=outcome["retry_after"])
            if status >= 500:
                logic.fail("provider_unavailable", "The Microsoft token endpoint is temporarily unavailable (HTTP 5xx).",
                           http_status=status, retry_after=outcome["retry_after"])
            if status != 200 or outcome["body_error"] or not isinstance(body, dict):
                error = body.get("error") if isinstance(body, dict) else None
                description = body.get("error_description") if isinstance(body, dict) else None
                logic.fail("auth_failed", http_status=status,
                           provider_code=logic.clean(error, 64) if isinstance(error, str) else None,
                           provider_message=logic.clean(logic.redact(description, self.secrets), 200)
                           if isinstance(description, str) else None)
            token = body.get("access_token")
            if (str(body.get("token_type", "")).lower() != "bearer" or not isinstance(token, str)
                    or not logic.TOKEN_TEXT.fullmatch(token)):
                logic.fail("auth_failed")
            lifetime = body.get("expires_in")
            lifetime = lifetime if type(lifetime) is int and 0 < lifetime <= 86400 else 300
            return token, lifetime

        token, lifetime = self.retrying(attempt, mutation=False)
        self.token = token
        self.secrets.append(token)
        self.cache.update(token=token, fingerprint=secret_fingerprint(self.secret),
                          expires_at=time.monotonic() + lifetime)
        return self.token

    def call(self, method, target, query=None, payload=None, *, mutation, headers=None, raw=None, content_type=None,
             sink=None, empty=False, payload_limit=None):
        """Call a Graph path (or validated continuation path); return its JSON body.

        empty=True accepts a bodyless 2xx (202/204) and returns None. With a sink the
        body is streamed and the size/SHA-256 summary is returned.
        """
        token = self.authorize()
        path = target + ("?" + urlencode(query, quote_via=quote, safe="$,'():") if query else "")
        sent = {"Authorization": "Bearer " + token, "Accept": "*/*" if sink is not None else "application/json",
                "Accept-Encoding": "identity", "User-Agent": USER_AGENT, "Connection": "close", **(headers or {})}
        body = None
        if payload is not None:
            body = logic.encoded(payload, payload_limit or self.config["max_request_bytes"], "request_too_large")
            sent["Content-Type"] = "application/json"
        elif raw is not None:
            body = raw
            sent["Content-Type"] = content_type
        return self.retrying(lambda: self.receive(self.config["api_base_url"], method, path, sent, body, mutation=mutation,
                                                  sink=sink, empty=empty), mutation=mutation)

    def receive(self, origin, method, path, headers, body, *, mutation, sink=None, empty=False):
        """One attempt: map the response to a body, a stream summary or a coded error."""
        outcome = self.send(origin, method, path, headers, body, mutation=mutation, sink=sink)
        status = outcome["status"]
        if 200 <= status < 300:
            if outcome["streamed"]:
                if outcome["body_error"] == "too_large":
                    logic.fail("file_too_large", "The download passed max_download_bytes; the partial file was removed.")
                if outcome["body_error"]:
                    logic.fail("protocol")
                return {"size": sink.size, "sha256": sink.sha256()}
            if outcome["body_error"] == "empty" and empty:
                return None
            if outcome["body_error"] or not isinstance(outcome["body"], dict):
                logic.fail("protocol", uncertain=mutation)
            return outcome["body"]
        if status >= 500 and outcome["body_error"]:
            logic.fail("provider_unavailable", logic.MESSAGES["provider_unavailable_uncertain" if mutation else
                                                              "provider_unavailable"], uncertain=mutation, http_status=status,
                       retry_after=outcome["retry_after"])
        logic.provider_failure(status, outcome["body"], self.secrets, mutation=mutation, retry_after=outcome["retry_after"])

    def follow(self, link, *, headers=None):
        """GET a validated continuation path (nextLink/deltaLink) on the API origin."""
        token = self.authorize()
        sent = {"Authorization": "Bearer " + token, "Accept": "application/json", "Accept-Encoding": "identity",
                "User-Agent": USER_AGENT, "Connection": "close", **(headers or {})}
        return self.retrying(lambda: self.receive(self.config["api_base_url"], "GET", link, sent, None, mutation=False),
                             mutation=False)

    def transfer(self, origin, method, target, *, body=None, headers=None, mutation, sink=None, empty=False):
        """Exchange with a pre-authenticated download/upload URL: never with the Authorization header."""
        sent = {"Accept": "*/*" if sink is not None else "application/json", "Accept-Encoding": "identity",
                "User-Agent": USER_AGENT, "Connection": "close", **(headers or {})}
        return self.retrying(lambda: self.receive(origin, method, target, sent, body, mutation=mutation, sink=sink,
                                                  empty=empty), mutation=mutation)

    def discard_session(self, origin, target):
        """Best-effort cancellation of an upload session; True only when Graph confirms it was still open."""
        deadline = min(self.deadline, time.monotonic() + 3)
        if deadline - time.monotonic() <= 0.2:
            return False
        try:
            outcome = exchange(origin, self.config["ca_file"], "DELETE", target,
                               {"User-Agent": USER_AGENT, "Accept-Encoding": "identity", "Connection": "close"}, None,
                               deadline=deadline, cancel=lambda: False, limit=65536)
        except Failure:
            return False
        self.requests += 1
        return outcome["status"] == 204
