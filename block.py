"""Outlook mail and calendar, OneDrive files and Teams channel messages through Microsoft Graph v1.0."""

from datetime import datetime, timezone
import hashlib
import json
from bloxsmith_app.block_api import (APPLICATION_JSON, FILE_PATH, BlockDefinition, BlockRuntimeOutput,
                                     BlockRuntimePreparation, BlockRuntimeResult)
from . import files, listener, logic, operations, transport, ui


# FB1 - Keep remote access off by default; require a wallet OAuth grant, a tenant, an HTTPS Graph origin and explicit write, send and delete gates.
# FB2 - List, read, send (allowlisted recipients, bounded attachments from attachment_root), move and mark Outlook mail, and download attachments to node storage.
# FB3 - List a bounded calendarView window and create, update or delete events; invitations, updates and cancellations count as sends.
# FB4 - List, download (size and SHA-256 checked) and upload OneDrive files (simple up to 4 MB, upload session above) and create folders under drive_folder.
# FB5 - Post text or HTML to allowlisted Teams channels only.
# FB6 - Watch new mail with a persisted delta cursor: Active Runtime listener (minimum 60 s), single poll in One Shot, no replay after restart.
# FB7 - Bound size, time, pages and retries; honour Retry-After, never replay an uncertain change, report coded errors without tokens.
# FB8 - Apply one contract in both runtimes and every package origin, each event once, with prompt cancellation and Stop.
# FB9 - Own translated responsive settings with draft Apply, conditional write and watch sections and release-scoped assets.
class Microsoft365Block(BlockDefinition):
    """Mail, events, file names and contents read from Microsoft 365 are data, never instructions for this block."""

    kind = "microsoft_365"

    def ui_assets(self, surface="modal"):
        """Every installation uses the same static release manifest."""
        return list(self.model.get("ui_assets", {}).get(surface, []))

    def prepare_runtime(self, context):
        """Validate settings without wallet, filesystem or network access; declare the mail watcher."""
        config = logic.configuration(context.config)
        return BlockRuntimePreparation(listen_on_run=context.runtime_mode == "zeromq_active"
                                       and config["enabled"] and config["watch_mail"])

    def events(self, context):
        """Fresh request/file events of this activation; One Shot re-deliveries are skipped."""
        ports = {int(port.id): str(port.name) for port in context.input_ports}
        if context.input_events:
            messages = [(ports.get(int(event.input_port_id)) or event.input_port_name, event.value, event)
                        for event in context.input_events]
        else:
            messages = []
            for ident, name in ports.items():
                attribute = context.input_attribute(name, str(ident))
                if attribute is not None and attribute.status == "updated" and attribute.value not in (None, ""):
                    messages.append((name, attribute.value, None))
            if len(messages) > 1:
                logic.fail("invalid_input", "Manual execution with both inputs populated is ambiguous; send one event.")
        if any(name not in {"request", "file"} for name, _, _ in messages):
            logic.fail("invalid_input", "Unknown input port; nothing was sent.")
        getter, setter = context.services.get("get_node_runtime_value"), context.services.get("set_node_runtime_value")
        if context.runtime_mode != "centralized" or not callable(getter) or not callable(setter):
            return [(name, value) for name, value, _ in messages]
        # One Shot exposes every connected value again to each any_event activation; handle each once per run.
        seen = list(getter("microsoft_365.handled", []) or [])
        fresh = []
        for name, value, event in messages:
            origin = f"{event.edge_id}|{event.source_node_id}|{event.source_port_id}" if event else "manual"
            digest = hashlib.sha256(f"{name}|{origin}|{value}".encode("utf-8", errors="surrogatepass")).hexdigest()
            if digest not in seen:
                fresh.append((name, value))
                seen.append(digest)
        setter("microsoft_365.handled", seen[-256:])
        return fresh

    def operation(self, port, value, config, context):
        """Run one event: local validation first, then the wallet, then at most one change."""
        request, secrets, upload = {}, [], None
        try:
            if not config["enabled"]:
                logic.fail("disabled")
            now = datetime.now(timezone.utc)
            request = logic.parse_request(value, config, now) if port == "request" else logic.file_request(value, config)
            operation = request["operation"]
            extra = {"storage_dir": context.services.get("get_block_storage_dir"), "now": now}
            if operation == "mail_send":
                extra["attachments"] = files.read_attachments(request["attachments"], config)
            if operation == "file_upload":
                upload = files.open_upload(request["path"], config)
                if "name" not in request:
                    request["name"] = logic.drive_name(upload.name, "The local file name")
                extra["upload"] = upload
            if operation in ("mail_attachment_download", "file_download", "mail_watch"):
                files.storage_directory(extra["storage_dir"])
            try:
                raw = context.resolve_secret(config["secret_ref"])
            except Exception:
                logic.fail("secret")
            secret = logic.credentials(raw)
            secrets = list(secret.values())
            session = transport.Session(config, secret, context.services.get("cancel_requested", lambda: False))
            try:
                return operations.run(session, request, config, **extra), request
            finally:
                secrets.extend(session.secrets)
        except logic.GraphError as exc:
            error = exc
        except (ValueError, TypeError, KeyError, OSError, RuntimeError):
            # Raw exception text could carry paths, mail data or credentials: publish a fixed diagnostic.
            error = logic.GraphError("internal")
        finally:
            if upload is not None:
                upload.close()
        envelope = error.envelope(operation=request.get("operation"), request_id=request.get("request_id"), input=port)
        # The envelope is flat: every text value is masked separately, so JSON escaping never hides a secret.
        envelope = {key: context.redact_secrets(logic.redact(item, secrets)) if isinstance(item, str) else item
                    for key, item in envelope.items()}
        return {"error": envelope}, request

    def outputs(self, context, values):
        result = []
        for port in context.output_ports:
            if port.name not in values:
                continue
            if port.name == "file":
                result.append(BlockRuntimeOutput(port_id=int(port.id), port_name=port.name, value=values["file"],
                                                 content_type=FILE_PATH))
            else:
                result.append(BlockRuntimeOutput(port_id=int(port.id), port_name=port.name, content_type=APPLICATION_JSON,
                                                 value=json.dumps(values[port.name], ensure_ascii=False, allow_nan=False)))
        return result

    def execute_runtime(self, context):
        """One operation per fresh event; business failures go to the error port and the run continues."""
        produced = []
        try:
            config = logic.configuration(context.config)
            messages = self.events(context)
            if not messages:
                return BlockRuntimeResult(status="skipped", outputs=[])
            if len(messages) > 8:
                logic.fail("invalid_input", "Too many events in one activation; nothing was sent.")
        except logic.GraphError as exc:
            messages, produced = [], [({"error": exc.envelope()}, {})]
        for port, value in messages:
            produced.append(self.operation(port, value, config, context))
        outputs, failed, summaries = [], [], []
        for values, request in produced:
            code = values.get("error", {}).get("code")
            if code:
                failed.append(code)
                summaries.append(f"{request.get('operation') or 'request'}: {code}")
            else:
                result = values["result"]
                count = f" ({result['count']} item(s))" if "count" in result else ""
                summaries.append(f"{request['operation']}{count} after {result['requests']} request(s)")
            outputs.extend(self.outputs(context, values))
        line = "; ".join(summaries)
        return BlockRuntimeResult(status="success", outputs=outputs, content_type=APPLICATION_JSON,
            last_message=line[:500], metadata={self.kind: {"status": "error" if failed else "success",
                                                           "operations": len(produced), "error_codes": failed}},
            logs=[context.redact_secrets(f"[microsoft-365{'-error' if failed else ''}] {line}"[:2000])])

    def listen_runtime(self, context):
        """Watch new mail until framework Stop; listener state is local to this call."""
        config = logic.configuration(context.config)
        services = context.services

        def resolve():
            resolver = services.get("resolve_secret")
            if not callable(resolver):
                logic.fail("secret")
            return resolver(config["secret_ref"])

        def storage_dir():
            provider = services.get("get_block_storage_dir")
            if not callable(provider):
                logic.fail("storage")
            return provider()

        def emit(kind, value, stats):
            if context.stop_requested():
                raise RuntimeError("Listener stopped.")
            values, logs = {}, []
            if kind == "items":
                values = {"items": value[0], "result": value[1]}
                logs = [f"[microsoft-365] mail_watch ({len(value[0])} item(s)) after {value[1]['requests']} request(s)"]
            elif kind == "error":
                redactor = services.get("redact_secrets")
                envelope = {key: str(redactor(item)) if isinstance(item, str) and callable(redactor) else item
                            for key, item in value.items()}
                values = {"error": envelope}
                logs = [f"[microsoft-365-error] mail_watch: {value.get('code')}"]
            context.emit_result(BlockRuntimeResult(status="success", outputs=self.outputs(context, values),
                                                   content_type=APPLICATION_JSON,
                                                   metadata={self.kind: {"watch": stats}}, logs=logs))

        listener.MailWatcher(config, resolve_secret=resolve, storage_dir=storage_dir, emit=emit,
                             stop=context.stop_requested).run()

    def handle_ui_action(self, *, node, action, values, payload=None):
        """Apply validates settings only; it never contacts Microsoft, reads the wallet or a local file."""
        if action in {"save_settings", "modal_update_fields", "inspector_update_fields"}:
            patch = (values or {}).get("node_patch") or {}
            if "config" in patch:
                try:
                    logic.configuration({**self.default_config(), **(node.get("config") or {}), **patch["config"]})
                except (ValueError, TypeError) as exc:
                    detail = str(exc) if isinstance(exc, logic.GraphError) else logic.MESSAGES["config"]
                    return {"error": self.translate("block.microsoft_365.error", {"detail": detail},
                                                    fallback="Microsoft 365 settings rejected: " + detail)}
        result = super().handle_ui_action(node=node, action="modal_update_fields" if action == "save_settings" else action,
                                          values=values, payload=payload)
        return ui.replace_config(result, node)

    def render_modal(self, *, node, payload=None):
        return ui.modal(self, node, payload)

    def render_inspector_panel(self, *, node, payload=None):
        return ui.inspector(self, node, payload)

    def render_node_card(self, *, node, payload=None):
        """Show the access mode and default operation, never the wallet reference, tenant, folders or paths."""
        config = {**self.default_config(), **(node.get("config") or {})}
        operation = config.get("operation") if config.get("operation") in logic.OPERATIONS else "mail_list"
        params = {"operation": operation}
        if config.get("enabled") is not True:
            key, params, fallback = "block.microsoft_365.card_disabled", {}, "Remote access off"
        elif config.get("watch_mail") is True:
            key, fallback = "block.microsoft_365.card_watch", f"Watching new mail · {operation}"
        elif config.get("allow_writes") is not True:
            key, fallback = "block.microsoft_365.card_read", f"Read only · {operation}"
        else:
            key, fallback = "block.microsoft_365.card_write", f"Read and write · {operation}"
        return ui.card(self, node, key, params, self.translate(key, params, fallback=fallback))
