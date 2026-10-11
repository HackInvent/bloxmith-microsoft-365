"""FB9: real release UI in English and French, responsive geometry, conditional write and watch sections, drafts and persistence."""

import json
from pathlib import Path
import sys
import time

ROOT = Path(__file__).resolve().parents[3]
sys.path[:0] = [str(ROOT), str(ROOT / "tests")]
from playwright.sync_api import sync_playwright, expect, TimeoutError as PlaywrightTimeoutError
from block_test_artifacts import artifact_path
from block_test_packages import install_test_package, surface_payload
from ui_smoke_common import isolated_server, graph_payload, create_project_api, project_editor_url
from blocs.microsoft_365.block import Microsoft365Block

VIEWPORTS = ((1440, 900), (800, 700), (390, 740), (320, 568))
DRIVE_FOLDER = "Shared/Reports 2026"
ATTACHMENT_ROOT = "/srv/bloxsmith/attachments"


def open_modal(card, modal):
    """Open the modal from the canvas card, retrying while the editor binds its handlers on a loaded host.

    A retry is sent only after a long wait: on a slow host an early second double-click
    can open the modal again later, after the test has closed it.
    """
    expect(card.locator('h3')).to_be_visible(timeout=30000)
    for _ in range(4):
        if modal.is_visible():
            return
        try:
            card.locator('h3').dblclick(timeout=5000)
            expect(modal).to_be_visible(timeout=15000)
            return
        except (AssertionError, PlaywrightTimeoutError):
            continue
    expect(modal).to_be_visible()


def settle_closed(page, modal):
    """Keep the modal closed before using the inspector, even if a retried open arrives late."""
    for _ in range(4):
        page.wait_for_timeout(500)
        if not modal.is_visible():
            return
        close_modal(modal)
    expect(modal).to_be_hidden()


def close_modal(modal, which="first"):
    """Close the modal and wait until it is gone, so a reopen binds to a fresh surface."""
    buttons = modal.locator('[data-close-block-modal]')
    (buttons.last if which == "last" else buttons.first).click()
    expect(modal).to_be_hidden()


def geometry(modal):
    """Require an opaque panel with reachable actions, no page overflow and labelled controls."""
    deadline = time.monotonic() + 5
    while True:
        measured = measure(modal)
        if (measured["inside"] and measured["controls"] and not measured["overflow"]) or time.monotonic() > deadline:
            break
        time.sleep(0.1)
    assert measured["inside"] and measured["controls"] and not measured["overflow"], measured
    assert not measured["unlabeled"], measured
    assert measured["background"] not in {"transparent", "rgba(0, 0, 0, 0)"}, measured


def measure(modal):
    return modal.evaluate("""el => {
        const r=el.getBoundingClientRect(), body=el.querySelector('.owned-body');
        const actions=[...el.querySelectorAll('[data-close-block-modal], [data-owned-apply]')];
        const visible=[...el.querySelectorAll('input,select,textarea')].filter(c=>c.getClientRects().length);
        return {inside:r.left>=-1&&r.right<=innerWidth+1&&r.top>=-1&&r.bottom<=innerHeight+1,
          overflow:el.scrollWidth>el.clientWidth+2||body.scrollWidth>body.clientWidth+2,
          controls:actions.every(a=>{const b=a.getBoundingClientRect();return b.top>=0&&b.bottom<=innerHeight+1}),
          background:getComputedStyle(el).backgroundColor,
          unlabeled:visible.filter(c=>!c.labels?.length&&!c.getAttribute('aria-label')&&!c.getAttribute('aria-labelledby')).length};
    }""")


def apply_expecting(page, modal, fragment):
    """Apply and return the refusal, which must name the setting that is missing."""
    with page.expect_response(lambda r: r.url.endswith('/ui-action') and r.request.method == 'POST') as answer:
        modal.locator('[data-owned-apply]').click()
    error = answer.value.json().get("error", "")
    assert fragment in error, answer.value.json()
    expect(modal.locator('.owned-form-status')).to_contain_text(fragment)


def main():
    """Exercise current package files in real managed and linked browser hosts."""
    block = Microsoft365Block()
    expect.set_options(timeout=15000)
    for origin in ("managed", "linked"):
        with isolated_server() as server, sync_playwright() as playwright:
            model = install_test_package(server, "microsoft_365", origin=origin)
            candidate = block.build_node_payload(node_id="ui-case", position={"x": 160, "y": 160})
            candidate["block_version"] = model["version"]
            for surface in ("modal", "inspector_panel", "node_card"):
                html = surface_payload(server, model, candidate, surface=surface)["html"]
                assert "{{" not in html
            project = create_project_api(server, document=graph_payload("Owned form", [candidate], []))["project"]
            url = project_editor_url(server.base_url, project["project_id"], workspace_project_id=project["workspace_project_id"])
            browser = playwright.chromium.launch(headless=True)
            try:
                page = browser.new_page(viewport={"width": 1440, "height": 900})
                page.add_init_script("window.localStorage.setItem('bloxsmith.inspectorPinned','true')")
                errors = []
                page.on("pageerror", lambda error: errors.append(str(error)))
                page.goto(url)
                card = page.locator('.canvas-node[data-node-id="ui-case"]')
                expect(card.locator('h3')).to_be_visible()
                expect(card.locator('.node-data-preview')).to_have_text("Remote access off")
                modal = page.locator('.owned-modal')
                open_modal(card, modal)
                expect(modal.locator('[data-owned-apply]')).to_be_disabled()
                field = lambda key: modal.locator(f'[data-block-config-field="{key}"]')
                writes, send, recipients = field("allow_writes"), field("allow_send"), field("allowed_recipients")
                watch, interval = field("watch_mail"), field("watch_interval_sec")
                for hidden in (send, recipients, field("attachment_root"), field("allow_delete"), interval,
                               field("emit_existing")):
                    expect(hidden).to_be_hidden()
                writes.check()
                watch.check()
                for shown in (send, recipients, field("allowed_channels"), field("attachment_root"), field("upload_root"),
                              field("allow_delete"), field("conflict_behavior"), interval, field("existing_since_hours")):
                    expect(shown).to_be_visible()
                expect(modal.locator('[data-owned-apply]')).to_be_enabled()
                close_modal(modal, "last")
                open_modal(card, modal)
                expect(writes).not_to_be_checked()
                expect(send).to_be_hidden()
                expect(interval).to_be_hidden()
                # Invalid combinations are refused on Apply without changing the saved settings.
                field("enabled").check()
                apply_expecting(page, modal, "secret_ref")
                field("secret_ref").fill("secret://workspace/microsoft_365")
                field("tenant_id").fill("contoso.onmicrosoft.com")
                field("operation").select_option("mail_send")
                apply_expecting(page, modal, "allow_writes")
                writes.check()
                apply_expecting(page, modal, "allow_send")
                send.check()
                apply_expecting(page, modal, "allowed_recipients")
                recipients.fill('["team@example.com", "@example.org"]')
                field("allowed_operations").fill('["mail_list", "files_list"]')
                field("attachment_root").fill(ATTACHMENT_ROOT)
                field("drive_folder").fill(DRIVE_FOLDER)
                watch.check()
                interval.fill("30")
                modal.locator('[data-owned-apply]').click()
                # The 60 s minimum is enforced before any request leaves the form (and again on the server).
                expect(modal.locator('.owned-form-status')).to_contain_text("Check the numeric value")
                interval.fill("120")
                with page.expect_response(lambda r: r.url.endswith('/ui-action') and r.request.method == 'POST') as saved:
                    modal.locator('[data-owned-apply]').click()
                assert not saved.value.json().get("error"), saved.value.json()
                if modal.is_visible():
                    close_modal(modal)
                page.reload()
                expect(card.locator('.node-data-preview')).to_have_text("Watching new mail · mail_send")
                for private in (DRIVE_FOLDER, ATTACHMENT_ROOT, "contoso", "team@example.com"):
                    assert private not in card.inner_text()
                open_modal(card, modal)
                assert json.loads(recipients.input_value()) == ["team@example.com", "@example.org"]
                assert json.loads(field("allowed_operations").input_value()) == ["mail_list", "files_list"]
                expect(field("operation")).to_have_value("mail_send")
                expect(field("tenant_id")).to_have_value("contoso.onmicrosoft.com")
                expect(field("drive_folder")).to_have_value(DRIVE_FOLDER)
                expect(field("attachment_root")).to_have_value(ATTACHMENT_ROOT)
                expect(interval).to_have_value("120")
                expect(writes).to_be_checked()
                expect(send).to_be_checked()
                expect(watch).to_be_checked()
                for width, height in VIEWPORTS:
                    page.set_viewport_size({"width": width, "height": height})
                    page.wait_for_timeout(100)
                    geometry(modal)
                    page.screenshot(path=artifact_path("microsoft-365-" + origin + "-" + str(width) + ".png"))
                summary = modal.locator('.owned-advanced').first.locator('summary')
                summary.scroll_into_view_if_needed()
                summary.focus()
                summary.press('Enter')
                expect(modal.locator('.owned-advanced').first).to_have_attribute('open', '')
                expect(field("page_size")).to_be_visible()
                expect(field("api_base_url")).to_have_value("https://graph.microsoft.com")
                geometry(modal)
                page.screenshot(path=artifact_path("microsoft-365-" + origin + "-advanced-320.png"))
                page.set_viewport_size({"width": 1440, "height": 900})
                close_modal(modal)
                settle_closed(page, modal)
                card.click(position={"x": 20, "y": 20})
                settle_closed(page, modal)
                inspector = page.locator('[data-properties-surface="inspector"][data-node-id="ui-case"]:visible')
                expect(inspector).to_be_visible()
                assert inspector.evaluate('el => el.scrollWidth <= el.clientWidth+2')
                panel_writes = inspector.locator('[data-block-config-field="allow_writes"]')
                expect(panel_writes).to_be_checked()
                expect(inspector.locator('[data-block-config-field="allowed_recipients"]')).to_be_visible()
                panel_writes.uncheck()
                expect(inspector.locator('[data-block-config-field="allowed_recipients"]')).to_be_hidden()
                expect(inspector.locator('[data-block-config-field="watch_interval_sec"]')).to_be_visible()
                inspector.locator('[data-block-config-field="watch_mail"]').uncheck()
                expect(inspector.locator('[data-block-config-field="watch_interval_sec"]')).to_be_hidden()
                inspector.locator('[data-inspector-tab="ports"]').click()
                expect(inspector.locator('[data-inspector-panel-tab="ports"]')).to_be_visible()
                for port in ("request", "file", "items", "result", "error"):
                    expect(inspector.locator('[data-inspector-panel-tab="ports"]')).to_contain_text(port)
                inspector.locator('[data-inspector-tab="general"]').click()
                expect(inspector.locator('[data-inspector-panel-tab="general"]')).to_be_visible()
                page.screenshot(path=artifact_path("microsoft-365-" + origin + "-inspector.png"))
                page.goto(server.base_url + "/")
                page.locator("#homeApplicationSettingsButton").click()
                page.locator("#applicationLanguageSelect").select_option("fr")
                page.wait_for_function("window.CWMessages.getLanguage() === 'fr'")
                page.goto(url)
                # The card preview follows the interface language (the inspector draft above was not applied).
                expect(card.locator('.node-data-preview')).to_have_text("Surveillance du courrier · mail_send")
                open_modal(card, modal)
                expect(modal.locator('.owned-advanced').first.locator('summary')).to_have_text("Réglages avancés")
                expect(modal.locator('.owned-notice')).to_contain_text("L’autorisation OAuth reste dans le wallet")
                expect(modal.locator('label[for$="-enabled"]')).to_have_text("Activer l’accès distant")
                expect(modal.locator('label[for$="-allow_writes"]')).to_have_text("Autoriser l’écriture")
                expect(modal.locator('label[for$="-allow_send"]')).to_have_text("Autoriser l’envoi")
                expect(modal.locator('label[for$="-watch_mail"]')).to_have_text("Surveiller le nouveau courrier")
                expect(modal.locator('[data-owned-apply]')).to_have_text("Appliquer")
                for width, height in ((1440, 900), (390, 740)):
                    page.set_viewport_size({"width": width, "height": height})
                    page.wait_for_timeout(100)
                    geometry(modal)
                    page.screenshot(path=artifact_path("microsoft-365-" + origin + "-fr-" + str(width) + ".png"))
                assert not errors, errors
                assert page.evaluate("!window.CWBlockUiBlocks?.microsoft_365")
            finally:
                browser.close()
        print("[ok] microsoft_365 " + origin + " UI", flush=True)


if __name__ == "__main__":
    main()
