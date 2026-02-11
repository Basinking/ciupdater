// src/content-change-request.ts
import { sleep } from "./common";

const log = (...args: any[]) =>
  console.log("[CI Updater][Change-Request]", ...args);

const CLOSE_ACTION_KEY = "ciUpdaterCloseAction";
const CLOSE_TASK_ACTION_KEY = "ciUpdaterCloseTaskAction";
const CLOSE_ACTION_TTL_MS = 3 * 60 * 1000;

type CloseAction = {
  chg?: string;
  ts?: number;
};
type CloseTaskAction = {
  chg?: string;
  ctask?: string;
  ts?: number;
};

function normalize(value: string): string {
  return (value || "").replace(/\s+/g, " ").trim();
}

function isChangeRequestPage(): boolean {
  const href = location.href;
  return (
    /\/change_request\.do(?:$|\?|&)/i.test(href) ||
    /\buri=change_request\.do/i.test(href)
  );
}

function readCloseAction(): CloseAction | null {
  try {
    const raw = sessionStorage.getItem(CLOSE_ACTION_KEY);
    if (!raw) return null;
    return JSON.parse(raw) as CloseAction;
  } catch {
    return null;
  }
}

function clearCloseAction() {
  try {
    sessionStorage.removeItem(CLOSE_ACTION_KEY);
  } catch {}
}

function setCloseTaskAction(payload: CloseTaskAction) {
  try {
    sessionStorage.setItem(CLOSE_TASK_ACTION_KEY, JSON.stringify(payload));
  } catch {}
}

async function notifySkipToNext(reason: string) {
  try {
    const { ciUpdaterClosing, ciUpdaterRunId } = await chrome.storage.local.get([
      "ciUpdaterClosing",
      "ciUpdaterRunId",
    ]);
    const resumeIndex = Number(ciUpdaterClosing?.resumeIndex);
    if (!Number.isFinite(resumeIndex)) return;
    await chrome.runtime.sendMessage({
      type: "CLOSE_TASK_DONE",
      runId: ciUpdaterRunId || ciUpdaterClosing?.runId,
      resumeIndex,
      reason,
    });
  } catch {}
}

function getPageChgNumber(): string {
  const input =
    document.querySelector<HTMLInputElement>(
      'input[name="change_request.number"], input[id$="change_request.number"]'
    ) || null;
  if (input?.value) return normalize(input.value).toUpperCase();

  const breadcrumb =
    document.querySelector<HTMLElement>(
      "#change_request\\.change_task\\.change_request_breadcrumb"
    ) || document.querySelector<HTMLElement>(".breadcrumb_container");
  const text = breadcrumb?.textContent || "";
  const m = text.match(/CHG\d+/i);
  if (m) return m[0].toUpperCase();

  const title =
    document.querySelector<HTMLElement>("h1, .navbar-title, .form_header") ||
    null;
  const mt = (title?.textContent || "").match(/CHG\d+/i);
  if (mt) return mt[0].toUpperCase();

  return "";
}

async function openChangeTasksTabIfNeeded() {
  const tab = document.querySelector<HTMLElement>(
    'span.tabs2_tab[aria-controls="change_request.change_task.change_request_list"]'
  );
  if (!tab) return;
  const selected = tab.getAttribute("aria-selected") === "true";
  if (!selected) {
    tab.click();
    await sleep(150);
  }
}

async function waitForChangeTasksTable(
  timeoutMs = 12000,
  pollMs = 200
): Promise<HTMLTableElement | null> {
  const t0 = Date.now();
  while (Date.now() - t0 < timeoutMs) {
    const table = document.querySelector<HTMLTableElement>(
      "#change_request\\.change_task\\.change_request_table"
    );
    if (table) return table;
    await sleep(pollMs);
  }
  return null;
}

function hasNumberColumn(table: HTMLTableElement): boolean {
  return !!table.querySelector(
    'thead th[name="number"], thead th[glide_label="Number"]'
  );
}

function findUpdateCiRow(
  table: HTMLTableElement
): HTMLTableRowElement | null {
  const tbody = table.tBodies?.[0] || table;
  const rows = Array.from(
    tbody.querySelectorAll<HTMLTableRowElement>("tr.list_row, tr")
  ).filter((tr) => tr.querySelectorAll("td").length > 0);
  return (
    rows.find((tr) =>
      normalize(tr.textContent || "").toLowerCase().includes("for update ci")
    ) || null
  );
}

function findNumberLink(row: HTMLTableRowElement): HTMLAnchorElement | null {
  const anchors = Array.from(row.querySelectorAll<HTMLAnchorElement>("a"));
  const byText = anchors.find((a) =>
    /CTASK\d+/i.test(normalize(a.textContent || ""))
  );
  if (byText) return byText;
  const byHref = anchors.find((a) =>
    /change_task\.do/i.test(a.getAttribute("href") || "")
  );
  return byHref || null;
}

(async function run() {
  try {
    if (!isChangeRequestPage()) return;

    const action = readCloseAction();
    if (!action) return;

    const ts = Number(action.ts || 0);
    if (!ts || Date.now() - ts > CLOSE_ACTION_TTL_MS) {
      clearCloseAction();
      return;
    }

    const pageChg = getPageChgNumber();
    if (action.chg && pageChg && action.chg.toUpperCase() !== pageChg) {
      log("CHG mismatch; skip", { flag: action.chg, page: pageChg });
      clearCloseAction();
      return;
    }

    await openChangeTasksTabIfNeeded();
    const table = await waitForChangeTasksTable();
    if (!table) {
      log("Change Tasks table not found; skip");
      clearCloseAction();
      return;
    }

    if (!hasNumberColumn(table)) {
      log("Number column not found; skip");
      clearCloseAction();
      await notifySkipToNext("number_column_missing");
      return;
    }

    const row = findUpdateCiRow(table);
    if (!row) {
      log("For Update CI row not found; skip");
      clearCloseAction();
      await notifySkipToNext("update_ci_row_missing");
      return;
    }

    const link = findNumberLink(row);
    if (!link) {
      log("Change Task link not found; skip");
      clearCloseAction();
      await notifySkipToNext("change_task_link_missing");
      return;
    }
    try {
      const ctaskMatch = normalize(link.textContent || "").match(/CTASK\d+/i);
      setCloseTaskAction({
        chg: (action.chg || pageChg || "").trim(),
        ctask: ctaskMatch ? ctaskMatch[0].toUpperCase() : undefined,
        ts: Date.now(),
      });
    } catch {}

    try {
      (link as HTMLElement).scrollIntoView?.({
        block: "center",
        inline: "nearest",
      });
    } catch {}
    link.click();
    log("Open Change Task", normalize(link.textContent || ""));
    clearCloseAction();
  } catch (e) {
    console.error("[CI Updater][Change-Request] error:", e);
    clearCloseAction();
  }
})();
