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

function isChangeTasksTabSelected(): boolean {
  const tab = findChangeTasksTab();
  if (!tab) return false;
  return tab.getAttribute("aria-selected") === "true";
}

async function openChangeTasksTabIfNeeded(
  timeoutMs = 8000,
  pollMs = 180
): Promise<boolean> {
  const t0 = Date.now();
  while (Date.now() - t0 < timeoutMs) {
    const tab = findChangeTasksTab();
    if (!tab) {
      await sleep(pollMs);
      continue;
    }
    const selected = tab.getAttribute("aria-selected") === "true";
    if (selected) return true;
    tab.click();
    await sleep(220);
    if (tab.getAttribute("aria-selected") === "true") return true;
  }
  return isChangeTasksTabSelected();
}

function findChangeTasksTab(): HTMLElement | null {
  const exact = document.querySelector<HTMLElement>(
    'span.tabs2_tab[aria-controls="change_request.change_task.change_request_list"]'
  );
  if (exact) return exact;

  const tabs = Array.from(
    document.querySelectorAll<HTMLElement>("span.tabs2_tab, .tabs2_tab")
  );
  return (
    tabs.find((tab) => {
      const controls = normalize(tab.getAttribute("aria-controls") || "").toLowerCase();
      if (controls.includes("change_task")) return true;
      const label = normalize(tab.textContent || "").toLowerCase();
      return label.includes("change tasks");
    }) || null
  );
}

function isChangeTasksTable(table: HTMLTableElement): boolean {
  const id = normalize(table.id || "").toLowerCase();
  if (id.includes("change_task")) return true;
  if (table.querySelector('a[href*="change_task.do"]')) return true;
  const listRegion = table.closest<HTMLElement>('[id*="change_task"]');
  return !!listRegion;
}

function getChangeTasksListRegion(): HTMLElement | null {
  return (
    document.querySelector<HTMLElement>(
      "#change_request\\.change_task\\.change_request_list"
    ) ||
    document.querySelector<HTMLElement>(
      '[id*="change_request.change_task.change_request_list"]'
    ) ||
    null
  );
}

async function waitForChangeTasksTable(
  timeoutMs = 12000,
  pollMs = 200
): Promise<HTMLTableElement | null> {
  const t0 = Date.now();
  while (Date.now() - t0 < timeoutMs) {
    const exact =
      document.querySelector<HTMLTableElement>(
        "#change_request\\.change_task\\.change_request_table"
      ) ||
      getChangeTasksListRegion()?.querySelector<HTMLTableElement>(
        "table.list2_table, table.data_list_table.list_table, table"
      ) ||
      null;
    if (exact) return exact;

    if (!isChangeTasksTabSelected()) {
      await openChangeTasksTabIfNeeded(1200, pollMs);
    }

    const region = getChangeTasksListRegion();
    const tables = region
      ? Array.from(
          region.querySelectorAll<HTMLTableElement>(
            "table.list2_table, table.data_list_table.list_table, table"
          )
        )
      : [];
    const match = tables.find(isChangeTasksTable);
    if (match) return match;

    await sleep(pollMs);
  }
  return null;
}

function getTaskRows(table: HTMLTableElement): HTMLTableRowElement[] {
  const tbody = table.tBodies?.[0] || table;
  return Array.from(
    tbody.querySelectorAll<HTMLTableRowElement>("tr.list_row, tr")
  ).filter((tr) => tr.querySelectorAll("td").length > 0);
}

function findUpdateCiRow(
  table: HTMLTableElement
): HTMLTableRowElement | null {
  const rows = getTaskRows(table);
  const matches = rows.filter((tr) => isForUpdateCiRow(tr));
  const openUpdateCi = matches.find((tr) => !isClosedTaskRow(tr));
  if (openUpdateCi) return openUpdateCi;
  return matches[0] || null;
}

function getTaskStateText(row: HTMLTableRowElement): string {
  const cells = Array.from(row.querySelectorAll<HTMLTableCellElement>("td"));
  const stateCell =
    cells.find((td) => !!td.querySelector(".list2_cell_background")) || null;
  if (stateCell) return normalize(stateCell.textContent || "").toLowerCase();
  return normalize(row.textContent || "").toLowerCase();
}

function isClosedTaskRow(row: HTMLTableRowElement): boolean {
  const text = getTaskStateText(row);
  return (
    /\bclosed\b/i.test(text) ||
    text.includes("closed complete") ||
    text.includes("closed incomplete") ||
    text.includes("cancelled") ||
    text.includes("canceled")
  );
}

function isForUpdateCiText(textRaw: string): boolean {
  const text = normalize(textRaw).toLowerCase();
  if (!text) return false;
  return (
    /\bfor\s+update\s+ci\b/i.test(text) ||
    /\bfor\s+updating\s+ci\b/i.test(text) ||
    /\bfor\s+update\s+configuration\s+item\b/i.test(text) ||
    /\bupdate\s+ci\b/i.test(text)
  );
}

function getShortDescriptionText(row: HTMLTableRowElement): string {
  const link = findNumberLink(row);
  const numberCell = link?.closest("td") || null;
  if (numberCell) {
    const nextCell = numberCell.nextElementSibling as HTMLTableCellElement | null;
    if (nextCell) {
      const text = normalize(nextCell.textContent || "");
      if (text) return text;
    }
  }

  const cells = Array.from(row.querySelectorAll<HTMLTableCellElement>("td"));
  if (cells.length >= 4) {
    const text = normalize(cells[3]?.textContent || "");
    if (text) return text;
  }
  return normalize(row.textContent || "");
}

function isForUpdateCiRow(row: HTMLTableRowElement): boolean {
  const shortDesc = getShortDescriptionText(row);
  if (isForUpdateCiText(shortDesc)) return true;
  return isForUpdateCiText(normalize(row.textContent || ""));
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

function findForUpdateCiTaskLink(table: HTMLTableElement): HTMLAnchorElement | null {
  const updateRow = findUpdateCiRow(table);
  if (updateRow) {
    const link = findNumberLink(updateRow);
    if (link) return link;
  }
  return null;
}

function findForUpdateCiTaskLinkAnywhere(): HTMLAnchorElement | null {
  const rows = Array.from(
    document.querySelectorAll<HTMLTableRowElement>("tr.list_row, tr")
  ).filter((tr) => tr.querySelectorAll("td").length > 0);

  for (const row of rows.filter((tr) => isForUpdateCiRow(tr) && !isClosedTaskRow(tr))) {
    const link = findNumberLink(row);
    if (link) return link;
  }

  for (const row of rows.filter((tr) => isForUpdateCiRow(tr))) {
    const link = findNumberLink(row);
    if (link) return link;
  }

  const anchors = Array.from(
    document.querySelectorAll<HTMLAnchorElement>('a[href*="change_task.do"], a')
  );
  const updateAnchor = anchors.find((a) => {
    const rowText = normalize(a.closest("tr")?.textContent || "");
    const text = normalize(a.textContent || "");
    return (
      /CTASK\d+/i.test(text) &&
      (/\b(?:for\s+)?(?:update|updating)\s+ci\b/i.test(rowText) ||
        /\bci\s*update\b/i.test(rowText))
    );
  });
  if (updateAnchor) return updateAnchor;
  return null;
}

async function waitForUpdateCiTaskLink(
  timeoutMs = 15000,
  pollMs = 220
): Promise<HTMLAnchorElement | null> {
  const t0 = Date.now();
  while (Date.now() - t0 < timeoutMs) {
    await openChangeTasksTabIfNeeded(1200, pollMs);
    const table = await waitForChangeTasksTable(1200, pollMs);
    const link = (table ? findForUpdateCiTaskLink(table) : null)
      || findForUpdateCiTaskLinkAnywhere();
    if (link) return link;
    await sleep(pollMs);
  }
  return null;
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

    const tabReady = await openChangeTasksTabIfNeeded(12000, 220);
    if (!tabReady) {
      log("Unable to open Change Tasks tab; keep close phase");
      return;
    }

    const link = await waitForUpdateCiTaskLink(16000, 220);
    if (!link) {
      log("For Update CI task not found; keep close phase");
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
  }
})();
