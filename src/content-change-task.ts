// src/content-change-task.ts
import { sleep, waitForElement } from "./common";

const log = (...args: any[]) =>
  console.log("[CI Updater][Change-Task]", ...args);

const CLOSE_TASK_ACTION_KEY = "ciUpdaterCloseTaskAction";
const CLOSE_TASK_TTL_MS = 3 * 60 * 1000;
const DEFAULT_CLOSE_NOTES = "Update CI เรียบร้อยครับ";

type CloseTaskAction = {
  chg?: string;
  ctask?: string;
  ts?: number;
};

function normalize(value: string): string {
  return (value || "").replace(/\s+/g, " ").trim();
}

function isChangeTaskPage(): boolean {
  const href = location.href;
  return /\/change_task\.do(?:$|\?|&)/i.test(href) || /\buri=change_task\.do/i.test(href);
}

function readCloseTaskAction(): CloseTaskAction | null {
  try {
    const raw = sessionStorage.getItem(CLOSE_TASK_ACTION_KEY);
    if (!raw) return null;
    return JSON.parse(raw) as CloseTaskAction;
  } catch {
    return null;
  }
}

function clearCloseTaskAction() {
  try {
    sessionStorage.removeItem(CLOSE_TASK_ACTION_KEY);
  } catch {}
}

function getPageCTaskNumber(): string {
  const input =
    document.querySelector<HTMLInputElement>(
      'input[name="change_task.number"], input[id$="change_task.number"]'
    ) || null;
  if (input?.value) return normalize(input.value).toUpperCase();

  const title =
    document.querySelector<HTMLElement>("h1, .navbar-title, .form_header") ||
    null;
  const mt = (title?.textContent || "").match(/CTASK\d+/i);
  if (mt) return mt[0].toUpperCase();

  return "";
}

async function waitForClosureTab(timeoutMs = 8000, pollMs = 200) {
  const t0 = Date.now();
  while (Date.now() - t0 < timeoutMs) {
    const tabs = Array.from(
      document.querySelectorAll<HTMLElement>(".tabs2_tab")
    );
    const tab = tabs.find((t) => {
      const label = t.querySelector<HTMLElement>(".tab_caption_text");
      const text = normalize(label?.textContent || "");
      return text.toLowerCase() === "closure information";
    });
    if (tab) return tab;
    await sleep(pollMs);
  }
  return null;
}

async function openClosureTabIfNeeded() {
  const tab = await waitForClosureTab();
  if (!tab) return;
  const selected = tab.getAttribute("aria-selected") === "true";
  if (!selected) {
    tab.click();
    await sleep(150);
  }
}

function setTextAreaValue(el: HTMLTextAreaElement, value: string) {
  const setter = Object.getOwnPropertyDescriptor(
    Object.getPrototypeOf(el),
    "value"
  )?.set;
  setter?.call(el, value);
  el.dispatchEvent(new Event("input", { bubbles: true }));
  el.dispatchEvent(new Event("change", { bubbles: true }));
}

async function setCloseCodeSuccessful() {
  try {
    const gf: any = (window as any).g_form;
    if (gf && typeof gf.setValue === "function") {
      gf.setValue("close_code", "successful");
      return;
    }
  } catch {}

  try {
    const select = await waitForElement<HTMLSelectElement>(
      "select#change_task\\.close_code, select[name=\"change_task.close_code\"]",
      8000,
      120
    );
    if (!select) return;
    const byValue = Array.from(select.options).find(
      (opt) => opt.value === "successful"
    );
    const byText = Array.from(select.options).find(
      (opt) => normalize(opt.textContent || "").toLowerCase() === "successful"
    );
    const target = byValue || byText;
    if (!target) return;
    select.focus();
    select.value = target.value;
    select.dispatchEvent(new Event("input", { bubbles: true }));
    select.dispatchEvent(new Event("change", { bubbles: true }));
  } catch {}
}

async function setCloseNotes(text: string) {
  const value = normalize(text);
  if (!value) return;
  try {
    const textarea = await waitForElement<HTMLTextAreaElement>(
      "textarea#change_task\\.close_notes, textarea[name=\"change_task.close_notes\"]",
      8000,
      120
    );
    if (!textarea) return;
    textarea.focus();
    setTextAreaValue(textarea, value);
    try {
      textarea.blur();
    } catch {}
  } catch {}
}

async function clickCloseTaskButton(): Promise<boolean> {
  try {
    const btn = await waitForElement<HTMLButtonElement>(
      "#change_task_to_closed_bottom, button[data-action-name=\"change_task_to_closed\"], #change_task_to_closed",
      8000,
      120
    );
    if (!btn) return false;
    btn.focus();
    btn.click();
    return true;
  } catch {}
  return false;
}

async function notifyCloseTaskDone() {
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
    });
  } catch {}
}

(async function run() {
  try {
    if (!isChangeTaskPage()) return;
    if ((window as any).__ciUpdaterCloseTaskDone) return;

    const action = readCloseTaskAction();
    if (!action) return;

    const ts = Number(action.ts || 0);
    if (!ts || Date.now() - ts > CLOSE_TASK_TTL_MS) {
      clearCloseTaskAction();
      return;
    }

    const pageCTask = getPageCTaskNumber();
    if (action.ctask && pageCTask && action.ctask.toUpperCase() !== pageCTask) {
      log("CTASK mismatch; skip", { flag: action.ctask, page: pageCTask });
      clearCloseTaskAction();
      return;
    }

    await openClosureTabIfNeeded();
    await setCloseCodeSuccessful();

    let closeNotes = "";
    let hasCloseNotesKey = false;
    try {
      const res = await chrome.storage.local.get("ciUpdaterCloseNotes");
      hasCloseNotesKey = Object.prototype.hasOwnProperty.call(
        res,
        "ciUpdaterCloseNotes"
      );
      closeNotes = normalize(res.ciUpdaterCloseNotes || "");
    } catch {}
    if (!hasCloseNotesKey) closeNotes = DEFAULT_CLOSE_NOTES;

    await setCloseNotes(closeNotes);

    (window as any).__ciUpdaterCloseTaskDone = true;
    clearCloseTaskAction();
    await sleep(200);
    const clicked = await clickCloseTaskButton();
    await sleep(100);
    await notifyCloseTaskDone();
    if (!clicked) {
      log("Close Task button not found; skip to next if possible");
    }
    log("Close fields set");
  } catch (e) {
    console.error("[CI Updater][Change-Task] error:", e);
    clearCloseTaskAction();
  }
})();
