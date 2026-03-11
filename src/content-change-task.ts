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
  if (
    /\/change_task\.do(?:$|\?|&)/i.test(href) ||
    /\buri=(?:%2f)?change_task\.do/i.test(href)
  ) {
    return true;
  }
  try {
    const uri = new URL(href).searchParams.get("uri") || "";
    if (/change_task\.do/i.test(decodeURIComponent(uri))) return true;
  } catch {}
  return false;
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
      const text = normalize(label?.textContent || "").toLowerCase();
      const controls = normalize(t.getAttribute("aria-controls") || "").toLowerCase();
      return (
        text === "closure information" ||
        text.includes("closure information") ||
        text.includes("closure") ||
        controls.includes("closure")
      );
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

function isClosedText(value: string): boolean {
  const text = normalize(value).toLowerCase();
  if (!text) return false;
  if (/^-?\d+$/.test(text)) {
    // ServiceNow commonly uses numeric state values for Change Task.
    // Closed-like values observed across instances: 3/4/7/8
    return ["3", "4", "7", "8"].includes(text);
  }
  return (
    text.includes("closed") ||
    text.includes("closed complete") ||
    text.includes("closed incomplete") ||
    text.includes("cancelled") ||
    text.includes("canceled")
  );
}

function getStateTextFromGForm(): string {
  try {
    const gf: any = (window as any).g_form;
    if (!gf) return "";
    const displayValue =
      typeof gf.getDisplayValue === "function"
        ? normalize(gf.getDisplayValue("state") || "")
        : "";
    if (displayValue) return displayValue;
    const value =
      typeof gf.getValue === "function" ? normalize(gf.getValue("state") || "") : "";
    return value;
  } catch {
    return "";
  }
}

function getStateTextFromDom(): string {
  const select =
    document.querySelector<HTMLSelectElement>(
      'select#change_task\\.state, select[name="change_task.state"], select[id$=".state"], select[name$=".state"]'
    ) || null;
  if (select) {
    const option = select.selectedOptions?.[0] || null;
    const text = normalize(option?.textContent || select.value || "");
    if (text) return text;
  }

  const input =
    document.querySelector<HTMLInputElement>(
      'input#change_task\\.state, input[name="change_task.state"], input[id$=".state"], input[name$=".state"]'
    ) || null;
  if (input?.value) return normalize(input.value);

  return "";
}

function isTaskClosedNow(): boolean {
  const fromGForm = getStateTextFromGForm();
  if (isClosedText(fromGForm)) return true;
  const fromDom = getStateTextFromDom();
  return isClosedText(fromDom);
}

async function waitForCloseConfirmation(
  timeoutMs = 30000,
  pollMs = 250
): Promise<boolean> {
  const t0 = Date.now();
  while (Date.now() - t0 < timeoutMs) {
    if (!isChangeTaskPage()) return true;
    if (isTaskClosedNow()) return true;

    await sleep(pollMs);
  }
  return false;
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
      'select#change_task\\.close_code, select[name="change_task.close_code"], select[id$=".close_code"], select[name$=".close_code"]',
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
      'textarea#change_task\\.close_notes, textarea[name="change_task.close_notes"], textarea[id$=".close_notes"], textarea[name$=".close_notes"]',
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
  const t0 = Date.now();
  while (Date.now() - t0 < 10000) {
    const btn = findCloseTaskButton();
    if (btn) {
      try {
        btn.focus();
      } catch {}
      btn.click();
      return true;
    }
    await sleep(120);
  }
  return false;
}

function getActionHaystack(el: HTMLElement): string {
  const haystack = [
    el.id || "",
    el.getAttribute("name") || "",
    el.getAttribute("data-action-name") || "",
    (el as HTMLInputElement).value || "",
    el.getAttribute("onclick") || "",
    el.getAttribute("href") || "",
    el.getAttribute("title") || "",
    el.getAttribute("aria-label") || "",
    normalize(el.textContent || ""),
  ]
    .join(" ")
    .toLowerCase();
  return haystack;
}

function isCloseTaskLikeElement(el: HTMLElement): boolean {
  const haystack = getActionHaystack(el);
  return (
    haystack.includes("change_task_to_closed") ||
    haystack.includes("to_closed") ||
    haystack.includes("close_task") ||
    /\bclose\s*task\b/i.test(haystack) ||
    /\bclose\s*complete\b/i.test(haystack)
  );
}

function isClickable(el: HTMLElement): boolean {
  if ((el as HTMLButtonElement).disabled) return false;
  if (el.getAttribute("aria-disabled") === "true") return false;
  const rect = el.getBoundingClientRect();
  if (rect.width < 2 || rect.height < 2) return false;
  const style = getComputedStyle(el);
  if (style.display === "none" || style.visibility === "hidden") return false;
  return true;
}

function closeTaskScore(el: HTMLElement): number {
  const haystack = getActionHaystack(el);
  let score = 0;
  if (haystack.includes("change_task_to_closed")) score += 100;
  if (haystack.includes("to_closed")) score += 60;
  if (haystack.includes("close_task")) score += 50;
  if (/\bclose\s*task\b/i.test(haystack)) score += 40;
  if (/\bclose\s*complete\b/i.test(haystack)) score += 30;
  return score;
}

function findCloseTaskButton(): HTMLElement | null {
  const exact = document.querySelector<HTMLElement>(
    '#change_task_to_closed_bottom, #change_task_to_closed, button[data-action-name="change_task_to_closed"], input[data-action-name="change_task_to_closed"], [id*="change_task_to_closed"], [name*="change_task_to_closed"]'
  );
  if (exact && isClickable(exact) && isCloseTaskLikeElement(exact)) return exact;

  const candidates = Array.from(
    document.querySelectorAll<HTMLElement>(
      'button, input[type="button"], input[type="submit"], a[role="button"], a.button, a.btn'
    )
  )
    .filter(isClickable)
    .map((el) => ({ el, score: closeTaskScore(el) }))
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score);
  return candidates[0]?.el || null;
}

function isCloseTaskTrigger(target: EventTarget | null): boolean {
  const el =
    target instanceof Element
      ? target.closest<HTMLElement>(
          'button, a[role="button"], input[type="button"], input[type="submit"], a.button, a.btn'
        )
      : null;
  if (!el) return false;
  return isCloseTaskLikeElement(el);
}

function armManualCloseWatcher() {
  if ((window as any).__ciUpdaterManualCloseWatcherArmed) return;
  (window as any).__ciUpdaterManualCloseWatcherArmed = true;

  const onClick = (ev: Event) => {
    if (!isCloseTaskTrigger(ev.target)) return;
    setTimeout(() => {
      void (async () => {
        const confirmed = await waitForCloseConfirmation(20000, 250);
        if (!confirmed) return;
        const notified = await notifyCloseTaskDone(true, "manual_close_click");
        if (!notified) return;
        (window as any).__ciUpdaterCloseTaskDone = true;
        clearCloseTaskAction();
        document.removeEventListener("click", onClick, true);
        log("Manual Close Task click detected; proceed to next CI");
      })();
    }, 120);
  };
  document.addEventListener("click", onClick, true);
}

async function notifyCloseTaskDone(
  success: boolean,
  reason?: string
): Promise<boolean> {
  try {
    const { ciUpdaterClosing, ciUpdaterRunId } = await chrome.storage.local.get([
      "ciUpdaterClosing",
      "ciUpdaterRunId",
    ]);
    const resumeIndex = Number(ciUpdaterClosing?.resumeIndex);
    if (!Number.isFinite(resumeIndex)) return false;
    await chrome.runtime.sendMessage({
      type: "CLOSE_TASK_DONE",
      runId: ciUpdaterRunId || ciUpdaterClosing?.runId,
      resumeIndex,
      success,
      reason,
    });
    return true;
  } catch {}
  return false;
}

function armPassiveCloseWatcher() {
  if ((window as any).__ciUpdaterPassiveCloseWatcherArmed) return;
  (window as any).__ciUpdaterPassiveCloseWatcherArmed = true;

  const timer = setInterval(() => {
    void (async () => {
      if ((window as any).__ciUpdaterCloseTaskDone) {
        clearInterval(timer);
        return;
      }
      if (!isChangeTaskPage()) return;
      if (!isTaskClosedNow()) return;
      const notified = await notifyCloseTaskDone(true, "passive_state_closed");
      if (!notified) return;
      (window as any).__ciUpdaterCloseTaskDone = true;
      clearCloseTaskAction();
      clearInterval(timer);
      log("Passive close watcher confirmed closed task");
    })();
  }, 1200);
}

(async function run() {
  try {
    if (!isChangeTaskPage()) return;
    if ((window as any).__ciUpdaterCloseTaskDone) return;

    const action = readCloseTaskAction();
    if (!action) return;
    armPassiveCloseWatcher();

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

    await sleep(200);
    const clicked = await clickCloseTaskButton();
    await sleep(100);
    if (!clicked) {
      await notifyCloseTaskDone(false, "close_task_button_missing");
      armManualCloseWatcher();
      log("Close Task button not found; hold workflow until task is closed");
      return;
    }
    const confirmed = await waitForCloseConfirmation();
    if (!confirmed) {
      await notifyCloseTaskDone(false, "close_task_not_confirmed");
      armManualCloseWatcher();
      log("Close Task clicked but closure not confirmed; hold workflow");
      return;
    }
    const notified = await notifyCloseTaskDone(true);
    if (!notified) {
      log("Unable to notify close completion; keep close action for retry");
      return;
    }
    (window as any).__ciUpdaterCloseTaskDone = true;
    clearCloseTaskAction();
    log("Close fields set and Close Task clicked");
  } catch (e) {
    console.error("[CI Updater][Change-Task] error:", e);
  }
})();
