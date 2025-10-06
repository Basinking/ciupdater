// src/background.ts
import type { ParsedData } from "./common";

type Messages =
  | { type: "RUN_UPDATE"; data: ParsedData }
  | { type: "OPEN_ADD_PAGE" }
  | { type: "STOP_NOW" }
  | { type: "SET_RUNNING"; value: boolean }
  | { type: "FINISHED_ONE" };

let workerTabId: number | null = null;
const RUN_STATE_KEYS = [
  "ciUpdaterQueue",
  "ciUpdaterBase",
  "ciUpdaterData",
  "ciUpdaterGoToCI",
];

chrome.tabs.onRemoved.addListener((tabId) => {
  if (workerTabId === tabId) workerTabId = null;
});

chrome.runtime.onMessage.addListener((msg: Messages, _sender, sendResponse) => {
  (async () => {
    try {
      if (msg.type === "SET_RUNNING") {
        // When starting a new run, clear leftover queue/data from previous runs
        if (msg.value === true) {
          await clearRunState();
        }
        await setRunning(msg.value);
        sendResponse({ ok: true });
        return;
      }

      if (msg.type === "STOP_NOW") {
        await stopNow();
        sendResponse({ ok: true });
        return;
      }

      const { isRunning } = await chrome.storage.local.get("isRunning");
      if (isRunning === false) {
        sendResponse({ ok: false, error: "stopped" });
        return;
      }

      if (msg.type === "RUN_UPDATE") {
        const { data } = msg;
        // Before preparing new data, reset any stale queue/info
        await clearRunState();
        // If multiple CIs are provided, build the queue from that list
        if (Array.isArray(data.cis) && data.cis.length > 1) {
          const base = { ...data } as any;
          delete base.ci;
          await chrome.storage.local.set({
            ciUpdaterQueue: { cis: data.cis, index: 0 },
            ciUpdaterBase: base,
          });
          await setCurrentCiIndex(0);
        } else {
          await chrome.storage.local.set({ ciUpdaterData: data });
          const listUrl = makeListUrl(data.ci, data.chg);
          await openOrReuseTab(listUrl);
        }

        sendResponse({ ok: true });
        return;
      }

      if (msg.type === "OPEN_ADD_PAGE") {
        const addUrl = "https://ricohap.service-now.com/now/nav/ui/classic/params/target/task_ci.do";
        await openOrReuseTab(addUrl);
        sendResponse({ ok: true });
        return;
      }

      if (msg.type === "FINISHED_ONE") {
        await handleFinishedOne();
        sendResponse({ ok: true });
        return;
      }
    } catch (e: any) {
      console.error(e);
      sendResponse({ ok: false, error: e?.message || String(e) });
    }
  })();

  return true;
});

async function clearRunState() {
  try { await chrome.storage.local.remove(RUN_STATE_KEYS); } catch {}
}

function makeListUrl(ci: string, chg: string) {
  const base = "https://ricohap.service-now.com/now/nav/ui/classic/params/target/task_ci_list.do";
  const parts: string[] = [];
  if (ci) parts.push(`ci_item.nameSTARTSWITH${ci}`);
  if (chg) parts.push(`task.numberSTARTSWITH${chg}`);
  const params = new URLSearchParams({
    ...(parts.length ? { sysparm_query: parts.join("^") } : {}),
    sysparm_first_row: "1",
    sysparm_list_header_search: "true",
  });
  return `${base}?${params.toString()}`;
}

async function openOrReuseTab(url: string) {
  try {
    if (workerTabId != null) {
      await chrome.tabs.update(workerTabId, { url, active: true });
      return;
    }
  } catch {
    workerTabId = null; // Reset the worker tab if the previous one disappeared
  }
  const tab = await chrome.tabs.create({ url, active: true });
  if (tab.id != null) workerTabId = tab.id;
}

async function setCurrentCiIndex(index: number) {
  const { ciUpdaterQueue, ciUpdaterBase } = await chrome.storage.local.get(["ciUpdaterQueue", "ciUpdaterBase"]);
  const cis: string[] = ciUpdaterQueue?.cis || [];
  const chg: string = ciUpdaterBase?.chg || "";
  const base = ciUpdaterBase || {};
  const ci = cis[index] || "";
  const data = { ...base, ci } as ParsedData;
  await chrome.storage.local.set({ ciUpdaterData: data, ciUpdaterQueue: { cis, index } });
  const listUrl = makeListUrl(ci, chg);
  await openOrReuseTab(listUrl);
}

async function handleFinishedOne() {
  const { isRunning } = await chrome.storage.local.get("isRunning");
  if (isRunning === false) return;
  const { ciUpdaterQueue } = await chrome.storage.local.get("ciUpdaterQueue");
  const q = ciUpdaterQueue as { cis: string[]; index: number } | undefined;
  if (!q || !Array.isArray(q.cis) || q.cis.length === 0) {
    // single CI flow: just stop
    await setRunning(false);
    return;
  }
  const next = q.index + 1;
  if (next >= q.cis.length) {
    await setRunning(false);
    return;
  }
  // Delay briefly to let the form submit before moving to the next CI
  setTimeout(() => { setCurrentCiIndex(next); }, 1500);
}

async function setRunning(value: boolean) {
  await chrome.storage.local.set({ isRunning: value });
  if (!value) {
    await clearRunState();
  }
  chrome.action.setBadgeBackgroundColor({ color: value ? "#0a84ff" : "#777" });
  chrome.action.setBadgeText({ text: value ? "RUN" : "" });
}

async function stopNow() {
  await setRunning(false);
  if (workerTabId != null) {
    try { await chrome.tabs.remove(workerTabId); } catch {}
    workerTabId = null;
  }
}
