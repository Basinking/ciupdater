// src/background.ts
import type { ParsedData } from "./common";

type Messages =
  | { type: "RUN_UPDATE"; data: ParsedData }
  | { type: "OPEN_ADD_PAGE" }
  | { type: "STOP_NOW" }
  | { type: "SET_RUNNING"; value: boolean }
  | { type: "FINISHED_ONE"; runId?: string }
  | { type: "REQUEST_LIST_RETRY"; runId?: string; reason?: string };

let workerTabId: number | null = null;
const NEXT_CI_ALARM = "ci-updater-next-ci";
const LIST_RETRY_ALARM = "ci-updater-list-retry";
const RUN_STATE_KEYS = [
  "ciUpdaterQueue",
  "ciUpdaterBase",
  "ciUpdaterData",
  "ciUpdaterGoToCI",
  "ciUpdaterNext",
  "ciUpdaterRetry",
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
          await resetForNewRun();
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
        const runId = await ensureRunId();
        // If multiple CIs are provided, build the queue from that list
        if (Array.isArray(data.cis) && data.cis.length > 1) {
          const base = { ...data, runId } as any;
          delete base.ci;
          await chrome.storage.local.set({
            ciUpdaterQueue: { cis: data.cis, index: 0, runId },
            ciUpdaterBase: base,
          });
          await setCurrentCiIndex(0, runId);
        } else {
          await chrome.storage.local.set({ ciUpdaterData: { ...data, runId } });
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
        await handleFinishedOne(msg.runId);
        sendResponse({ ok: true });
        return;
      }

      if (msg.type === "REQUEST_LIST_RETRY") {
        await handleListRetryRequest(msg.runId, msg.reason);
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

chrome.alarms.onAlarm.addListener((alarm) => {
  (async () => {
    try {
      if (alarm.name === NEXT_CI_ALARM) {
        const { ciUpdaterNext, ciUpdaterRunId, isRunning } =
          await chrome.storage.local.get(["ciUpdaterNext", "ciUpdaterRunId", "isRunning"]);
        if (isRunning === false) return;
        if (!ciUpdaterNext || !ciUpdaterRunId) return;
        if (ciUpdaterNext.runId && ciUpdaterRunId !== ciUpdaterNext.runId) return;
        const nextIndex = Number(ciUpdaterNext.index);
        if (!Number.isFinite(nextIndex)) return;
        await setCurrentCiIndex(nextIndex, ciUpdaterRunId);
        await chrome.storage.local.remove("ciUpdaterNext");
      }

      if (alarm.name === LIST_RETRY_ALARM) {
        await handleListRetryAlarm();
      }
    } catch (e) {
      console.error(e);
    }
  })();
});

async function clearRunState() {
  try {
    await chrome.alarms.clear(NEXT_CI_ALARM);
    await chrome.alarms.clear(LIST_RETRY_ALARM);
  } catch {}
  try { await chrome.storage.local.remove(RUN_STATE_KEYS); } catch {}
}

async function clearRunMeta() {
  try { await chrome.storage.local.remove("ciUpdaterRunId"); } catch {}
}

function createRunId(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

async function resetForNewRun(): Promise<string> {
  await clearRunState();
  const runId = createRunId();
  await chrome.storage.local.set({ ciUpdaterRunId: runId });
  return runId;
}

async function ensureRunId(): Promise<string> {
  const { ciUpdaterRunId } = await chrome.storage.local.get("ciUpdaterRunId");
  if (ciUpdaterRunId) return ciUpdaterRunId as string;
  return resetForNewRun();
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

async function setCurrentCiIndex(index: number, runId?: string) {
  const { ciUpdaterQueue, ciUpdaterBase, ciUpdaterRunId } =
    await chrome.storage.local.get(["ciUpdaterQueue", "ciUpdaterBase", "ciUpdaterRunId"]);
  if (runId && ciUpdaterRunId && runId !== ciUpdaterRunId) return;
  if (ciUpdaterQueue?.runId && ciUpdaterRunId && ciUpdaterQueue.runId !== ciUpdaterRunId) return;
  const cis: string[] = ciUpdaterQueue?.cis || [];
  const chg: string = ciUpdaterBase?.chg || "";
  const base = ciUpdaterBase || {};
  const ci = cis[index] || "";
  const data = { ...base, ci, runId: ciUpdaterRunId || runId } as ParsedData;
  await chrome.storage.local.set({
    ciUpdaterData: data,
    ciUpdaterQueue: { cis, index, runId: ciUpdaterRunId || runId },
  });
  const listUrl = makeListUrl(ci, chg);
  await openOrReuseTab(listUrl);
}

async function handleFinishedOne(runId?: string) {
  const { isRunning, ciUpdaterRunId } = await chrome.storage.local.get(["isRunning", "ciUpdaterRunId"]);
  if (isRunning === false) return;
  if (runId && ciUpdaterRunId && runId !== ciUpdaterRunId) return;
  const { ciUpdaterQueue } = await chrome.storage.local.get("ciUpdaterQueue");
  const q = ciUpdaterQueue as { cis: string[]; index: number; runId?: string } | undefined;
  if (q?.runId && ciUpdaterRunId && q.runId !== ciUpdaterRunId) return;
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
  await scheduleNextCi(next, ciUpdaterRunId || runId);
}

async function setRunning(value: boolean) {
  await chrome.storage.local.set({ isRunning: value });
  if (!value) {
    await clearRunState();
    await clearRunMeta();
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

async function getBetweenCiDelayMs(): Promise<number> {
  try {
    const { ciUpdaterTiming } = await chrome.storage.local.get("ciUpdaterTiming");
    const v = Number(ciUpdaterTiming?.betweenCiDelayMs);
    if (Number.isFinite(v) && v >= 500) return v;
  } catch {}
  return 3000;
}

async function scheduleNextCi(index: number, runId?: string) {
  if (!runId) return;
  const delayMs = await getBetweenCiDelayMs();
  await chrome.storage.local.set({ ciUpdaterNext: { index, runId, ts: Date.now() } });
  chrome.alarms.create(NEXT_CI_ALARM, { when: Date.now() + delayMs });
}

async function getListRetryDelayMs(count: number): Promise<number> {
  const base = 1200;
  const step = 900;
  const cappedCount = Math.max(1, Math.min(4, count));
  return base + step * (cappedCount - 1);
}

async function handleListRetryRequest(runId?: string, reason?: string) {
  const {
    isRunning,
    ciUpdaterRunId,
    ciUpdaterRetry,
    ciUpdaterData,
  } = await chrome.storage.local.get([
    "isRunning",
    "ciUpdaterRunId",
    "ciUpdaterRetry",
    "ciUpdaterData",
  ]);
  if (isRunning === false) return;
  if (!ciUpdaterRunId) return;
  if (runId && ciUpdaterRunId !== runId) return;

  const data = ciUpdaterData as ParsedData | undefined;
  if (!data || !data.ci || !data.chg) return;

  const prev = (ciUpdaterRetry || {}) as { runId?: string; count?: number; lastTs?: number };
  const now = Date.now();
  if (prev.runId && prev.runId !== ciUpdaterRunId) {
    // reset counter across runs
    prev.count = 0;
    prev.lastTs = 0;
  }
  const count = Number(prev.count || 0) + 1;
  if (count > 3) {
    console.warn("[CI Updater] list retry limit reached", { runId: ciUpdaterRunId, reason });
    return;
  }
  if (prev.lastTs && now - prev.lastTs < 700) return; // avoid spam

  const delayMs = await getListRetryDelayMs(count);
  await chrome.storage.local.set({
    ciUpdaterRetry: { runId: ciUpdaterRunId, count, lastTs: now },
  });
  chrome.alarms.create(LIST_RETRY_ALARM, { when: now + delayMs });
}

async function handleListRetryAlarm() {
  const {
    ciUpdaterRunId,
    ciUpdaterRetry,
    ciUpdaterData,
    isRunning,
  } = await chrome.storage.local.get([
    "ciUpdaterRunId",
    "ciUpdaterRetry",
    "ciUpdaterData",
    "isRunning",
  ]);
  if (isRunning === false) return;
  if (!ciUpdaterRunId) return;
  const retry = (ciUpdaterRetry || {}) as { runId?: string };
  if (retry.runId && retry.runId !== ciUpdaterRunId) return;

  const data = ciUpdaterData as ParsedData | undefined;
  if (!data || !data.ci || !data.chg) return;

  const listUrl = makeListUrl(data.ci, data.chg);
  await openOrReuseTab(listUrl);
}

async function resetOnStartup() {
  try {
    await setRunning(false);
  } catch (e) {
    console.error(e);
  }
}

chrome.runtime.onInstalled.addListener(() => {
  void resetOnStartup();
});

chrome.runtime.onStartup.addListener(() => {
  void resetOnStartup();
});
