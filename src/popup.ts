import { parseTxtContent, ParsedData } from "./common";

const fileInput = document.getElementById("file") as HTMLInputElement;
const textInput = document.getElementById("textInput") as HTMLTextAreaElement;
const parseBtn = document.getElementById("parseBtn") as HTMLButtonElement;
const preview = document.getElementById("preview") as HTMLDivElement;
const runBtn = document.getElementById("runBtn") as HTMLButtonElement;
const stopBtn = document.getElementById("stopBtn") as HTMLButtonElement;
const resetBtn = document.getElementById("resetBtn") as HTMLButtonElement | null;
const statusEl = document.getElementById("status") as HTMLDivElement;
const testModeEl = document.getElementById("testMode") as HTMLInputElement | null;
const addCommentsEl = document.getElementById("addComments") as HTMLInputElement | null;
const timingBetweenEl = document.getElementById("timingBetween") as HTMLInputElement | null;
const timingInitialEl = document.getElementById("timingInitial") as HTMLInputElement | null;
const timingTableWaitEl = document.getElementById("timingTableWait") as HTMLInputElement | null;
const timingScanEl = document.getElementById("timingScan") as HTMLInputElement | null;
const timingObserverEl = document.getElementById("timingObserver") as HTMLInputElement | null;
const timingPollEl = document.getElementById("timingPoll") as HTMLInputElement | null;
const saveTimingBtn = document.getElementById("saveTimingBtn") as HTMLButtonElement | null;
const resetTimingBtn = document.getElementById("resetTimingBtn") as HTMLButtonElement | null;

const DEFAULT_TIMING = {
  initialDelayMs: 180,
  tableWaitBudgetMs: 2500,
  scanBudgetMs: 2500,
  pollMs: 80,
  observerBudgetMs: 3200,
  betweenCiDelayMs: 3000,
};

function showToast(message: string, type: "success" | "error" | "info" = "info") {
  const container = document.getElementById("toast-container")!;
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.textContent = message;

  container.appendChild(toast);

  // แสดงพร้อม animation
  setTimeout(() => toast.classList.add("show"), 50);

  // ซ่อนและลบออกอัตโนมัติ
  setTimeout(() => {
    toast.classList.remove("show");
    setTimeout(() => toast.remove(), 300);
  }, 3000);
}
let parsed: ParsedData | null = null;

function resetUiState() {
  parsed = null;
  if (fileInput) fileInput.value = "";
  if (textInput) textInput.value = "";
  preview.textContent = "";
  runBtn.disabled = true;
  statusEl.textContent = "พร้อมเริ่มรอบใหม่";
}

resetUiState();

function setTimingInput(
  el: HTMLInputElement | null,
  value: number | undefined,
  placeholder: number
) {
  if (!el) return;
  el.placeholder = String(placeholder);
  el.value = typeof value === "number" ? String(value) : "";
}

function readTimingInput(el: HTMLInputElement | null, min: number) {
  if (!el) return undefined;
  const raw = (el.value || "").trim();
  if (!raw) return undefined;
  const n = Number(raw);
  if (!Number.isFinite(n) || n < min) return undefined;
  return Math.round(n);
}

async function loadTimingInputs() {
  try {
    const { ciUpdaterTiming } = await chrome.storage.local.get("ciUpdaterTiming");
    const t = (ciUpdaterTiming || {}) as Record<string, number>;
    setTimingInput(timingBetweenEl, t.betweenCiDelayMs, DEFAULT_TIMING.betweenCiDelayMs);
    setTimingInput(timingInitialEl, t.initialDelayMs, DEFAULT_TIMING.initialDelayMs);
    setTimingInput(timingTableWaitEl, t.tableWaitBudgetMs, DEFAULT_TIMING.tableWaitBudgetMs);
    setTimingInput(timingScanEl, t.scanBudgetMs, DEFAULT_TIMING.scanBudgetMs);
    setTimingInput(timingObserverEl, t.observerBudgetMs, DEFAULT_TIMING.observerBudgetMs);
    setTimingInput(timingPollEl, t.pollMs, DEFAULT_TIMING.pollMs);
  } catch {}
}

void loadTimingInputs();

// init test mode checkbox from storage
(async () => {
  try {
    if (!testModeEl) return;
    const { ciUpdaterTestMode } = await chrome.storage.local.get("ciUpdaterTestMode");
    testModeEl.checked = Boolean(ciUpdaterTestMode);
  } catch {}
})();

// persist on change
testModeEl?.addEventListener("change", async () => {
  try {
    await chrome.storage.local.set({ ciUpdaterTestMode: !!testModeEl.checked });
  } catch {}
});

// init add comments toggle (default true)
(async () => {
  try {
    if (!addCommentsEl) return;
    const { ciUpdaterAddComments } = await chrome.storage.local.get("ciUpdaterAddComments");
    addCommentsEl.checked = ciUpdaterAddComments === false ? false : true;
  } catch {}
})();

// persist on change
addCommentsEl?.addEventListener("change", async () => {
  try {
    await chrome.storage.local.set({ ciUpdaterAddComments: !!addCommentsEl.checked });
  } catch {}
});

saveTimingBtn?.addEventListener("click", async () => {
  try {
    const timing: Record<string, number> = {};
    const between = readTimingInput(timingBetweenEl, 500);
    const initial = readTimingInput(timingInitialEl, 0);
    const tableWait = readTimingInput(timingTableWaitEl, 500);
    const scan = readTimingInput(timingScanEl, 500);
    const observer = readTimingInput(timingObserverEl, 500);
    const poll = readTimingInput(timingPollEl, 30);

    if (between != null) timing.betweenCiDelayMs = between;
    if (initial != null) timing.initialDelayMs = initial;
    if (tableWait != null) timing.tableWaitBudgetMs = tableWait;
    if (scan != null) timing.scanBudgetMs = scan;
    if (observer != null) timing.observerBudgetMs = observer;
    if (poll != null) timing.pollMs = poll;

    if (Object.keys(timing).length === 0) {
      await chrome.storage.local.remove("ciUpdaterTiming");
      showToast("ใช้ค่า Timing เริ่มต้นแล้ว", "success");
    } else {
      await chrome.storage.local.set({ ciUpdaterTiming: timing });
      showToast("บันทึก Timing แล้ว", "success");
    }
    await loadTimingInputs();
  } catch {
    showToast("บันทึก Timing ไม่สำเร็จ", "error");
  }
});

resetTimingBtn?.addEventListener("click", async () => {
  try {
    await chrome.storage.local.remove("ciUpdaterTiming");
    await loadTimingInputs();
    showToast("รีเซ็ต Timing เป็นค่าเริ่มต้นแล้ว", "success");
  } catch {
    showToast("รีเซ็ต Timing ไม่สำเร็จ", "error");
  }
});

function renderPreview(pd: ParsedData) {
  const cis = (pd.cis && pd.cis.length ? pd.cis : (pd.ci ? [pd.ci] : []));
  preview.textContent = [
    `Header: ${pd.header}`,
    `Mode: ${pd.mode}`,
    `CHG: ${pd.chg}`,
    `CI(s): ${cis.join(', ')}`,
    `Current Status: ${pd.currentStatus}`,
    `Contact Name (clean): ${pd.contact}`,
    `Location: ${pd.location}`,
    `Note: ${pd.otherDesc}`
  ].join("\n");

  const valid = Boolean((cis.length > 0) && pd.chg);
  runBtn.disabled = !valid;
  statusEl.textContent = valid ? "พร้อมอัปเดต" : "กรุณาตรวจสอบ CHG และ CI ในข้อมูล";
}

fileInput.addEventListener("change", async (e) => {
  const file = (e.target as HTMLInputElement).files?.[0];
  if (!file) return;
  const text = await file.text();
  parsed = parseTxtContent(text);
  renderPreview(parsed);
});

parseBtn.addEventListener("click", () => {
  const text = (textInput?.value || "").trim();
  if (!text) {
    preview.textContent = "";
    runBtn.disabled = true;
    statusEl.textContent = "กรุณาใส่ข้อความเพื่อตรวจสอบ";
    return;
  }
  parsed = parseTxtContent(text);
  renderPreview(parsed);
});

runBtn.addEventListener("click", async () => {
  if (!parsed) return;
  // persist current test mode before running
  if (testModeEl) {
    try { await chrome.storage.local.set({ ciUpdaterTestMode: !!testModeEl.checked }); } catch {}
  }
  // persist add comments setting before running
  if (addCommentsEl) {
    try { await chrome.storage.local.set({ ciUpdaterAddComments: !!addCommentsEl.checked }); } catch {}
  }
  await chrome.runtime.sendMessage({ type: "SET_RUNNING", value: true });
  showToast("กำลังเปิดหน้ารายการ Task CI...", "info");

  chrome.runtime.sendMessage({ type: "RUN_UPDATE", data: parsed }, (res) => {
    if (res?.ok) showToast("เปิดหน้ารายการแล้ว", "success");
    else showToast("เกิดข้อผิดพลาด: " + (res?.error || "unknown"), "error");
  });
});

stopBtn.addEventListener("click", async () => {
  await chrome.runtime.sendMessage({ type: "STOP_NOW" });
  showToast("⛔ หยุดและปิดแท็บที่เปิดโดย Extension แล้ว", "error");
  resetUiState();
});

// Reset data in extension storage except user options
resetBtn?.addEventListener("click", async () => {
  try {
    if (!confirm("ยืนยันล้างข้อมูลทั้งหมด (ยกเว้นตัวเลือก)?")) return;
    const KEEP_KEYS = ["ciUpdaterTestMode", "ciUpdaterAddComments", "ciUpdaterTiming"];
    const all = await chrome.storage.local.get(null as any);
    const keys = Object.keys(all || {});
    const removeKeys = keys.filter(k => !KEEP_KEYS.includes(k));
    if (removeKeys.length) await chrome.storage.local.remove(removeKeys);
    // ensure stopped state
    await chrome.storage.local.set({ isRunning: false });
    showToast("✅ ล้างข้อมูลเรียบร้อย (เก็บตัวเลือกไว้)", "success");
    resetUiState();
  } catch (e) {
    showToast("เกิดข้อผิดพลาดในการล้างข้อมูล", "error");
    console.error(e);
  }
});

chrome.storage.onChanged.addListener((changes, area) => {
  if (area !== "local") return;
  const change = changes.isRunning;
  if (change && change.oldValue === true && change.newValue === false) {
    resetUiState();
  }
});
