// src/content-outlook.ts
import { parseTxtContent, showPageToast, sleep } from "./common";

type OutlookMessage = { type: "OUTLOOK_GET_EMAIL" };

function normalizeLine(line: string) {
  return (line || "")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/[\u00A0\u1680\u180E\u2000-\u200A\u202F\u205F\u3000]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function getLines(text: string): string[] {
  return (text || "")
    .replace(/\r/g, "\n")
    .split("\n")
    .map(normalizeLine);
}

function hasRequiredFields(bodyText: string, subjectText = "") {
  const combined = `${bodyText || ""}\n${subjectText || ""}`;
  return /CI-\d+/i.test(combined);
}

function stripLeadingNumber(line: string): string {
  return (line || "").replace(/^\s*\d+\s*[.)]?\s*/i, "").trim();
}

function findHeader(lines: string[]): string {
  const line = lines.find((l) => /Change\s+Account\s+User/i.test(l));
  if (line) return line;
  const chgLine = lines.find((l) => /^\s*CHG\d+\b/i.test(l));
  if (!chgLine) return "";
  const m = chgLine.match(/CHG\d+/i);
  return m ? m[0].toUpperCase() : stripLeadingNumber(chgLine);
}

function extractBetween(
  text: string,
  labelPattern: string,
  stopPatterns: string[]
): string {
  const labelRe = new RegExp(
    `(?:\\b\\d+\\s*[.)]?\\s*)?(?:${labelPattern})\\s*[:：]\\s*`,
    "i"
  );
  const match = labelRe.exec(text);
  if (!match) return "";
  const rest = text.slice(match.index + match[0].length);
  let end = rest.length;
  for (const stop of stopPatterns) {
    const stopRe = new RegExp(
      `\\b\\d+\\s*[.)]?\\s*(?:${stop})\\s*[:：]|\\b(?:${stop})\\s*[:：]`,
      "i"
    );
    const m2 = stopRe.exec(rest);
    if (m2 && m2.index < end) end = m2.index;
  }
  return normalizeLine(rest.slice(0, end));
}

function findValue(
  lines: string[],
  rawText: string,
  labelPattern: string,
  stopPatterns: string[]
): string {
  const lineRe = new RegExp(
    `^(?:\\d+\\s*[.)]?\\s*)?(?:${labelPattern})\\s*[:：]?\\s*(.*)$`,
    "i"
  );
  for (const line of lines) {
    const cleaned = stripLeadingNumber(line);
    const m = cleaned.match(lineRe);
    if (m && m[1]) return m[1].trim();
  }
  return extractBetween(rawText, labelPattern, stopPatterns);
}

function buildExtractedText(rawText: string, subjectText = ""): string {
  const cleanText = (rawText || "")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\r/g, "\n");
  const lines = getLines(cleanText);
  const flat = normalizeLine(cleanText);
  const combinedText = `${cleanText}\n${subjectText || ""}`;

  const chg = (combinedText.match(/CHG\d+/i)?.[0] || "").toUpperCase();
  const ciMatches = Array.from(combinedText.matchAll(/CI-\d+/gi)).map((m) =>
    (m[0] || "").toUpperCase()
  );
  const cis: string[] = [];
  for (const c of ciMatches) {
    if (!c) continue;
    if (cis.includes(c)) continue;
    cis.push(c);
  }

  const labelPatterns = {
    currentStatus: "Current\\s*Status|Install\\s*Status|Status",
    toClient: "To\\s*Client",
    contact: "Contact\\s*Name|Owned\\s*by|Owner\\s*by",
    location: "Location",
    otherDesc: "Other\\s*Desc\\.?|Other\\s*Description|Other|Note",
  };
  const allStops = [...Object.values(labelPatterns), "CI\\s*Description"];

  const countLabel = (pattern: string) => {
    const re = new RegExp(`^(?:${pattern})\\s*[:：]`, "i");
    let count = 0;
    for (const line of lines) {
      const cleaned = stripLeadingNumber(line);
      if (re.test(cleaned)) count += 1;
    }
    return count;
  };

  const hasRepeatedLabels =
    cis.length > 1 &&
    [
      labelPatterns.currentStatus,
      labelPatterns.contact,
      labelPatterns.location,
      labelPatterns.otherDesc,
      labelPatterns.toClient,
    ].some((p) => countLabel(p) > 1);

  if (hasRepeatedLabels) {
    const subjectLine = normalizeLine(subjectText);
    let raw = cleanText.trim();
    if (subjectLine && !raw.includes(subjectLine)) {
      raw = `${subjectLine}\n\n${raw}`;
    }
    return raw.trim();
  }

  const currentStatus = findValue(lines, cleanText, labelPatterns.currentStatus, allStops);
  const toClient = findValue(lines, cleanText, labelPatterns.toClient, allStops);
  const contact = findValue(lines, cleanText, labelPatterns.contact, allStops);
  const location = findValue(lines, cleanText, labelPatterns.location, allStops);
  const otherDesc = findValue(lines, cleanText, labelPatterns.otherDesc, allStops);

  const out: string[] = [];
  let header = findHeader(lines);
  if (!header && chg) header = chg;
  if (header) out.push(header, "");
  const headerHasChg =
    !!header && !!chg && header.toUpperCase().includes(chg.toUpperCase());
  if (chg && !headerHasChg) out.push(`Change #${chg}`, "");
  if (cis.length) {
    for (const c of cis) out.push(c);
    out.push("");
  }
  if (currentStatus) out.push(`Current Status : ${currentStatus}`);
  if (toClient) out.push(`To Client : ${toClient}`);
  if (contact) out.push(`Contact Name : ${contact}`);
  if (location) out.push(`Location : ${location}`);
  if (otherDesc) out.push(`Other Desc. : ${otherDesc}`);

  if (!out.length && hasRequiredFields(flat, subjectText)) return flat;
  return out.join("\n").trim();
}

function extractFromElement(el: Element | null, subjectText: string): string | null {
  if (!el) return null;
  const raw = (el as HTMLElement).innerText || "";
  if (!raw) return null;
  if (!hasRequiredFields(raw, subjectText)) return null;
  const t = buildExtractedText(raw, subjectText);
  return t || null;
}

function getSubjectText(): string {
  const pickChgText = (text: string | null | undefined) => {
    const t = normalizeLine(text || "");
    if (!t) return "";
    const m = t.match(/CHG\d+/i);
    if (!m) return "";
    if (t.length <= 200) return t;
    return m[0].toUpperCase();
  };

  const pickFromElement = (el: HTMLElement) => {
    const fromTitle = pickChgText(el.getAttribute("title"));
    if (fromTitle) return fromTitle;
    const fromAria = pickChgText(el.getAttribute("aria-label"));
    if (fromAria) return fromAria;

    const nestedAttr = Array.from(el.querySelectorAll<HTMLElement>(
      '[title*="chg" i], [aria-label*="chg" i]'
    ));
    for (const node of nestedAttr) {
      const nestedTitle = pickChgText(node.getAttribute("title"));
      if (nestedTitle) return nestedTitle;
      const nestedAria = pickChgText(node.getAttribute("aria-label"));
      if (nestedAria) return nestedAria;
      const nestedText = pickChgText(node.textContent);
      if (nestedText) return nestedText;
    }

    const fromText = pickChgText(el.textContent);
    if (fromText) return fromText;
    return "";
  };

  const selectors = [
    '[data-testid="message-subject"]',
    '[data-testid="readingPaneSubject"]',
    '[data-testid*="subject"]',
    '[id*="_SUBJECT"]',
    '[aria-label="Message subject"]',
    '[aria-label*="Subject"]',
    'div[role="heading"]',
    'span[role="heading"]',
    "h1",
  ];

  for (const sel of selectors) {
    const els = document.querySelectorAll<HTMLElement>(sel);
    for (const el of els) {
      const picked = pickFromElement(el);
      if (picked) return picked;
    }
  }

  const title = normalizeLine(document.title || "");
  if (/CHG\d+/i.test(title)) return title;

  return "";
}

function getMessageBodyText(): string | null {
  const subjectText = getSubjectText();
  const selectors = [
    'div[role="document"]',
    'div[aria-label="Message body"]',
    'div[aria-label^="Message body"]',
    'div[aria-label*="Message body"]',
  ];

  for (const sel of selectors) {
    const el = document.querySelector(sel);
    const t = extractFromElement(el, subjectText);
    if (t) return t;
  }

  // If running inside an iframe, fall back to the body text.
  if (window.top !== window) {
    const raw = document.body?.innerText || "";
    if (raw && hasRequiredFields(raw, subjectText)) return buildExtractedText(raw, subjectText);
  }

  return null;
}

type OutlookWatchOptions = {
  watch: boolean;
  autoRun: boolean;
  onlyUnread: boolean;
};

const OUTLOOK_STORAGE_KEYS = [
  "ciUpdaterOutlookWatch",
  "ciUpdaterOutlookAutoRun",
  "ciUpdaterOutlookOnlyUnread",
  "ciUpdaterOutlookProcessed",
] as const;

let outlookOptions: OutlookWatchOptions = {
  watch: false,
  autoRun: false,
  onlyUnread: true,
};

let processedIds = new Set<string>();
let observer: MutationObserver | null = null;
let scanTimer: number | null = null;
let scanPending = false;
let inProgressId: string | null = null;
let lastScanTs = 0;

async function loadOutlookOptions() {
  try {
    const res = await chrome.storage.local.get(OUTLOOK_STORAGE_KEYS as any);
    outlookOptions = {
      watch: Boolean(res.ciUpdaterOutlookWatch),
      autoRun: Boolean(res.ciUpdaterOutlookAutoRun),
      onlyUnread: res.ciUpdaterOutlookOnlyUnread === false ? false : true,
    };
    const stored = (res.ciUpdaterOutlookProcessed || []) as string[];
    processedIds = new Set(stored.filter(Boolean).slice(-200));
  } catch {}
}

async function saveProcessedIds() {
  try {
    const arr = Array.from(processedIds).slice(-200);
    await chrome.storage.local.set({ ciUpdaterOutlookProcessed: arr });
  } catch {}
}

function rememberProcessed(id: string) {
  if (!id) return;
  processedIds.add(id);
  void saveProcessedIds();
}

function normalizeText(text: string) {
  return (text || "").replace(/\s+/g, " ").trim().toLowerCase();
}

function getFolderSelectedText(): string {
  const folderRoot =
    document.querySelector('[data-testid="folderPane"]') ||
    document.querySelector('[aria-label="Folders"]') ||
    document.querySelector('[role="tree"]');
  if (!folderRoot) return "";
  const selected =
    folderRoot.querySelector<HTMLElement>('[role="treeitem"][aria-selected="true"]') ||
    folderRoot.querySelector<HTMLElement>('[aria-current="true"]') ||
    folderRoot.querySelector<HTMLElement>('.is-selected, .selected');
  return selected?.textContent?.trim() || "";
}

function isInUpdateCiFolder(): boolean {
  const selectedText = normalizeText(getFolderSelectedText());
  if (!selectedText) return true; // best effort: allow if cannot detect
  return selectedText.includes("update ci");
}

function getMessageListRoot(): ParentNode {
  return (
    document.querySelector('[data-testid="virtuoso-item-list"]') ||
    document.querySelector('[role="listbox"]') ||
    document
  );
}

function getMessageRows(): HTMLElement[] {
  const root = getMessageListRoot();
  const rows = Array.from(
    root.querySelectorAll<HTMLElement>('div[role="option"][data-convid], div[role="option"][data-item-index], div[role="option"]')
  );
  return rows;
}

function getRowTimeText(row: HTMLElement): string {
  const timeEl =
    row.querySelector<HTMLElement>("span._rWRU") ||
    row.querySelector<HTMLElement>('[title*=":"]') ||
    row.querySelector<HTMLElement>('span[aria-label*=":"]');
  const titled = timeEl?.getAttribute("title")?.trim() || "";
  if (titled) return titled;
  return timeEl?.textContent?.trim() || "";
}

function getRowId(row: HTMLElement): string {
  const base =
    row.getAttribute("data-convid") ||
    row.id ||
    row.getAttribute("data-item-index") ||
    "";
  const timeText = getRowTimeText(row);
  const subject = getRowSubject(row).slice(0, 120);
  const parts = [base || "row", timeText, subject].filter(Boolean);
  if (parts.length) return parts.join("|");
  const aria = row.getAttribute("aria-label") || "";
  return aria.slice(0, 200);
}

function getRowSubject(row: HTMLElement): string {
  const subjectEl =
    row.querySelector<HTMLElement>(".TtcXM") ||
    row.querySelector<HTMLElement>('[data-testid="message-subject"]') ||
    row.querySelector<HTMLElement>('[data-testid*="subject"]');
  const subject = subjectEl?.textContent?.trim();
  if (subject) return subject;
  const aria = row.getAttribute("aria-label") || "";
  return aria;
}

function getRowPreview(row: HTMLElement): string {
  const previewEl = row.querySelector<HTMLElement>(".FqgPc");
  return previewEl?.textContent?.trim() || "";
}

function isUpdateCiText(text: string): boolean {
  return /\bupdate\s*ci\b/i.test(text || "");
}

function parseRgb(color: string): [number, number, number] | null {
  const m = color.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
  if (!m) return null;
  return [Number(m[1]), Number(m[2]), Number(m[3])];
}

function isUnreadByColor(row: HTMLElement): boolean {
  const target =
    row.querySelector<HTMLElement>(".TtcXM") ||
    row.querySelector<HTMLElement>(".FqgPc");
  if (!target) return false;
  const color = getComputedStyle(target).color || "";
  const rgb = parseRgb(color);
  if (!rgb) return false;
  const [r, g, b] = rgb;
  return r > 150 && r - g > 40 && r - b > 40;
}

function isUnreadByWeight(row: HTMLElement): boolean {
  const target =
    row.querySelector<HTMLElement>(".TtcXM") ||
    row.querySelector<HTMLElement>(".FqgPc");
  if (!target) return false;
  const weight = Number(getComputedStyle(target).fontWeight);
  return Number.isFinite(weight) && weight >= 600;
}

function isUnreadRow(row: HTMLElement): boolean {
  const markReadBtn = row.querySelector(
    'button[title*="Mark as read" i], button[aria-label*="Mark as read" i], button[title*="อ่านแล้ว" i], button[aria-label*="อ่านแล้ว" i]'
  );
  if (markReadBtn) return true;
  const aria = row.getAttribute("aria-label") || "";
  if (/unread/i.test(aria)) return true;
  if (isUnreadByColor(row)) return true;
  if (isUnreadByWeight(row)) return true;
  return false;
}

function openRow(row: HTMLElement) {
  try {
    row.scrollIntoView({ block: "center" });
  } catch {}
  row.dispatchEvent(new MouseEvent("click", { bubbles: true }));
}

function getRowHints(row: HTMLElement) {
  const subject = getRowSubject(row);
  const preview = getRowPreview(row);
  const combined = `${subject} ${preview}`;
  const ci = combined.match(/CI-\d+/i)?.[0]?.toUpperCase() || "";
  const chg = combined.match(/CHG\d+/i)?.[0]?.toUpperCase() || "";
  return { subject, preview, ci, chg };
}

async function waitForMessageText(hintCi: string, timeoutMs = 8000) {
  const t0 = Date.now();
  while (Date.now() - t0 < timeoutMs) {
    const text = getMessageBodyText();
    if (text) {
      if (!hintCi || text.toUpperCase().includes(hintCi.toUpperCase())) return text;
    }
    await sleep(250);
  }
  return null;
}

async function startRunFromText(rawText: string) {
  const parsed = parseTxtContent(rawText);
  const hasCi = Boolean(parsed.ci || (parsed.cis && parsed.cis.length));
  if (!hasCi) {
    showPageToast("ไม่พบ CI ในเมลนี้", "error");
    return { ok: false, reason: "no_ci" };
  }
  await chrome.runtime.sendMessage({ type: "SET_RUNNING", value: true });
  const res = await chrome.runtime.sendMessage({ type: "RUN_UPDATE", data: parsed });
  if (res?.ok) {
    showPageToast("เริ่มอัปเดต CI แล้ว", "success");
    return { ok: true };
  }
  showPageToast("เริ่มอัปเดตไม่สำเร็จ", "error");
  return { ok: false, reason: res?.error || "run_failed" };
}

async function handleCandidate(row: HTMLElement) {
  const rowId = getRowId(row);
  if (!rowId) return;
  inProgressId = rowId;
  try {
    const { subject, preview, ci } = getRowHints(row);
    if (!outlookOptions.autoRun) {
      const ok = confirm(`พบ Update CI: ${subject || preview}\nเริ่มทำงานเลยไหม?`);
      if (!ok) {
        rememberProcessed(rowId);
        showPageToast("ข้ามเมลนี้แล้ว", "info");
        return;
      }
    }
    showPageToast("กำลังเปิดอีเมล Update CI...", "info");
    openRow(row);
    const text =
      (await waitForMessageText(ci, 9000)) ||
      buildExtractedText(preview, subject);
    if (!text || !hasRequiredFields(text, subject)) {
      showPageToast("ไม่พบข้อมูล Update CI จากเมลนี้", "error");
      rememberProcessed(rowId);
      return;
    }
    const result = await startRunFromText(text);
    if (result.ok) rememberProcessed(rowId);
  } finally {
    inProgressId = null;
  }
}

async function scanForUpdateEmail() {
  if (!outlookOptions.watch) return;
  if (inProgressId) return;
  const now = Date.now();
  if (now - lastScanTs < 600) return;
  lastScanTs = now;
  if (!isInUpdateCiFolder()) return;
  const { isRunning } = await chrome.storage.local.get("isRunning");
  if (isRunning === true) return;

  const rows = getMessageRows();
  for (const row of rows) {
    const subject = getRowSubject(row);
    if (!isUpdateCiText(subject)) continue;
    const id = getRowId(row);
    if (!id || processedIds.has(id) || id === inProgressId) continue;
    if (outlookOptions.onlyUnread && !isUnreadRow(row)) continue;
    await handleCandidate(row);
    break;
  }
}

function scheduleScan() {
  if (scanPending) return;
  scanPending = true;
  setTimeout(() => {
    scanPending = false;
    void scanForUpdateEmail();
  }, 450);
}

function startWatcher() {
  if (observer) return;
  const root = getMessageListRoot();
  observer = new MutationObserver(scheduleScan);
  observer.observe(root, { childList: true, subtree: true });
  if (scanTimer) clearInterval(scanTimer);
  scanTimer = window.setInterval(() => {
    void scanForUpdateEmail();
  }, 1500);
  void scanForUpdateEmail();
}

function stopWatcher() {
  if (observer) {
    observer.disconnect();
    observer = null;
  }
  if (scanTimer) {
    clearInterval(scanTimer);
    scanTimer = null;
  }
}

async function initOutlookWatch() {
  await loadOutlookOptions();
  if (!outlookOptions.watch) {
    stopWatcher();
    return;
  }
  startWatcher();
}

if (window.top === window) {
  void initOutlookWatch();
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== "local") return;
    const keys = Object.keys(changes || {});
    if (keys.some((k) => k.startsWith("ciUpdaterOutlook"))) {
      void loadOutlookOptions().then(() => {
        if (outlookOptions.watch) startWatcher();
        else stopWatcher();
      });
    }
  });
}

chrome.runtime.onMessage.addListener((msg: OutlookMessage, _sender, sendResponse) => {
  if (!msg || msg.type !== "OUTLOOK_GET_EMAIL") return;
  try {
    const text = getMessageBodyText();
    if (!text) return;
    sendResponse({ ok: true, text });
  } catch (e: any) {
    sendResponse({ ok: false, error: e?.message || String(e) });
  }
  return true;
});
