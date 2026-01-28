// src/content-outlook.ts
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

    const nestedAttr = el.querySelectorAll<HTMLElement>(
      '[title*="chg" i], [aria-label*="chg" i]'
    );
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
