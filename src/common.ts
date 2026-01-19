export const sleep = (ms: number) => new Promise(res => setTimeout(res, ms));

export async function waitForElement<T extends Element>(
  selector: string,
  timeoutMs = 20000,
  pollMs = 200
): Promise<T> {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const el = document.querySelector<T>(selector);
    if (el) return el;
    await sleep(pollMs);
  }
  throw new Error(`Timeout waiting for ${selector}`);
}

export function setInputValue(el: HTMLInputElement, value: string) {
  const proto = Object.getPrototypeOf(el);
  const setter = Object.getOwnPropertyDescriptor(proto, "value")?.set;
  setter?.call(el, value);
  el.dispatchEvent(new Event("input", { bubbles: true }));
  el.dispatchEvent(new Event("change", { bubbles: true }));
}

export function cleanContactName(nameRaw: string): string {
  if (!nameRaw) return "";
  let s = nameRaw.trim();
  // ตัดอักขระนำหน้าเช่น ":", "：", "-", "–", "—", จุด bullet
  s = s.replace(/^[\s]*[:：\-–—•·]+\s*/u, "");
  // ตัดคำนำหน้าที่พบบ่อย (ไทย/อังกฤษย่อ)
  s = s.replace(/^(?:คุณ|นาย|นางสาว|นาง|น\.ส\.)\s*/u, "");
  s = s.replace(/^Mr\.?\s*/i, "");
  s = s.replace(/^Mrs\.?\s*/i, "");
  s = s.replace(/^Ms\.?\s*/i, "");
  s = s.replace(/^K[.\s]\s*/i, "");

  // เอาตัวเลขและสัญลักษณ์ออก ให้เหลือเฉพาะตัวอักษรอังกฤษและช่องว่าง
  s = s.replace(/[^A-Za-z\s]+/g, " ");
  s = s.replace(/\s+/g, " ").trim();

  // กรณีผู้ใช้กรอก IT Stock (ตัวเล็ก/ใหญ่ ไม่สำคัญ) โดยไม่มี RTH นำหน้า
  if (/^it\s*stock$/i.test(s)) {
    s = "RTH IT Stock";
  }

  return s.trim();
}

function normalizeCurrentStatus(statusRaw: string): string {
  let v = (statusRaw || "").trim();
  if (!v) return "";
  v = v.replace(/\s+/g, " ").trim();
  if (/\binstock\b/i.test(v) || /^in\s*stock$/i.test(v)) {
    v = v.replace(/\binstock\b/ig, "In Stock");
    if (/^in\s*stock$/i.test(v)) v = "In Stock";
  }
  return v;
}

function normalizeLocation(locationRaw: string): string {
  let v = (locationRaw || "").trim();
  if (!v) return "";
  v = v.replace(/\s+/g, " ").trim();
  const upper = v.toUpperCase();
  if (upper === "DHS-B1") return "DHS-B1-1";
  if (upper === "DHS-B2") return "DHS-B2-1";
  return v;
}

// Simple page toast (content-script friendly)
export function showPageToast(
  message: string,
  type: "info" | "success" | "error" = "info",
  durationMs = 2500
) {
  try {
    const doc = document;
    const id = "ci-updater-toast-container";
    let container = doc.getElementById(id) as HTMLDivElement | null;
    if (!container) {
      container = doc.createElement("div");
      container.id = id;
      container.style.position = "fixed";
      container.style.top = "12px";
      container.style.right = "12px";
      container.style.zIndex = "2147483647"; // on top
      container.style.display = "flex";
      container.style.flexDirection = "column";
      container.style.gap = "8px";
      doc.body.appendChild(container);
    }

    const toast = doc.createElement("div");
    toast.textContent = message;
    toast.style.padding = "10px 14px";
    toast.style.borderRadius = "8px";
    toast.style.fontSize = "13px";
    toast.style.color = "#fff";
    toast.style.boxShadow = "0 6px 18px rgba(0,0,0,.18)";
    toast.style.opacity = "0";
    toast.style.transform = "translateY(-6px)";
    toast.style.transition = "opacity .18s ease, transform .18s ease";
    toast.style.maxWidth = "360px";
    toast.style.wordBreak = "break-word";
    toast.style.borderLeft = "3px solid #e53935"; // red accent

    // Red/Black/White theme mapping
    if (type === "success") {
      toast.style.background = "#e53935"; // red
      toast.style.borderLeftColor = "#b71c1c";
    } else if (type === "error") {
      toast.style.background = "#000000"; // black
      toast.style.borderLeftColor = "#e53935";
    } else {
      toast.style.background = "#111111"; // dark gray/black
      toast.style.borderLeftColor = "#e53935";
    }

    container.appendChild(toast);
    requestAnimationFrame(() => {
      toast.style.opacity = "1";
      toast.style.transform = "translateY(0)";
    });
    setTimeout(() => {
      toast.style.opacity = "0";
      toast.style.transform = "translateY(-6px)";
      setTimeout(() => toast.remove(), 200);
    }, Math.max(1200, durationMs));
  } catch {}
}

export interface CiOverride {
  currentStatus?: string;
  toClient?: string;
  contact?: string;
  location?: string;
  otherDesc?: string;
}

export interface ParsedData {
  runId?: string;
  header: string;
  chg: string;
  mode: string;
  ci: string;
  cis?: string[];
  ciOverrides?: Record<string, CiOverride>;
  currentStatus: string;
  toClient: string;
  contact: string;
  location: string;
  otherDesc: string;
}

export function parseTxtContent(text: string): ParsedData {
  // ปรับช่องว่างทุกชนิดให้เป็น space ปกติ, ตัด zero-width
  const normalizeSpaces = (s: string) =>
    s
      .replace(/[\u200B-\u200D\uFEFF]/g, "")
      .replace(/[\u0009\u00A0\u1680\u180E\u2000-\u200A\u202F\u205F\u3000]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();

  const rawLines = text.split(/\r?\n/);
  const lines = rawLines.map(normalizeSpaces).filter(Boolean);
  const whole = normalizeSpaces(text);

  const labelDefs: Array<{
    key: keyof CiOverride;
    re: RegExp;
    normalize: (v: string) => string;
  }> = [
    {
      key: "currentStatus",
      re: /^(?:\d+\.\s*)?(?:Current\s*Status|Install\s*Status|Status)\s*[:：]\s*(.*)$/i,
      normalize: normalizeCurrentStatus,
    },
    {
      key: "toClient",
      re: /^(?:\d+\.\s*)?(?:To\s*Client)\s*[:：]\s*(.*)$/i,
      normalize: (v) => v.trim(),
    },
    {
      key: "contact",
      re: /^(?:\d+\.\s*)?(?:Contact\s*Name|Owned\s*by|Owner\s*by)\s*[:：]\s*(.*)$/i,
      normalize: cleanContactName,
    },
    {
      key: "location",
      re: /^(?:\d+\.\s*)?(?:Location)\s*[:：]\s*(.*)$/i,
      normalize: normalizeLocation,
    },
    {
      key: "otherDesc",
      re: /^(?:\d+\.\s*)?(?:Comments?|Other\s*Desc\.?|Other\s*Description|Other|Note)\s*[:：]\s*(.*)$/i,
      normalize: (v) => v.trim(),
    },
  ];

  const isSectionHeader = (line: string) => /^\d+\s*[.)]?$/.test(line);
  const extractCis = (line: string) =>
    Array.from(line.matchAll(/CI-\d+/gi)).map((m) => m[0].toUpperCase());
  const hasAnyField = (data: CiOverride | null | undefined) =>
    !!data &&
    Object.values(data).some(
      (v) => typeof v === "string" && v.trim().length > 0
    );

  const orderedCis: string[] = [];
  const addOrderedCi = (ciValue: string) => {
    if (!orderedCis.includes(ciValue)) orderedCis.push(ciValue);
  };

  const ciOverrides: Record<string, CiOverride> = {};
  let current: { cis: string[]; data: CiOverride } | null = null;
  let pending: CiOverride | null = null;

  const applyCurrentBlock = () => {
    if (!current) return;
    if (!current.cis.length || !hasAnyField(current.data)) {
      current = null;
      return;
    }
    for (const ciItem of current.cis) {
      const key = ciItem.toUpperCase();
      ciOverrides[key] = { ...(ciOverrides[key] || {}), ...current.data };
    }
    current = null;
  };

  const findField = (line: string): { key: keyof CiOverride; value: string } | null => {
    for (const def of labelDefs) {
      const m = line.match(def.re);
      if (!m) continue;
      const normalized = def.normalize(m[1] || "");
      if (!normalized) return null;
      return { key: def.key, value: normalized };
    }
    return null;
  };

  for (const line of lines) {
    if (isSectionHeader(line)) {
      applyCurrentBlock();
      current = null;
      pending = null;
      continue;
    }

    const cisInLine = extractCis(line);
    if (cisInLine.length) {
      if (current && hasAnyField(current.data)) {
        applyCurrentBlock();
      }
      if (!current) current = { cis: [], data: {} };
      if (pending && hasAnyField(pending) && !hasAnyField(current.data)) {
        current.data = { ...pending };
        pending = null;
      }
      for (const c of cisInLine) {
        if (!current.cis.includes(c)) current.cis.push(c);
        addOrderedCi(c);
      }
    }

    const field = findField(line);
    if (field) {
      if (!current) {
        pending = { ...(pending || {}), [field.key]: field.value };
      } else {
        current.data[field.key] = field.value;
      }
    }
  }
  applyCurrentBlock();

  // ดึง CHG จากทั้งเอกสาร รองรับ "Change #CHG0039650" หรือรูปแบบอื่น ๆ
  const chgMatch = whole.match(/CHG\d+/i);
  const chg = chgMatch ? chgMatch[0].toUpperCase() : "";

  // ดึง CI ทั้งหมดจากเอกสาร เช่น "CI-191003" รองรับหลายบรรทัด
  const ciAll = Array.from(whole.matchAll(/CI-\d+/gi)).map(m => m[0].toUpperCase());
  const cis: string[] = [...orderedCis];
  for (const ciItem of ciAll) {
    if (!cis.includes(ciItem)) cis.push(ciItem);
  }
  const ci = cis[0] || "";

  // helper: label ที่มีช่องว่างแปลก ๆ, เครื่องหมาย : หรือ ： และอาจมีเลขลำดับนำหน้า
  const getBy = (labelPattern: string) => {
    // ครอบ labelPattern ด้วย (?: ... ) เพื่อให้ | ทำงานเฉพาะในส่วน label และยังคง anchor ต้นบรรทัด
    const re = new RegExp(`^(?:\\d+\\.\\s*)?(?:${labelPattern})\\s*[:：]\\s*(.*)$`, "i");
    const found = lines.find(l => re.test(l));
    return found ? found.replace(re, "$1").trim() : "";
  };

  // ดึงค่าตามฟิลด์ โดยยืดหยุ่นกับรูปแบบที่แตกต่าง
  // ใช้ Install Status เป็นแหล่งข้อมูล Current Status ด้วย
  const currentStatusRaw = getBy("Current\\s*Status|Install\\s*Status|Status");
  const currentStatus = normalizeCurrentStatus(currentStatusRaw);
  const toClient = getBy("To\\s*Client");

  // Contact: อาจมาเป็น Contact Name, Owned by, Owner by
  const contactRaw =
    getBy("Contact\\s*Name") ||
    getBy("Owned\\s*by|Owner\\s*by");
  const contact = cleanContactName(contactRaw);

  const locationRaw = getBy("Location");
  const location = normalizeLocation(locationRaw);

  // Note/Comments/Other Desc → เก็บไว้ที่ otherDesc
  const comments =
    getBy("Comments?") ||
    getBy("Other\\s*Desc\\.?") ||
    getBy("Other") ||
    getBy("Note");
  const otherDesc = comments;

  // Mode: ถ้ามีบรรทัด Mode ให้ดึง, ถ้าไม่มีกับยังมี CI ให้ตั้งเป็น Update ตาม requirement
  const modeLine = getBy("Mode");
  let mode = modeLine || "";
  if (!mode && ci) mode = "Update";

  // Header: ถ้าเจอข้อความหัวเรื่องเดิม ใช้เลย, ไม่งั้นลองหา "Update CI-..." หรือใช้บรรทัดแรก
  const header =
    lines.find(l => /Change\s+Account\s+User/i.test(l)) ||
    lines.find(l => /\bUpdate\s+CI-\d+/i.test(l)) ||
    lines[0] || "";

  return {
    header,
    chg,
    mode,
    ci,
    cis,
    ciOverrides: Object.keys(ciOverrides).length > 0 ? ciOverrides : undefined,
    currentStatus,
    toClient,
    contact,
    location,
    otherDesc
  };
}
