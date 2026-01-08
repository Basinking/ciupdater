// src/content-list.ts
import { sleep, showPageToast } from "./common";

type UpdaterData = { ci: string; chg: string; runId?: string };
type QueueData = { cis: string[]; index: number; runId?: string };
type Timing = {
  initialDelayMs: number;
  tableWaitBudgetMs: number;
  scanBudgetMs: number;
  pollMs: number;
  observerBudgetMs: number;
};

const DEFAULT_TIMING: Timing = {
  initialDelayMs: 180,     // หน่วงเริ่มต้นสั้น ๆ
  tableWaitBudgetMs: 2500, // เวลารอให้ "ตาราง" โผล่
  scanBudgetMs: 2500,      // เวลาสแกน/สังเกตแถวรวม
  pollMs: 80,              // ความถี่ polling
  observerBudgetMs: 3200,  // เพดานเวลาให้ MutationObserver ทำงาน
};

function isRealListPage(): boolean {
  // ทำงานเฉพาะหน้า list จริง (ใน iframe)
  return /\/task_ci_list\.do(?:$|\?)/i.test(location.pathname + location.search);
}

function normalize(s: string) {
  return (s || "").trim().replace(/\s+/g, " ");
}

function findRowMatch(tr: HTMLTableRowElement, ci: string, chg: string) {
  let ciText = "";
  let chgText = "";

  tr.querySelectorAll<HTMLAnchorElement>("a.linked, a.formlink, a").forEach(a => {
    const t = normalize(a.textContent || "");
    if (!ciText && /^CI-\d+/i.test(t)) ciText = t;
    if (!chgText && /^CHG\d+/i.test(t)) chgText = t;
  });

  if (!ciText) {
    const td = tr.querySelectorAll<HTMLTableCellElement>("td")[2];
    if (td) ciText = normalize(td.textContent || "");
  }
  if (!chgText) {
    const td = tr.querySelectorAll<HTMLTableCellElement>("td")[9];
    if (td) chgText = normalize(td.textContent || "");
  }

  const okCI  = ci ? ciText.toUpperCase().startsWith(ci.toUpperCase())   : true;
  const okCHG = chg ? chgText.toUpperCase().startsWith(chg.toUpperCase()) : true;
  return okCI && okCHG;
}

function getListTable(): HTMLTableElement | null {
  return (
    document.querySelector<HTMLTableElement>("#task_ci_table") ||
    document.querySelector<HTMLTableElement>("table.data_list_table.list_table") ||
    document.querySelector<HTMLTableElement>("table.list2_table")
  );
}

function parseCountAttr(tbl: HTMLTableElement): number | null {
  const v =
    tbl.getAttribute("grand_total_rows") ??
    tbl.getAttribute("total_rows") ??
    tbl.getAttribute("last_row");
  if (v == null) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

async function getTiming(): Promise<Timing> {
  const { ciUpdaterTiming } = await chrome.storage.local.get("ciUpdaterTiming");
  return { ...DEFAULT_TIMING, ...(ciUpdaterTiming || {}) };
}

// รอ “ตาราง” แบบเร็ว: ผสม polling + observer สำหรับการแทรกตารางใหม่
async function waitForTableFast(budgetMs: number, pollMs: number): Promise<HTMLTableElement | null> {
  // 1) ลอง polling สั้น ๆ ก่อน
  const t0 = Date.now();
  while (Date.now() - t0 < Math.min(300, budgetMs)) {
    const tbl = getListTable();
    if (tbl) return tbl;
    await sleep(pollMs);
  }

  // 2) ตั้ง observer รอดูการแทรก table แบบ real-time พร้อมเพดานเวลา
  return new Promise((resolve) => {
    let resolved = false;
    const kill = () => {
      if (resolved) return;
      resolved = true;
      resolve(null);
    };

    const obs = new MutationObserver(() => {
      const tbl = getListTable();
      if (tbl) {
        if (!resolved) {
          resolved = true;
          obs.disconnect();
          resolve(tbl);
        }
      }
    });

    obs.observe(document.documentElement, { childList: true, subtree: true });

    setTimeout(() => {
      if (!resolved) {
        obs.disconnect();
        // 3) ปิดท้ายด้วย polling รอบสุดท้าย (กัน observer พลาด)
        const t1 = Date.now();
        (async () => {
          while (Date.now() - t1 < Math.max(0, budgetMs - 300)) {
            const tbl = getListTable();
            if (tbl) return resolve(tbl);
            await sleep(pollMs);
          }
          kill();
        })();
      }
    }, 300);
  });
}

// สแกน/สังเกตการเพิ่ม “แถว” อย่างไว พร้อมเพดานเวลา
async function waitForMatchOrNoRows(
  table: HTMLTableElement,
  ci: string,
  chg: string,
  scanBudgetMs: number,
  pollMs: number,
  observerBudgetMs: number
): Promise<"match" | "no-rows" | "timeout"> {
  const tbody = table.tBodies?.[0] || table;

  // 1) Fast-path จาก attribute count ถ้ามี
  const count0 = parseCountAttr(table);
  if (count0 != null) {
    if (count0 > 0) return "match";   // มีแถวแล้ว ตาม query เดิม → ถือว่าเจอ
    if (count0 === 0) return "no-rows"; // ไม่มีแถวจริง → ไป Add ได้เลย
  }

  // 2) เช็คแถวที่มีอยู่ทันที
  const scanNow = () => {
    const rows = Array.from(
      tbody.querySelectorAll<HTMLTableRowElement>("tr.list_row, tr")
    ).filter(tr => tr.querySelectorAll("td").length > 0);
    if (rows.length === 0) return false;
    // เคร่งครัด: ต้องมีแถวที่ match CI/CHG จริงเท่านั้น
    return rows.some(tr => findRowMatch(tr, ci, chg));
  };
  if (scanNow()) return "match";

  // 3) ผสม observer + polling ภายในเพดานเวลา (ไวกว่าเดิม)
  return await new Promise<"match" | "no-rows" | "timeout">((resolve) => {
    let done = false;
    const finish = (v: "match" | "no-rows" | "timeout") => {
      if (done) return;
      done = true;
      obs.disconnect();
      resolve(v);
    };

    // a) observe การเพิ่ม/แก้ไขใน tbody
    const obs = new MutationObserver(() => {
      if (scanNow()) finish("match");
      // ถ้า attribute count เปลี่ยนเป็น 0 ให้ปิดงานไว
      const c = parseCountAttr(table);
      if (c === 0) finish("no-rows");
    });
    obs.observe(tbody, { childList: true, subtree: true, characterData: true, attributes: true });

    // b) polling สั้น ๆ ควบคู่ (กัน observer ไม่ยิง)
    const t0 = Date.now();
    (async () => {
      while (!done && Date.now() - t0 < scanBudgetMs) {
        if (scanNow()) return finish("match");
        const c = parseCountAttr(table);
        if (c === 0) return finish("no-rows");
        await sleep(pollMs);
      }
      // c) เผื่อ attribute count โผล่ช้า ให้รออีกนิดด้วย observerBudgetMs
      const t1 = Date.now();
      while (!done && Date.now() - t1 < observerBudgetMs) {
        const c = parseCountAttr(table);
        if (c === 0) return finish("no-rows");
        await sleep(pollMs);
      }
      finish("timeout");
    })();
  });
}

async function waitForTableWithBackoff(
  baseBudgetMs: number,
  pollMs: number
): Promise<HTMLTableElement | null> {
  const attempts = [
    { budgetMs: baseBudgetMs, delayMs: 0 },
    { budgetMs: Math.max(3000, Math.round(baseBudgetMs * 1.5)), delayMs: 250 },
    { budgetMs: Math.max(5000, Math.round(baseBudgetMs * 2.2)), delayMs: 450 },
  ];

  for (const a of attempts) {
    if (a.delayMs) await sleep(a.delayMs);
    const tbl = await waitForTableFast(a.budgetMs, pollMs);
    if (tbl) return tbl;
  }
  return null;
}

function requestListRetry(runId?: string, reason?: string) {
  try {
    chrome.runtime.sendMessage({ type: "REQUEST_LIST_RETRY", runId, reason });
  } catch {}
}

// --- Main -----------------------------------------------------------------
(async function run() {
  try {
    if (!isRealListPage()) return;

    const { isRunning, ciUpdaterRunId } = await chrome.storage.local.get([
      "isRunning",
      "ciUpdaterRunId",
    ]);
    if (isRunning === false) return;

    const { ciUpdaterData, ciUpdaterQueue: qRaw } = await chrome.storage.local.get([
      "ciUpdaterData",
      "ciUpdaterQueue",
    ]);
    const q = qRaw as QueueData | undefined;
    let data = (ciUpdaterData || {}) as UpdaterData;
    if (ciUpdaterRunId && data.runId && data.runId !== ciUpdaterRunId) return;
    if (ciUpdaterRunId && q?.runId && q.runId !== ciUpdaterRunId) return;

    // รอบแรกอาจเกิด race: หน้า list โผล่ก่อน storage เขียนค่าเสร็จ → รอสั้น ๆ ให้มีข้อมูล
    if (!data.ci && !(q?.cis && Array.isArray(q.cis) && q.cis.length > 0)) {
      const t0 = Date.now();
      while (Date.now() - t0 < 2800) {
        const g = await chrome.storage.local.get(["ciUpdaterData","ciUpdaterQueue"]);
        data = (g.ciUpdaterData || {}) as UpdaterData;
        if (data.ci || (g.ciUpdaterQueue?.cis && g.ciUpdaterQueue.cis.length > 0)) break;
        await sleep(100);
      }
    }

    // Show progress on list page as well
    try {
      const cis: string[] = (q?.cis && Array.isArray(q.cis)) ? q.cis : (data.ci ? [data.ci] : []);
      const total = Math.max(1, cis.length || 1);
      const index = Math.min(Math.max(0, q?.index ?? 0), total - 1);
      const current = index + 1;
      showPageToast(`CI ${current}/${total} • Finding in list...`, "info", 1800);
    } catch {}

    const timing = await getTiming();

    // เพิ่ม budget ให้ CI ลำดับแรก (มักโหลดช้ากว่าปกติ)
    let boost = 1;
    try {
      const isFirst = !q || ((q?.index ?? 0) === 0);
      if (isFirst) boost = 1.8;
    } catch {}
    const eff = {
      ...timing,
      initialDelayMs: Math.round(timing.initialDelayMs * boost),
      tableWaitBudgetMs: Math.round(timing.tableWaitBudgetMs * boost),
      scanBudgetMs: Math.round(timing.scanBudgetMs * boost),
      observerBudgetMs: Math.round(timing.observerBudgetMs * boost),
    };

    await sleep(eff.initialDelayMs);

    const table = await waitForTableWithBackoff(eff.tableWaitBudgetMs, eff.pollMs);
    if (!table) {
      // ไม่เห็นตารางเลย → อย่ารีบไป Add (กัน false positive)
      // ให้ user สั่งรันใหม่ หรือระบบจะเปิด Add เมื่อหน้า wrapper บอกชัดว่าไม่มี
      console.warn("[CI Updater] table not found; giving up this pass to avoid false Add");
      requestListRetry(ciUpdaterRunId || data.runId, "table_not_found");
      return;
    }

    // Fast-path: ถ้าตารางระบุ count แล้ว
    const count = parseCountAttr(table);
    if (count != null) {
      if (count > 0) {
        // ✅ ต้องตรวจให้ตรง CI/CHG จริงเท่านั้น
        try {
          const ok = openMatchedRow(data.ci || "", data.chg || "");
          if (!ok) {
            // แถวอาจยังไม่วาดครบ ให้รอรอบสั้น ๆ หา match ก่อนค่อยตัดสินใจ
            const quick = await waitForMatchOrNoRows(
              table,
              data.ci || "",
              data.chg || "",
              Math.min(1800, eff.scanBudgetMs),
              eff.pollMs,
              Math.min(1800, eff.observerBudgetMs)
            );
            if (quick === "match") {
              const ok2 = openMatchedRow(data.ci || "", data.chg || "");
              if (!ok2) chrome.runtime.sendMessage({ type: "OPEN_ADD_PAGE" });
            } else if (quick === "no-rows") {
              chrome.runtime.sendMessage({ type: "OPEN_ADD_PAGE" });
            } else {
              console.warn("[CI Updater] undecided after quick pass on count>0; skip to avoid false Add");
            }
          }
        } catch (e) {
          console.warn("[CI Updater] open row by count failed:", e);
        }
        return;
      } else {
        // ❌ 0 แถว ชัดเจน → ไป Add ไว
        chrome.runtime.sendMessage({ type: "OPEN_ADD_PAGE" });
        return;
      }
    }

    // ถ้ายังไม่ทราบ → race แบบเร็ว
    const result = await waitForMatchOrNoRows(
      table,
      data.ci || "",
      data.chg || "",
      eff.scanBudgetMs,
      eff.pollMs,
      eff.observerBudgetMs
    );

    if (result === "match") {
      // ✅ เจอแล้ว → คลิกเข้าเรคคอร์ดที่ตรงเท่านั้น
      try {
        const ok = openMatchedRow(data.ci || "", data.chg || "");
        if (!ok) {
          // หากไม่เจอแถวที่ตรง ให้ไปเพิ่มใหม่
          chrome.runtime.sendMessage({ type: "OPEN_ADD_PAGE" });
        }
      } catch (e) {
        console.warn("[CI Updater] open matched row failed:", e);
      }
      return;
    }
    if (result === "no-rows") {
      // ❌ ไม่เจอ → ไป Add
      const { isRunning: still } = await chrome.storage.local.get("isRunning");
      if (still === false) return;
      chrome.runtime.sendMessage({ type: "OPEN_ADD_PAGE" });
      return;
    }

    // timeout: ไม่ชัด → ลองสแกนรอบสุดท้ายแบบช้าลงอีกนิดก่อนยอมแพ้
    console.warn("[CI Updater] undecided (timeout); trying slow-pass scan");
    await sleep(250);
    const late = await waitForMatchOrNoRows(
      table,
      data.ci || "",
      data.chg || "",
      Math.max(2600, Math.round(eff.scanBudgetMs * 1.6)),
      eff.pollMs,
      Math.max(2600, Math.round(eff.observerBudgetMs * 1.6))
    );
    if (late === "match") {
      try {
        const ok = openMatchedRow(data.ci || "", data.chg || "");
        if (!ok) chrome.runtime.sendMessage({ type: "OPEN_ADD_PAGE" });
      } catch (e) {
        console.warn("[CI Updater] open matched row (late) failed:", e);
      }
      return;
    }
    if (late === "no-rows") {
      const { isRunning: still } = await chrome.storage.local.get("isRunning");
      if (still !== false) chrome.runtime.sendMessage({ type: "OPEN_ADD_PAGE" });
      return;
    }
    console.warn("[CI Updater] still undecided after late-pass; skip to avoid false Add");
    requestListRetry(ciUpdaterRunId || data.runId, "timeout_undecided");
  } catch (e) {
    console.error("[CI Updater] content-list error:", e);
    requestListRetry(ciUpdaterRunId, "exception");
  }
})();

function openMatchedRow(ci: string, chg: string): boolean {
  const tbl = getListTable();
  if (!tbl) return false;
  const tbody = (tbl.tBodies && tbl.tBodies[0]) || tbl;
  const rows = Array.from(
    tbody.querySelectorAll<HTMLTableRowElement>("tr.list_row, tr")
  ).filter(tr => tr.querySelectorAll("td").length > 0);

  const match = rows.find(tr => findRowMatch(tr, ci, chg));
  if (!match) return false;
  const link = getCILinkFromRow(match, ci) || match.querySelector<HTMLAnchorElement>("a.linked, a.formlink, a[href]");
  if (!link) return false;
  (link as HTMLAnchorElement).click();
  return true;
}

function openFirstRowLink(): boolean {
  const tbl = getListTable();
  if (!tbl) return false;
  const tbody = (tbl.tBodies && tbl.tBodies[0]) || tbl;
  const firstRow = Array.from(
    tbody.querySelectorAll<HTMLTableRowElement>("tr.list_row, tr")
  ).find(tr => tr.querySelectorAll("td").length > 0);
  if (!firstRow) return false;
  const link = getCILinkFromRow(firstRow, "") || firstRow.querySelector<HTMLAnchorElement>("a.linked, a.formlink, a[href]");
  if (!link) return false;
  link.click();
  return true;
}

function getCILinkFromRow(tr: HTMLTableRowElement, ci: string): HTMLAnchorElement | null {
  // เลือกลิงก์ที่ชี้ไปหน้า CI โดยเฉพาะ และถ้ามี ci ให้เทียบข้อความตรงตัว
  const anchors = Array.from(tr.querySelectorAll<HTMLAnchorElement>("a.linked, a.formlink, a[href]"));
  const norm = (s: string) => (s || "").trim().replace(/\s+/g, " ").toUpperCase();

  if (ci) {
    const wanted = norm(ci);
    const byText = anchors.find(a => norm(a.textContent || "") === wanted);
    if (byText) return byText;
    const byAria = anchors.find(a => norm(a.getAttribute("aria-label") || "").endsWith(wanted));
    if (byAria) return byAria;
  }

  // เผื่อกรณีไม่มีข้อความตรง แต่ href เป็น cmdb_ci.do (หรือ cmdb_ci_computer.do)
  // รองรับทุกหมวด: cmdb_ci*.do เช่น cmdb_ci_server.do, cmdb_ci_network_gear.do ฯลฯ
  const byHref = anchors.find(a => /cmdb_ci(?:_[a-z0-9_]+)?\.do/i.test(a.getAttribute("href") || ""));
  if (byHref) return byHref;

  return null;
}
