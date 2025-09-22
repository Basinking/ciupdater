// src/content-add.ts
import { waitForElement, setInputValue, sleep, showPageToast } from "./common";

function isRealAddPage(): boolean {
  return /\/task_ci\.do(?:$|\?)/i.test(location.pathname + location.search);
}

(async function run() {
  try {
    if (!isRealAddPage()) return;

    const { isRunning } = await chrome.storage.local.get("isRunning");
    if (isRunning === false) return;

    // ถ้ามีคำสั่งให้นำทางไปหน้า CI โดยตรง (หลัง submit) ให้ทำก่อน
    try {
      const { ciUpdaterGoToCI } = await chrome.storage.local.get("ciUpdaterGoToCI");
      if (ciUpdaterGoToCI && ciUpdaterGoToCI.sysId) {
        // เคลียร์ flag เพื่อป้องกันวนซ้ำ
        await chrome.storage.local.remove("ciUpdaterGoToCI");
        const sysId: string = ciUpdaterGoToCI.sysId;
        const url = `${location.origin}/cmdb_ci.do?sys_id=${encodeURIComponent(sysId)}`;
        location.href = url;
        return;
      }
    } catch {}

    const { ciUpdaterData: data, ciUpdaterQueue: q } = await chrome.storage.local.get(["ciUpdaterData","ciUpdaterQueue"]);
    if (!data) {
      console.warn("[CI Updater] ไม่มีข้อมูลใน storage (เปิดจาก popup ก่อน)");
      return;
    }

    // แสดงความคืบหน้าขณะกำลังเพิ่มความสัมพันธ์ CI
    try {
      const cis: string[] = (q?.cis && Array.isArray(q.cis)) ? q.cis : (data.ci ? [data.ci] : []);
      const total = Math.max(1, cis.length || 1);
      const index = Math.min(Math.max(0, q?.index ?? 0), total - 1);
      const current = index + 1;
      showPageToast(`Adding CI ${current}/${total} ${data.ci || cis[index] || ""}`, "info", 2200);
    } catch {}

    // Helpers for Tab navigation (simulate real keydown/keyup Tab like CI form)
    const isVisible = (el: HTMLElement) => {
      const rects = el.getClientRects();
      return !!rects && rects.length > 0;
    };
    const getFocusableList = (): HTMLElement[] => {
      const nodes = Array.from(document.querySelectorAll<HTMLElement>('input, select, textarea, button, a[href], [tabindex]'));
      return nodes.filter(n => !n.hasAttribute('disabled') && isVisible(n) && (n.tabIndex ?? 0) > -1);
    };
    const focusNextByTab = (fromEl: HTMLElement, reverse = false): HTMLElement | null => {
      const list = getFocusableList();
      const idx = list.indexOf(fromEl);
      if (idx === -1) { const first = reverse ? list[list.length - 1] : list[0]; first?.focus(); return first || null; }
      const next = reverse ? list[idx - 1] : list[idx + 1];
      if (next) { next.focus(); return next; }
      return null;
    };
    const sendTab = (el: HTMLElement, reverse = false) => {
      el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, shiftKey: reverse }));
      focusNextByTab(el, reverse);
      el.dispatchEvent(new KeyboardEvent('keyup', { key: 'Tab', bubbles: true, shiftKey: reverse }));
    };
    const tabExitEnterExit = async (el: HTMLElement, waitMs = 80) => {
      try { (el as any).scrollIntoView?.({ block: 'center', inline: 'nearest' }); } catch {}
      el.focus();
      sendTab(el);
      await sleep(waitMs);
      (el as HTMLElement).click();
      el.dispatchEvent(new Event('focus', { bubbles: false }));
      await sleep(waitMs);
      sendTab(el);
    };

    // Direct set for reference-like inputs on task_ci form
    const setRefDirect = async (displaySelector: string, hiddenSelector: string, value: string): Promise<boolean> => {
      const disp = await waitForElement<HTMLInputElement>(displaySelector);
      const hidden = document.querySelector<HTMLInputElement>(hiddenSelector);
      if (hidden) {
        hidden.value = "";
        hidden.dispatchEvent(new Event("change", { bubbles: true }));
      }
      disp.focus();
      setInputValue(disp, value);
      disp.dispatchEvent(new Event("input", { bubbles: true }));
      disp.dispatchEvent(new Event("change", { bubbles: true }));
      await tabExitEnterExit(disp);
      if (!hidden) return true;
      const t0 = Date.now();
      while (Date.now() - t0 < 2000) {
        if ((hidden.value || "").length > 0) return true;
        await sleep(120);
      }
      return (hidden.value || "").length > 0;
    };

    // sys_display.task_ci.ci_item (Configuration Item), hidden #task_ci.ci_item
    await setRefDirect('input#sys_display\\.task_ci\\.ci_item', 'input#task_ci\\.ci_item', data.ci);

    // sys_display.task_ci.task (Task/CHG), hidden #task_ci.task
    await setRefDirect('input#sys_display\\.task_ci\\.task', 'input#task_ci\\.task', data.chg);

    let submitBtn =
      document.querySelector<HTMLButtonElement>("#sysverb_insert_bottom") ||
      document.querySelector<HTMLButtonElement>('button[type="submit"][value="sysverb_insert"]') ||
      document.querySelector<HTMLButtonElement>("button.form_action_button");

    if (!submitBtn) {
      submitBtn = await waitForElement<HTMLButtonElement>('button[type="submit"]');
    }

    const { isRunning: still } = await chrome.storage.local.get("isRunning");
    if (still === false) return;

    // เตรียมข้อมูลนำทางไปหน้า CI หลัง submit: ดึง sys_id จาก hidden ของฟิลด์ CI
    try {
      const hiddenCI = document.querySelector<HTMLInputElement>('input#task_ci\\.ci_item');
      const sysId = hiddenCI?.value || "";
      if (sysId) {
        await chrome.storage.local.set({ ciUpdaterGoToCI: { sysId, ts: Date.now() } });
      }
    } catch {}

    await sleep(120);
    submitBtn.click();
    console.log("[CI Updater] Submit done");

    // พยายามคลิกลิงก์ CI ทันทีถ้าเห็น (กรณีไม่เกิด reload)
    try {
      const ciLink = await waitForElement<HTMLAnchorElement>('a.linked[href*="cmdb_ci.do"], a[href*="cmdb_ci.do"]', 6000, 120);
      ciLink?.click();
      return;
    } catch {}
    // ถ้าไม่เจอ จะใช้กลไกนำทางตอนโหลดหน้า task_ci.do ถัดไป (ดูด้านบน)
  } catch (e) {
    console.error("[CI Updater] content-add error:", e);
  }
})();
