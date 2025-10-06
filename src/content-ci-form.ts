// src/content-ci-form.ts
import {
  waitForElement,
  setInputValue,
  sleep,
  cleanContactName,
  showPageToast,
} from "./common";
const log = (...args: any[]) => console.log("[CI Updater][CI-Form]", ...args);

function isCIFormPage(): boolean {
  const href = location.href;
  // รองรับทั้งแบบโหลดโดยตรง และผ่าน classic wrapper เช่น
  // /now/nav/ui/classic/params/target/cmdb_ci.do?sys_id=... หรือ nav_to.do?uri=cmdb_ci.do?...
  return (
    /\/cmdb_ci(?:_[a-z0-9_]+)?\.do(?:$|\?|&)/i.test(href) ||
    /\buri=cmdb_ci(?:_[a-z0-9_]+)?\.do/i.test(href) ||
    !!getCIClassFromDom()
  );
}

function getCIClassFromDom(): string | null {
  const el = document.querySelector<HTMLElement>('div[id$=".form_scroll"]');
  if (!el) return null;
  const id = el.id || "";
  const m = id.match(/^(cmdb_ci(?:_[A-Za-z0-9_]+)?)\.form_scroll$/i);
  return m ? m[1] : null;
}

function getDetectedCIClass(): string | null {
  return getFormTableName() || getCIClassFromDom();
}

async function runCategoryExtras(ciClass: string, data: any) {
  const cls = (ciClass || "").toLowerCase();
  if (cls.includes("_ip_switch")) {
    await handleIpSwitch(data);
    return;
  }
  if (cls.includes("_ip_phone")) {
    await handleIpPhone(data);
    return;
  }
  if (cls.endsWith("_acc") || cls.includes("_acc")) {
    await handleAcc(data);
    return;
  }
}

// Disabled: no longer set Company from To Client
async function setCompanyFromToClient(_data: any) {
  return;
}

async function handleIpSwitch(data: any) {
  await setCompanyFromToClient(data);
}
async function handleIpPhone(data: any) {
  await setCompanyFromToClient(data);
}
async function handleAcc(data: any) {
  await setCompanyFromToClient(data);
}

function byIdSuffix<T extends Element>(suffix: string): T | null {
  return document.querySelector<T>(`[id$="${suffix}"]`);
}

function byIdPrefixSuffix<T extends Element>(
  prefix: string,
  suffix: string
): T | null {
  return document.querySelector<T>(`[id^="${prefix}"][id$="${suffix}"]`);
}

function mapInstallStatus(valueRaw: string): string | null {
  let v = (valueRaw || "").trim();
  // ตัดอักขระนำหน้าแปลก ๆ เช่น ":" หรือ "-" แล้ว normalize lower-case
  v = v
    .replace(/^[^A-Za-z0-9]+/, "")
    .trim()
    .toLowerCase();
  // Return the <option value> expected by ServiceNow for Install Status
  switch (v) {
    case "absent":
      return "100";
    case "installed":
      return "1";
    case "on order":
      return "2";
    case "in maintenance":
      return "3";
    case "pending install":
      return "4";
    case "pending repair":
      return "5";
    case "in stock":
      return "6";
    case "retired":
      return "7";
    case "stolen":
      return "8";
    default:
      return null; // ไม่รู้จักค่า → ไม่แตะต้อง
  }
}

function getFormTableName(): string | null {
  try {
    const gf = (window as any).g_form;
    if (gf && typeof gf.getTableName === "function") return gf.getTableName();
  } catch {}
  const target =
    (document.getElementById("sys_target") as HTMLInputElement | null)?.value ||
    "";
  if (target) return target;
  const m = location.href.match(/\/(\w+?)\.do(?:$|\?|&)/i);
  return m ? m[1] : null;
}

function findDisplayRef(fieldSuffix: string): HTMLInputElement | null {
  const tbl = getFormTableName();
  if (tbl) {
    const exactId = `sys_display.${tbl}.${fieldSuffix}`;
    const el = document.getElementById(exactId) as HTMLInputElement | null;
    if (el) return el;
  }
  return document.querySelector<HTMLInputElement>(
    `input[id^="sys_display."][id$=".${fieldSuffix}"]`
  );
}

function findHiddenRef(fieldSuffix: string): HTMLInputElement | null {
  // hidden value field (sys_id) for reference
  const tbl = getFormTableName();
  if (tbl) {
    const exactId = `${tbl}.${fieldSuffix}`;
    const el = document.getElementById(exactId) as HTMLInputElement | null;
    if (el) return el;
  }
  return document.querySelector<HTMLInputElement>(
    `input[id$=".${fieldSuffix}"]:not([id^="sys_display"])`
  );
}

function getRefMeta(fieldSuffix: string) {
  const disp = findDisplayRef(fieldSuffix);
  if (!disp) return null;
  const table =
    disp.getAttribute("data-table") ||
    disp.getAttribute("data-ref-table") ||
    "";
  const nameAttr = disp.getAttribute("name") || ""; // e.g. sys_display.cmdb_ci_computer.owned_by
  const fieldName = nameAttr.split(".").pop() || fieldSuffix; // owned_by
  return { disp, table, fieldName };
}

function fieldCandidates(
  ciClass: string | null,
  semantic: "location" | "owned_by"
): string[] {
  const cls = (ciClass || "").toLowerCase();
  // defaults
  let base: string[] = [];
  if (semantic === "location")
    base = [
      "location",
      "install_location",
      "site",
      "u_location",
      "u_site",
      "u_install_location",
    ];
  if (semantic === "owned_by")
    base = [
      "owned_by",
      "u_owned_by",
      "assigned_to",
      "u_assigned_to",
      "u_owner",
    ];
  // tweaks per class (extend if needed) — apply only for the relevant semantic
  if (semantic === "location") {
    if (cls.includes("_ip_switch")) {
      // some envs use network specific naming
      base = [
        "location",
        "install_location",
        "site",
        "u_site",
        "u_location",
        ...base,
      ];
    } else if (cls.includes("_ip_phone")) {
      base = ["location", "u_location", "site", ...base];
    } else if (cls.endsWith("_acc") || cls.includes("_acc")) {
      base = ["location", "u_location", "install_location", ...base];
    }
  }
  // remove duplicates while preserving order
  const seen = new Set<string>();
  return base.filter((s) => (seen.has(s) ? false : (seen.add(s), true)));
}

async function waitForAnyDisplayRef(
  suffixes: string[],
  timeoutMs = 8000,
  pollMs = 60
): Promise<{ disp: HTMLInputElement; used: string } | null> {
  const t0 = Date.now();
  while (Date.now() - t0 < timeoutMs) {
    for (const s of suffixes) {
      const el = findDisplayRef(s);
      if (el) return { disp: el, used: s };
    }
    await sleep(pollMs);
  }
  return null;
}

async function typeText(el: HTMLInputElement, text: string, stepMs = 80) {
  const valueSetter = Object.getOwnPropertyDescriptor(
    Object.getPrototypeOf(el),
    "value"
  )?.set;
  const send = (type: string, init: any = {}) =>
    el.dispatchEvent(
      new KeyboardEvent(type as any, { bubbles: true, ...init })
    );
  // clear first
  valueSetter?.call(el, "");
  el.dispatchEvent(new Event("input", { bubbles: true }));
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    send("keydown", { key: ch });
    valueSetter?.call(el, text.slice(0, i + 1));
    el.dispatchEvent(new Event("input", { bubbles: true }));
    send("keyup", { key: ch });
    await sleep(stepMs);
  }
}

function isVisible(el: HTMLElement): boolean {
  const rects = el.getClientRects();
  return !!rects && rects.length > 0;
}

function getFocusableList(): HTMLElement[] {
  const nodes = Array.from(
    document.querySelectorAll<HTMLElement>(
      "input, select, textarea, button, a[href], [tabindex]"
    )
  );
  return nodes.filter(
    (n) => !n.hasAttribute("disabled") && isVisible(n) && (n.tabIndex ?? 0) > -1
  );
}

function focusNextByTab(
  fromEl: HTMLElement,
  reverse = false
): HTMLElement | null {
  const list = getFocusableList();
  const idx = list.indexOf(fromEl);
  if (idx === -1) {
    const first = reverse ? list[list.length - 1] : list[0];
    first?.focus();
    return first || null;
  }
  const next = reverse ? list[idx - 1] : list[idx + 1];
  if (next) {
    next.focus();
    return next;
  }
  return null;
}

function sendTab(el: HTMLElement, reverse = false) {
  el.dispatchEvent(
    new KeyboardEvent("keydown", {
      key: "Tab",
      bubbles: true,
      shiftKey: reverse,
    })
  );
  focusNextByTab(el, reverse);
  el.dispatchEvent(
    new KeyboardEvent("keyup", { key: "Tab", bubbles: true, shiftKey: reverse })
  );
}

async function tabExitEnterExit(el: HTMLElement, waitMs = 80) {
  try {
    (el as any).scrollIntoView?.({ block: "center", inline: "nearest" });
  } catch {}
  // exit
  el.focus();
  sendTab(el);
  await sleep(waitMs);
  // re-enter
  (el as HTMLElement).click();
  el.dispatchEvent(new Event("focus", { bubbles: false }));
  await sleep(waitMs);
  // exit again
  sendTab(el);
}

async function setReferenceFieldDirect(
  fieldSuffixOrList: string | string[],
  text: string,
  settleMs = 2000
): Promise<boolean> {
  const suffixes = Array.isArray(fieldSuffixOrList)
    ? fieldSuffixOrList
    : [fieldSuffixOrList];
  let disp: HTMLInputElement | null = null;
  let usedSuffix = "";
  // quick try
  for (const s of suffixes) {
    const el = findDisplayRef(s);
    if (el) {
      disp = el;
      usedSuffix = s;
      break;
    }
  }
  if (!disp) {
    const found = await waitForAnyDisplayRef(suffixes, 8000, 60);
    if (found) {
      disp = found.disp;
      usedSuffix = found.used;
    }
  }
  if (!disp) return false;
  const hidden = findHiddenRef(usedSuffix);
  const want = (text || "").trim();
  log(`setReferenceFieldDirect('${usedSuffix || suffixes[0]}')`, {
    want,
    dispId: disp?.id || "<none>",
    hiddenId: hidden?.id || "<none>",
  });

  // Clear previous hidden/display to avoid stale values
  try {
    const meta = getRefMeta(usedSuffix || suffixes[0]);
    const gf = (window as any).g_form;
    if (meta && gf && typeof gf.setValue === "function") {
      gf.setValue(meta.fieldName, "");
    }
  } catch {}
  if (hidden) {
    hidden.value = "";
    hidden.dispatchEvent(new Event("change", { bubbles: true }));
  }

  // Focus, set value, fire events, then leave field via Tab
  disp.focus();
  setInputValue(disp, want);
  disp.dispatchEvent(new Event("input", { bubbles: true }));
  disp.dispatchEvent(new Event("change", { bubbles: true }));
  sendTab(disp);

  // Re-enter the input then exit again to trigger UI policies/validators
  try {
    (disp as any).scrollIntoView?.({ block: "center", inline: "nearest" });
  } catch {}
  await sleep(80);
  disp.click();
  disp.dispatchEvent(new Event("focus", { bubbles: false }));
  await sleep(80);
  sendTab(disp);
  disp.dispatchEvent(new Event("change", { bubbles: true }));

  // Wait briefly for hidden to resolve if available
  if (hidden) {
    const t0 = Date.now();
    while (Date.now() - t0 < settleMs) {
      if ((hidden.value || "").length > 0) return true;
      await sleep(120);
    }
    return (hidden.value || "").length > 0;
  }
  return true;
}

// ใช้ g_form.setDisplayValue เพื่อให้ SN ดึง sys_id เอง (เร็วกว่าและเสถียรกว่า UI typing)
async function setReferenceFieldViaGForm(
  fieldSuffixOrList: string | string[],
  text: string,
  settleMs = 1800
): Promise<boolean> {
  const suffixes = Array.isArray(fieldSuffixOrList)
    ? fieldSuffixOrList
    : [fieldSuffixOrList];
  const want = (text || "").trim();
  if (!want) return false;

  let gf: any = null;
  try {
    gf = (window as any).g_form;
  } catch {}
  if (!gf || typeof gf.setDisplayValue !== "function") return false;

  for (const s of suffixes) {
    try {
      const ctl = gf.getControl?.(s);
      // ต้องเจอ control และต้องเป็นฟิลด์ที่ลงท้ายด้วย .<suffix> จริง ๆ
      const idOk =
        !!ctl &&
        ((ctl as any).id?.endsWith?.(`.${s}`) ||
          (ctl as any).name?.endsWith?.(`.${s}`));
      if (!ctl || !idOk) continue; // ฟิลด์นี้ไม่มีบนฟอร์มหรือระบุตัวผิด

      // เคลียร์ค่าก่อนเพื่อกันค้าง
      try {
        if (typeof gf.setValue === "function") gf.setValue(s, "");
      } catch {}

      gf.setDisplayValue(s, want);

      const hidden = findHiddenRef(s);
      const t0 = Date.now();
      while (Date.now() - t0 < settleMs) {
        if (hidden) {
          if ((hidden.value || "").length > 0) return true;
        } else if (typeof gf.getValue === "function") {
          const v = gf.getValue(s) as string;
          if ((v || "").length > 0) return true;
        }
        await sleep(80);
      }
    } catch {}
  }
  return false;
}

async function setReferenceField(
  fieldSuffixOrList: string | string[],
  text: string,
  settleMs = 2500
): Promise<boolean> {
  const suffixes = Array.isArray(fieldSuffixOrList)
    ? fieldSuffixOrList
    : [fieldSuffixOrList];
  let disp: HTMLInputElement | null = null;
  let hidden: HTMLInputElement | null = null;
  let usedSuffix = "";
  for (const s of suffixes) {
    disp = findDisplayRef(s);
    hidden = findHiddenRef(s);
    if (disp && hidden) {
      usedSuffix = s;
      break;
    }
  }
  if (!disp || !hidden) return false;
  const want = (text || "").trim();
  if (!want) return false;
  log(`setReferenceField('${usedSuffix || suffixes[0]}')`, {
    want,
    dispId: disp.id,
    hiddenId: hidden.id,
  });

  // Clear previous value (very important otherwise hidden already has value)
  try {
    const meta = getRefMeta(usedSuffix || suffixes[0]);
    const gf = (window as any).g_form;
    if (meta && gf && typeof gf.setValue === "function") {
      gf.setValue(meta.fieldName, "");
    }
  } catch {}
  hidden.value = "";
  hidden.dispatchEvent(new Event("change", { bubbles: true }));

  disp.focus();
  await typeText(disp, want, 70);

  // Ensure autocomplete is initialized and try to open the list
  try {
    const completerName = disp.getAttribute("data-completer") || ""; // e.g. AJAXTableCompleter
    const refName = disp.getAttribute("data-ref") || ""; // e.g. cmdb_ci_computer.owned_by
    const dependent = disp.getAttribute("data-dependent") || "";
    const dyn = disp.getAttribute("data-ref-dynamic") || "";
    const win: any = window as any;
    if (
      !(disp as any).ac &&
      completerName &&
      typeof win[completerName] === "function"
    ) {
      try {
        (disp as any).ac = new win[completerName](
          disp,
          refName,
          dependent,
          dyn
        );
        log(`Initialized ${completerName} for`, refName);
      } catch (e) {
        log("init completer error", e);
      }
    }
    if ((disp as any).ac && typeof (disp as any).ac.onFocus === "function") {
      try {
        (disp as any).ac.onFocus();
      } catch {}
    }
  } catch (e) {
    log("ensure AC error", e);
  }

  // พยายามเปิด dropdown และเลือกจากลิสต์ให้ตรงข้อความที่ต้องการ
  let ownsId = disp.getAttribute("aria-owns") || "";
  if (!ownsId) {
    const refName = disp.getAttribute("data-ref") || "";
    if (refName) ownsId = `AC.${refName}`; // ปรับตามรูปแบบของ ServiceNow
  }
  log(`aria-owns for ${usedSuffix || suffixes[0]}:`, ownsId || "<none>");
  if (ownsId) {
    const t0 = Date.now();
    let list: HTMLElement | null = null;
    while (Date.now() - t0 < 3000 && !list) {
      list = document.getElementById(ownsId);
      if (list) break;
      // กระตุ้น AC ให้แสดงลิสต์
      disp.dispatchEvent(
        new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true })
      );
      disp.dispatchEvent(new Event("input", { bubbles: true }));
      disp.dispatchEvent(
        new KeyboardEvent("keyup", { key: "ArrowDown", bubbles: true })
      );
      await sleep(80);
    }

    if (list) {
      log(`AC list ready for ${usedSuffix || suffixes[0]}`, {
        listId: list.id,
      });
      const norm = (s: string) =>
        (s || "").replace(/\s+/g, " ").trim().toLowerCase();
      const wantNorm = norm(want);

      const candidates = Array.from(
        list.querySelectorAll<HTMLElement>('[role="option"], li, a')
      );
      log(
        `AC candidates(${candidates.length}) for ${usedSuffix || suffixes[0]}`
      );

      let picked = false;
      // 1) exact match
      for (const el of candidates) {
        const txt = norm(el.textContent || "");
        if (txt === wantNorm) {
          (el as HTMLElement).dispatchEvent(
            new MouseEvent("mousedown", { bubbles: true })
          );
          (el as HTMLElement).dispatchEvent(
            new MouseEvent("click", { bubbles: true })
          );
          log(`picked (exact) '${txt}' for ${usedSuffix || suffixes[0]}`);
          picked = true;
          break;
        }
      }
      // 2) startsWith match (ขึ้นต้นด้วยชื่อที่ต้องการ)
      if (!picked) {
        const el = candidates.find((e) => {
          const txt = norm(e.textContent || "");
          return txt.startsWith(wantNorm);
        });
        if (el) {
          el.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
          el.dispatchEvent(new MouseEvent("click", { bubbles: true }));
          log(
            `picked (startsWith) '${(el.textContent || "").trim()}' for ${
              usedSuffix || suffixes[0]
            }`
          );
          picked = true;
        }
      }
      // 3) contains match (กรณีลิสต์แสดงฟิลด์อื่นต่อท้าย)
      if (!picked) {
        const el = candidates.find((e) =>
          norm(e.textContent || "").includes(wantNorm)
        );
        if (el) {
          el.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
          el.dispatchEvent(new MouseEvent("click", { bubbles: true }));
          log(
            `picked (contains) '${(el.textContent || "").trim()}' for ${
              usedSuffix || suffixes[0]
            }`
          );
          picked = true;
        }
      }

      // รอให้ hidden ได้ค่า sys_id ถ้าเลือกได้แล้ว และ display ใกล้เคียงที่ต้องการ
      if (picked) {
        const t1 = Date.now();
        while (Date.now() - t1 < settleMs) {
          const dispOk = (() => {
            const norm = (s: string) =>
              (s || "").replace(/\s+/g, " ").trim().toLowerCase();
            return norm(disp.value).includes(norm(want));
          })();
          if ((hidden.value || "").length > 0 && dispOk) return true;
          await sleep(100);
        }
        log(
          `timeout waiting hidden sys_id for ${
            usedSuffix || suffixes[0]
          } after pick`
        );
      }
    }
  }

  // Fallback: ใช้คีย์ลัด ArrowDown + Enter
  disp.dispatchEvent(
    new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true })
  );
  await sleep(150);
  disp.dispatchEvent(
    new KeyboardEvent("keydown", { key: "Enter", bubbles: true })
  );

  const t2 = Date.now();
  while (Date.now() - t2 < settleMs) {
    const norm = (s: string) =>
      (s || "").replace(/\s+/g, " ").trim().toLowerCase();
    if (
      (hidden.value || "").length > 0 &&
      norm(disp.value).includes(norm(want))
    )
      return true;
    await sleep(100);
  }
  log(
    `fallback failed to resolve ${
      usedSuffix || suffixes[0]
    } (no hidden value set)`
  );
  return (hidden.value || "").length > 0;
}

(async function run() {
  try {
    if (!isCIFormPage()) return;
    log("CI form detected", location.href);
    // ป้องกันรันซ้ำบนหน้าเดียวกัน (เช่น reload หลัง update)
    if ((window as any).__ciUpdaterFormDone) return;

    const { isRunning } = await chrome.storage.local.get("isRunning");
    if (isRunning === false) return;

    const { ciUpdaterData: data, ciUpdaterQueue: q } =
      await chrome.storage.local.get(["ciUpdaterData", "ciUpdaterQueue"]);
    if (!data) return;

    // Show progress toast (e.g., 2/13) on the target page
    try {
      const cis: string[] =
        q?.cis && Array.isArray(q.cis) ? q.cis : data.ci ? [data.ci] : [];
      const total = Math.max(1, cis.length || 1);
      const index = Math.min(Math.max(0, q?.index ?? 0), total - 1);
      const current = index + 1;
      const ciLabel = data.ci || cis[index] || "";
      showPageToast(
        `Processing CI ${current}/${total} ${ciLabel}`,
        "info",
        2500
      );
    } catch {}

    try {
      const ciClass = getDetectedCIClass();
      log("Detected CI class", ciClass || "<unknown>");
      if (ciClass) await runCategoryExtras(ciClass, data);
    } catch {}

    // 0) รอให้ปุ่ม Update โผล่ (ยืนยันว่าอยู่บนหน้าแบบฟอร์มจริง)
    const updateBtn = await waitForElement<HTMLButtonElement>(
      '#sysverb_update_bottom, button[value="sysverb_update"], #sysverb_update',
      15000,
      120
    );
    log("Update button ready", updateBtn?.id || "<unknown>");

    // 1) ตั้งค่า Install Status จาก Current Status (ถ้ามีฟิลด์นี้)
    try {
      const rawStatus = (data.currentStatus || "").toString();
      const mapped = mapInstallStatus(rawStatus);
      log("Install Status mapping", { rawStatus, mapped });
      if (mapped) {
        let applied = false;
        // Preferred: use g_form API (works even when DOM uses label wrappers)
        try {
          const gf: any = (window as any).g_form;
          if (gf && typeof gf.setValue === "function") {
            try {
              gf.setValue("install_status", "");
            } catch {}
            gf.setValue("install_status", mapped);
            const tS = Date.now();
            while (Date.now() - tS < 1500) {
              try {
                const cur =
                  typeof gf.getValue === "function"
                    ? (gf.getValue("install_status") as string)
                    : "";
                if ((cur || "") === mapped) {
                  applied = true;
                  break;
                }
              } catch {}
              await sleep(120);
            }
          }
        } catch {}

        // Fallback: direct select element (classic UI)
        if (!applied) {
          try {
            const selectStatus = await waitForElement<HTMLSelectElement>(
              'select[id$=".install_status"], select[name$=".install_status"]',
              8000,
              120
            );
            if (selectStatus) {
              selectStatus.focus();
              selectStatus.value = mapped;
              selectStatus.dispatchEvent(new Event("input", { bubbles: true }));
              selectStatus.dispatchEvent(
                new Event("change", { bubbles: true })
              );
              await tabExitEnterExit(selectStatus);
              selectStatus.dispatchEvent(
                new Event("change", { bubbles: true })
              );
              // ensure applied
              const tS2 = Date.now();
              while (Date.now() - tS2 < 1500 && selectStatus.value !== mapped) {
                await sleep(120);
              }
              applied = selectStatus.value === mapped;
            }
          } catch {}
        }

        if (applied) log("Install Status set");
        else log("Install Status not applied");
      }
    } catch (e) {
      log("Install Status error", e);
    }

    // 2) ใส่ Location จากไฟล์ (reference field ถ้ามี)
    try {
      if (data.location) {
        const locWant = (data.location || "").trim();
        const contactWant = cleanContactName((data.contact || "").trim());
        // กันสลับฟิลด์: ถ้า Location เท่ากับ Contact ให้ข้ามการตั้ง Location
        if (
          locWant &&
          contactWant &&
          locWant.toLowerCase() === contactWant.toLowerCase()
        ) {
          log("Skip Location because equals Contact", { locWant, contactWant });
        } else {
          log("Setting Location (UI)", locWant);
          let __cands = [
            "location",
            "install_location",
            "site",
            "u_location",
            "u_site",
            "u_install_location",
          ] as string[];
          try {
            __cands = fieldCandidates(getDetectedCIClass(), "location");
          } catch {}
          try {
            await waitForAnyDisplayRef(__cands, 8000, 60);
          } catch {}
          // Fast path: g_form API → direct UI set → AC fallback
          let ok = await setReferenceFieldViaGForm(__cands, locWant, 1800);
          if (!ok) {
            ok = await setReferenceFieldDirect(__cands, locWant, 2400);
          }
          if (!ok) {
            log("Location direct methods failed, try AC typing fallback");
            await sleep(100);
            ok = await setReferenceField(__cands, locWant, 2600);
          }
          log("Location result", ok);
        }
      }
    } catch (e) {
      log("Location error", e);
    }

    // 2.1) ใส่ Owned by จาก Contact (reference field ถ้ามี)
    try {
      const contact = (data.contact || "").trim();
      if (contact) {
        log("Setting Owned by (UI) from contact", contact);
        try {
          const __ci = getDetectedCIClass();
          const __c = fieldCandidates(__ci, "owned_by");
          await waitForAnyDisplayRef(__c, 8000, 60);
        } catch {}
        // ใช้ชื่อที่ได้รับมาเท่านั้น (ไม่สลับ/ไม่แปลงลำดับ)
        const cleaned = contact.replace(/\s+/g, " ").trim();
        const variants: string[] = [cleaned];

        log("Owned by variants", variants);
        let setOk = false;
        for (const v of variants) {
          log("Trying Owned by variant", v);
          try {
            var __ciClass = getDetectedCIClass();
          } catch {
            var __ciClass = null as any;
          }
          const __ownedCands = fieldCandidates(__ciClass, "owned_by");
          // Fast path: g_form API → direct UI set → AC fallback
          setOk = await setReferenceFieldViaGForm(__ownedCands, v, 1800);
          if (!setOk) {
            setOk = await setReferenceFieldDirect(__ownedCands, v, 2400);
          }
          if (!setOk) {
            log("Owned by direct methods failed, fallback to AC typing");
            await sleep(100);
            setOk = await setReferenceField(__ownedCands, v, 2600);
          }
          log("Variant result", { v, ok: setOk });
          if (setOk) break;
        }
      }
    } catch (e) {
      log("Owned by error", e);
    }

    // 3) Append Comments = ต่อท้ายขึ้นบรรทัดใหม่ด้วย `CHG + note` (ถ้ามีฟิลด์นี้)
    try {
      // อ่านตัวเลือกว่าต้องการให้เติม comments หรือไม่ (ค่าเริ่มต้น = เติม)
      try {
        const { ciUpdaterAddComments } = await chrome.storage.local.get(
          "ciUpdaterAddComments"
        );
        if (ciUpdaterAddComments === false) {
          log("Skip adding Comments due to setting");
          throw null; // ข้ามทั้งบล็อก
        }
      } catch (ignore) {}

      const note = (data.otherDesc || "").trim();
      const chg = (data.chg || "").trim();
      const extra = [chg, note].filter(Boolean).join(" ").trim();
      if (extra) {
        const comments = document.querySelector<HTMLTextAreaElement>(
          'textarea[id$=".comments"]'
        );
        if (comments) {
          const normalize = (s: string) =>
            (s || "").replace(/\s+/g, " ").trim();
          const addLine = extra; // candidate line to add

          // สร้างรายการบรรทัดเดิม + บรรทัดใหม่ แล้วลบรายการที่ซ้ำให้เหลือบรรทัดเดียว (รักษาลำดับแรกเจอก่อน)
          const existingLines = (comments.value || "")
            .replace(/\s+$/, "")
            .split(/\r?\n/)
            .map((l) => l.replace(/\s+$/, ""));
          const merged = [...existingLines, addLine].filter(Boolean);
          const seen = new Set<string>();
          const unique: string[] = [];
          for (const line of merged) {
            const key = normalize(line);
            if (!key) continue;
            if (seen.has(key)) continue;
            seen.add(key);
            unique.push(line.trim());
          }

          comments.value = unique.join("\n");
          comments.dispatchEvent(new Event("input", { bubbles: true }));
          comments.dispatchEvent(new Event("change", { bubbles: true }));
          await tabExitEnterExit(comments);
          comments.dispatchEvent(new Event("change", { bubbles: true }));

          // ensure the prepared line exists (ตรวจแบบ normalize)
          const want = normalize(addLine);
          const tC = Date.now();
          while (Date.now() - tC < 1000) {
            const has = (comments.value || "")
              .split(/\r?\n/)
              .some((l) => normalize(l) === want);
            if (has) break;
            await sleep(120);
          }
        }
      }
    } catch {}

    // 4) ปล่อย Update ตรวจสอบอัตโนมัติทุก ๆ background ประมาณ 1 นาที (รองรับ Test Mode)
    const { isRunning: still } = await chrome.storage.local.get("isRunning");
    if (still === false) return;
    (window as any).__ciUpdaterFormDone = true;

    let finishedNotified = false;
    const notifyFinished = async (): Promise<void> => {
      if (finishedNotified) return;
      finishedNotified = true;
      try {
        const { ciUpdaterQueue: q2 } = await chrome.storage.local.get([
          "ciUpdaterQueue",
        ]);
        const total = Math.max(1, q2?.cis?.length ?? 1);
        const idx = Math.max(0, q2?.index ?? 0);
        const cur = Math.min(idx + 1, total);
        showPageToast(`Updated CI ${cur}/${total}`, "success", 2200);
      } catch {}
      try {
        await chrome.runtime.sendMessage({ type: "FINISHED_ONE" });
        log("FINISHED_ONE sent");
      } catch (e) {
        log("FINISHED_ONE send error", e);
      }
    };

    if (updateBtn) {
      updateBtn.addEventListener(
        "click",
        () => {
          void notifyFinished();
        },
        { once: true }
      );
    }

    // กำลังทดสอบการอัปเดตอัตโนมัติในพื้นหลัง (สำหรับทดสอบ)
    await sleep(600);
    try {
      const { ciUpdaterTestMode } = await chrome.storage.local.get(
        "ciUpdaterTestMode"
      );
      if (ciUpdaterTestMode) {
        log("Test Mode enabled: skip clicking Update");
        showPageToast("Test Mode: ข้ามการ Update", "info", 2200);
        await notifyFinished();
        return; // กรณีทดสอบ: ข้ามการ Update ไม่ต้องคลิกจริง
      }
    } catch {}

    log("Click Update button");
    updateBtn?.click();
    await notifyFinished();
  } catch (e) {
    console.error("[CI Updater][CI-Form] error:", e);
  }
})();
