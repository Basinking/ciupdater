import { defineManifest } from "@crxjs/vite-plugin";

export default defineManifest({
  manifest_version: 3,
  name: "CI Updater Dev Team",
  version: "1.1.0",
  description: "Service Auto Update",
  // ใช้โลโก้ที่เตรียมไว้เป็นไอคอนของส่วนขยาย
  icons: {
    16: "U logo.png",
    32: "U logo.png",
    48: "U logo.png",
    128: "U logo.png"
  },
  permissions: ["tabs", "scripting", "storage", "activeTab", "contextMenus"],
  host_permissions: ["https://ricohap.service-now.com/*"],
  action: {
    default_title: "CI Updater",
    default_popup: "src/popup.html",
    // ตั้งไอคอนปุ่มแอคชัน (ไอคอนบนแถบเครื่องมือ)
    default_icon: {
      16: "U logo.png",
      32: "U logo.png"
    }
  },
  background: {
    service_worker: "src/background.ts",
    type: "module"
  },
  content_scripts: [
    {
      matches: ["https://ricohap.service-now.com/*task_ci_list.do*"],
      js: ["src/common.ts", "src/content-list.ts"],
      run_at: "document_idle",
      all_frames: true
    },
    {
      matches: ["https://ricohap.service-now.com/*task_ci.do*"],
      js: ["src/common.ts", "src/content-add.ts"],
      run_at: "document_idle",
      all_frames: true
    },
    {
      matches: [
        "https://ricohap.service-now.com/*cmdb_ci.do*",
        "https://ricohap.service-now.com/*cmdb_ci_computer.do*",
        // รองรับทุกหมวด CI ที่มี suffix ต่อท้าย cmdb_ci_*
        "https://ricohap.service-now.com/*cmdb_ci_*.do*"
      ],
      js: ["src/common.ts", "src/content-ci-form.ts"],
      run_at: "document_idle",
      all_frames: true
    }
  ],
  commands: {
    "stop-ci-updater": {
      "suggested_key": {
        "default": "Ctrl+Shift+Y",
        "mac": "Command+Shift+Y"
      },
      "description": "Stop CI Updater now"
    }
  }
});
