import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("@tauri-apps/api/core", () => ({
  invoke: vi.fn(),
}));

vi.mock("@tauri-apps/plugin-dialog", () => ({
  open: vi.fn(),
  save: vi.fn(),
}));

describe("edit lifecycle observability", () => {
  let main: typeof import("./main");

  beforeEach(async () => {
    vi.resetModules();
    document.body.innerHTML = `<div id="app"></div>`;
    main = await import("./main");
    main.clearEditLifecycleEntriesForTests();
  });

  afterEach(() => {
    document.body.innerHTML = "";
  });

  it("keeps an in-memory lifecycle window and surfaces the newest events first", () => {
    for (let index = 0; index < 28; index += 1) {
      main.logEditLifecycleEvent("interop_sheet_preview", "success", `preview-${index}`);
    }

    const entries = main.getEditLifecycleEntriesForTests();

    expect(entries).toHaveLength(24);
    expect(entries[0]?.message).toBe("preview-27");
    expect(entries[23]?.message).toBe("preview-4");
  });

  it("emits a start/success/error lifecycle trail and asserts errors assertively", async () => {
    const assertiveRegion = document.querySelector<HTMLDivElement>("#sr-announcer-assertive");
    const output = document.querySelector<HTMLPreElement>("#edit-lifecycle-output");

    main.logEditLifecycleEvent("interop_save_workbook", "start", "saving workbook");
    main.logEditLifecycleEvent("interop_save_workbook", "success", "saved workbook");
    main.logEditLifecycleEvent("interop_save_workbook", "error", "save failed");

    const text = output?.textContent ?? "";

    expect(text).toContain('"command": "interop_save_workbook"');
    expect(text).toContain('"phase": "start"');
    expect(text).toContain('"phase": "success"');
    expect(text).toContain('"phase": "error"');

    await vi.waitFor(() => {
      expect(assertiveRegion?.textContent).toBe("interop_save_workbook failed: save failed");
    });
  });
});
