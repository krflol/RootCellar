import { spawn } from "node:child_process";
import { access, mkdir } from "node:fs/promises";
import { constants } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { setTimeout as sleep } from "node:timers/promises";
import { chromium } from "playwright";

const desktopRoot = path.resolve(fileURLToPath(new URL("..", import.meta.url)));
const distIndexPath = path.join(desktopRoot, "dist", "index.html");
const capturesDir = path.join(desktopRoot, "artifacts", "ui-captures");
const host = "127.0.0.1";
const port = 4173;
const baseUrl = `http://${host}:${port}`;

function npmInvocation(args) {
  if (process.platform === "win32") {
    return {
      command: "cmd.exe",
      args: ["/d", "/s", "/c", "npm", ...args],
    };
  }
  return {
    command: "npm",
    args,
  };
}

function runNpm(args, cwd) {
  const invocation = npmInvocation(args);
  return new Promise((resolve, reject) => {
    const child = spawn(invocation.command, invocation.args, {
      cwd,
      stdio: "inherit",
      shell: false,
    });

    child.on("error", reject);
    child.on("exit", (code) => {
      if (code === 0) {
        resolve();
        return;
      }
      reject(new Error(`npm ${args.join(" ")} failed with exit code ${code}`));
    });
  });
}

function spawnNpm(args, cwd) {
  const invocation = npmInvocation(args);
  return spawn(invocation.command, invocation.args, {
    cwd,
    stdio: "ignore",
    shell: false,
  });
}

async function stopProcessTree(processHandle) {
  if (!processHandle || processHandle.exitCode != null) {
    return;
  }

  if (process.platform === "win32") {
    await new Promise((resolve) => {
      const killer = spawn(
        "taskkill",
        ["/PID", String(processHandle.pid), "/T", "/F"],
        { stdio: "ignore", shell: false },
      );
      killer.on("exit", () => resolve());
      killer.on("error", () => resolve());
    });
    return;
  }

  processHandle.kill("SIGTERM");
  await sleep(300);
  if (processHandle.exitCode == null) {
    processHandle.kill("SIGKILL");
  }
}

async function ensureBuilt() {
  try {
    await access(distIndexPath, constants.F_OK);
  } catch {
    // Fall through and build.
  }
  await runNpm(["run", "build"], desktopRoot);
}

async function waitForServer(url, timeoutMs = 30_000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) {
        return;
      }
    } catch {
      // Retry until timeout.
    }
    await sleep(250);
  }
  throw new Error(`preview server did not become ready within ${timeoutMs}ms`);
}

async function capturePage(page, url, outputPath, viewport, selector = null) {
  await page.setViewportSize(viewport);
  await page.goto(url, { waitUntil: "networkidle" });
  await page.waitForTimeout(400);
  if (selector) {
    const section = page.locator(selector);
    await section.waitFor({ state: "visible", timeout: 10_000 });
    await section.scrollIntoViewIfNeeded();
    await page.waitForTimeout(200);
    await section.screenshot({ path: outputPath });
    return;
  }
  await page.screenshot({ path: outputPath, fullPage: true });
}

async function main() {
  await ensureBuilt();
  await mkdir(capturesDir, { recursive: true });

  const preview = spawnNpm(
    ["run", "preview", "--", "--host", host, "--port", String(port), "--strictPort"],
    desktopRoot,
  );

  let browser;
  try {
    await waitForServer(`${baseUrl}/`);

    browser = await chromium.launch({ headless: true });
    const page = await browser.newPage();

    await capturePage(
      page,
      `${baseUrl}/?ui_capture=1&capture_state=fresh`,
      path.join(capturesDir, "desktop-fresh.png"),
      { width: 1440, height: 1800 },
    );
    await capturePage(
      page,
      `${baseUrl}/?ui_capture=1&capture_state=stale`,
      path.join(capturesDir, "desktop-stale.png"),
      { width: 1440, height: 1800 },
    );
    await capturePage(
      page,
      `${baseUrl}/?ui_capture=1&capture_state=stale&capture_section=edit-cell`,
      path.join(capturesDir, "desktop-edit-cell.png"),
      { width: 1440, height: 1200 },
      "#capture-section-edit-cell",
    );
    await capturePage(
      page,
      `${baseUrl}/?ui_capture=1&capture_state=stale&capture_section=save-recalc`,
      path.join(capturesDir, "desktop-save-recalc.png"),
      { width: 1440, height: 1200 },
      "#capture-section-save-recalc",
    );
    await capturePage(
      page,
      `${baseUrl}/?ui_capture=1&capture_state=pending`,
      path.join(capturesDir, "desktop-pending.png"),
      { width: 1440, height: 1800 },
    );
    await capturePage(
      page,
      `${baseUrl}/?ui_capture=1&capture_state=stale`,
      path.join(capturesDir, "mobile-stale.png"),
      { width: 390, height: 1800 },
    );

    console.log(`UI captures written to: ${capturesDir}`);
  } finally {
    if (browser) {
      await browser.close();
    }
    await stopProcessTree(preview);
  }
}

main().catch((error) => {
  console.error(`ui capture error: ${String(error)}`);
  process.exitCode = 1;
});
