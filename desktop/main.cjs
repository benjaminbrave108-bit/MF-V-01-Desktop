const { app, BrowserWindow, dialog, screen } = require("electron");
const http = require("http");
const fs = require("fs");
const path = require("path");
const { pathToFileURL } = require("url");

let mainWindow = null;
let server = null;

const MIME_TYPES = {
  ".css": "text/css; charset=utf-8",
  ".gif": "image/gif",
  ".html": "text/html; charset=utf-8",
  ".ico": "image/x-icon",
  ".jpeg": "image/jpeg",
  ".jpg": "image/jpeg",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml; charset=utf-8",
  ".webp": "image/webp",
  ".woff": "font/woff",
  ".woff2": "font/woff2",
};

function log(message) {
  try {
    fs.appendFileSync(path.join(app.getPath("userData"), "startup.log"), `[${new Date().toISOString()}] ${message}\n`);
  } catch {}
}

function configureWritablePaths() {
  const base = path.join(app.getPath("appData"), "MF-V-01");
  app.setPath("userData", path.join(base, "UserData"));
  app.setPath("cache", path.join(base, "Cache"));
  fs.mkdirSync(app.getPath("userData"), { recursive: true });
  fs.mkdirSync(app.getPath("cache"), { recursive: true });
}

function resolveClientAsset(urlValue) {
  let pathname;
  try {
    pathname = decodeURIComponent(new URL(urlValue, "http://127.0.0.1").pathname);
  } catch {
    return null;
  }
  if (pathname === "/") return null;
  const clientRoot = path.resolve(app.getAppPath(), "dist", "client");
  const candidate = path.resolve(clientRoot, `.${pathname}`);
  if (candidate !== clientRoot && !candidate.startsWith(`${clientRoot}${path.sep}`)) return null;
  try {
    return fs.statSync(candidate).isFile() ? candidate : null;
  } catch {
    return null;
  }
}

function fileResponse(filePath) {
  const body = fs.readFileSync(filePath);
  return new Response(body, {
    status: 200,
    headers: {
      "content-type": MIME_TYPES[path.extname(filePath).toLowerCase()] || "application/octet-stream",
      "cache-control": "no-cache",
    },
  });
}

function sendResponse(res, response) {
  res.statusCode = response.status;
  response.headers.forEach((value, key) => res.setHeader(key, value));
  return response.arrayBuffer().then((body) => res.end(Buffer.from(body)));
}

function startApplicationServer() {
  const modulePath = path.join(app.getAppPath(), "dist", "server", "index.js");
  return import(pathToFileURL(modulePath).href).then(({ default: worker }) => new Promise((resolve, reject) => {
    server = http.createServer(async (req, res) => {
      try {
        const origin = `http://127.0.0.1:${server.address().port}`;
        const staticFile = resolveClientAsset(req.url || "/");
        if (staticFile) {
          await sendResponse(res, fileResponse(staticFile));
          return;
        }
        const chunks = [];
        for await (const chunk of req) chunks.push(chunk);
        const method = req.method || "GET";
        const request = new Request(new URL(req.url || "/", origin), {
          method,
          headers: req.headers,
          body: method === "GET" || method === "HEAD" ? undefined : Buffer.concat(chunks),
          duplex: method === "GET" || method === "HEAD" ? undefined : "half",
        });
        // Vinext expects a Cloudflare-compatible ExecutionContext. Passing an
        // empty object makes rendered-page cache finalization call a missing
        // waitUntil() method and returns a blank 500 page on Windows.
        const backgroundTasks = [];
        const executionContext = {
          waitUntil(promise) {
            backgroundTasks.push(Promise.resolve(promise).catch((error) => {
              log(`Background task error: ${error?.stack || error}`);
            }));
          },
          passThroughOnException() {},
        };
        const assets = {
          fetch(assetRequest) {
            const file = resolveClientAsset(assetRequest.url);
            return Promise.resolve(file ? fileResponse(file) : new Response("Not found", { status: 404 }));
          },
        };
        const response = await worker.fetch(request, { ASSETS: assets }, executionContext);
        await sendResponse(res, response);
      } catch (error) {
        log(`Request error: ${error?.stack || error}`);
        res.statusCode = 500;
        res.end("MF-V-01 could not load.");
      }
    });
    server.once("error", reject);
    // Keep a stable origin so Chromium localStorage remains available after
    // the desktop app is closed and reopened. A random port creates a new
    // origin on every launch and makes previously saved finance data appear lost.
    const persistentPort = 47831;
    server.listen(persistentPort, "127.0.0.1", () => resolve(`http://127.0.0.1:${persistentPort}`));
  }));
}

function showMainWindow(url) {
  const area = screen.getPrimaryDisplay().workArea;
  const width = Math.min(area.width, Math.min(1440, Math.max(720, area.width - 40)));
  const height = Math.min(area.height, Math.min(960, Math.max(540, area.height - 40)));
  mainWindow = new BrowserWindow({
    width, height,
    x: area.x + Math.round((area.width - width) / 2),
    y: area.y + Math.round((area.height - height) / 2),
    minWidth: 680,
    minHeight: 500,
    show: false,
    autoHideMenuBar: true,
    backgroundColor: "#f3f6f8",
    icon: path.join(app.getAppPath(), "public", "kb-logo.png"),
    webPreferences: { contextIsolation: true, sandbox: true },
  });
  mainWindow.once("ready-to-show", () => { mainWindow.show(); mainWindow.focus(); });
  mainWindow.webContents.on("console-message", (_event, level, message, line, sourceId) => {
    log(`Renderer console [${level}] ${message} (${sourceId || "unknown"}:${line || 0})`);
  });
  mainWindow.webContents.on("did-fail-load", (_event, errorCode, errorDescription, validatedURL, isMainFrame) => {
    log(`Renderer load failed: ${errorCode} ${errorDescription}; mainFrame=${isMainFrame}; url=${validatedURL}`);
  });
  mainWindow.webContents.on("render-process-gone", (_event, details) => {
    log(`Renderer process gone: ${JSON.stringify(details)}`);
  });
  mainWindow.webContents.on("did-finish-load", () => {
    setTimeout(async () => {
      if (!mainWindow || mainWindow.isDestroyed()) return;
      try {
        const rendered = await mainWindow.webContents.executeJavaScript(
          "document.body && (document.body.innerText.trim().length > 0 || document.body.querySelectorAll('button,input,select,textarea').length > 0)",
          true,
        );
        log(`Renderer health check: ${rendered ? "OK" : "EMPTY"}`);
        if (!rendered) {
          const logPath = path.join(app.getPath("userData"), "startup.log");
          dialog.showErrorBox(
            "MF-V-01 - Arayüz yüklenemedi",
            `Uygulama penceresi açıldı ancak arayüz yüklenemedi.\n\nHata kaydı: ${logPath}`,
          );
        }
      } catch (error) {
        log(`Renderer health check error: ${error?.stack || error}`);
      }
    }, 5000);
  });
  mainWindow.loadURL(url).catch((error) => log(`loadURL error: ${error?.stack || error}`));
  setTimeout(() => {
    if (mainWindow && !mainWindow.isDestroyed() && !mainWindow.isVisible()) {
      mainWindow.center(); mainWindow.show(); mainWindow.focus();
    }
  }, 4000);
}

configureWritablePaths();
const hasLock = app.requestSingleInstanceLock();
if (!hasLock) app.quit();
else {
  app.on("second-instance", () => {
    if (mainWindow) { if (mainWindow.isMinimized()) mainWindow.restore(); mainWindow.center(); mainWindow.show(); mainWindow.focus(); }
  });
  app.whenReady().then(async () => {
    try {
      log(`Starting MF-V-01 ${app.getVersion()}`);
      showMainWindow(await startApplicationServer());
    } catch (error) {
      log(`Startup error: ${error?.stack || error}`);
      dialog.showErrorBox("MF-V-01", `The application could not start.\n\n${error?.message || error}\n\nLog: ${path.join(app.getPath("userData"), "startup.log")}`);
      app.quit();
    }
  });
  app.on("activate", () => { if (mainWindow) { mainWindow.center(); mainWindow.show(); mainWindow.focus(); } });
  app.on("window-all-closed", () => app.quit());
  app.on("before-quit", () => { if (server) server.close(); });
}
