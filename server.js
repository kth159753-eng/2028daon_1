"use strict";

const { spawn } = require("child_process");
const crypto = require("crypto");
const fs = require("fs");
const os = require("os");
const path = require("path");
const express = require("express");
const compression = require("compression");
const multer = require("multer");
const { createStore } = require("./store");
const { previewKind, previewMime, buildSlides } = require("./preview");

const IS_PROD = process.env.NODE_ENV === "production" || Boolean(process.env.RENDER || process.env.RAILWAY_ENVIRONMENT || process.env.FLY_APP_NAME);
const START_PORT = Number(process.env.PORT) || 3080;
const ROOT = __dirname;
const PUBLIC_DIR = ROOT;
const DATA_DIR = path.resolve(process.env.DATA_DIR || path.join(ROOT, "data"));
const UPLOAD_DIR = path.resolve(process.env.UPLOAD_DIR || (process.env.DATA_DIR ? path.join(DATA_DIR, "uploads") : path.join(ROOT, "uploads")));
const META_PATH = path.join(DATA_DIR, "meta.json");
const TEAM_PASSWORD = String(process.env.TEAM_PASSWORD || process.env.ACCESS_PASSWORD || "");
const AUTH_SECRET = process.env.SESSION_SECRET || TEAM_PASSWORD || "daon-local";

const store = createStore({
  dataDir: DATA_DIR,
  uploadDir: UPLOAD_DIR,
  metaPath: META_PATH,
  token: process.env.GITHUB_TOKEN || process.env.GH_TOKEN || "",
  repo: process.env.GITHUB_REPO || "",
  branch: process.env.GITHUB_BRANCH || "storage",
});
const MAX_FILE_BYTES = store.maxFileBytes;

const MENUS = [
  { id: "memo", label: "메모장(자유롭게)", icon: "note" },
  { id: "merged", label: "2028 대입 지도서(합본)", icon: "edit" },
  { id: "refs", label: "[집필 참조 소스 및 자료]", icon: "folder" },
  { id: "ch0", label: "2028 대입 총론 (신호철)", icon: "doc" },
  { id: "ch1", label: "학생부 종합 전형 (천연정)", icon: "doc" },
  { id: "ch3", label: "3. 교과전형 (이상원)", icon: "doc" },
  { id: "ch4", label: "4. 기균 및 농어촌 (김성배)", icon: "doc" },
  { id: "ch5a", label: "5. 논술 (김승범)", icon: "doc" },
  { id: "ch5b", label: "5. 논술 (노승연)", icon: "doc" },
  { id: "ch6", label: "6. 면접 전형 (박진)", icon: "doc" },
  { id: "ch7", label: "7. 의치한약수 (이동균)", icon: "doc" },
  { id: "ch8", label: "8. 정시 전형 (김태훈)", icon: "doc" },
  { id: "metro1", label: "수도권 대학 I (류승택)", icon: "doc" },
  { id: "metro2", label: "수도권 대학 II (이상원)", icon: "doc" },
];

const MENU_IDS = new Set(MENUS.map((m) => m.id));

const ALLOWED_EXT = new Set([
  "ppt", "pptx", "pot", "potx", "pps", "ppsx",
  "hwp", "hwpx", "hwt", "hml", "hwpml", "owpml",
  "cell", "show", "nxl", "nxls",
  "txt", "text", "pdf",
]);

function cookiesOf(req) {
  const out = {};
  for (const part of String(req.headers.cookie || "").split(";")) {
    const i = part.indexOf("=");
    if (i < 0) continue;
    const key = part.slice(0, i).trim();
    if (!key) continue;
    try {
      out[key] = decodeURIComponent(part.slice(i + 1).trim());
    } catch (_) {
      out[key] = part.slice(i + 1).trim();
    }
  }
  return out;
}

function authValue() {
  return crypto.createHmac("sha256", AUTH_SECRET).update("ok").digest("hex");
}

function sameText(a, b) {
  const left = Buffer.from(String(a || ""));
  const right = Buffer.from(String(b || ""));
  if (left.length !== right.length || !left.length) return false;
  return crypto.timingSafeEqual(left, right);
}

function isAuthed(req) {
  if (!TEAM_PASSWORD) return true;
  return sameText(cookiesOf(req).daon, authValue());
}

function setAuthCookie(res, req) {
  const secure = req.secure || String(req.headers["x-forwarded-proto"] || "").includes("https");
  res.setHeader(
    "Set-Cookie",
    `daon=${authValue()}; HttpOnly; Path=/; SameSite=Lax; Max-Age=2592000${secure ? "; Secure" : ""}`
  );
}

let meta = { files: [], labels: {}, sheet: null };

function hydrateMeta(parsed) {
  const next = parsed && typeof parsed === "object" ? parsed : { files: [], labels: {}, sheet: null };
  if (!Array.isArray(next.files)) next.files = [];
  if (!next.labels || typeof next.labels !== "object" || Array.isArray(next.labels)) next.labels = {};
  next.sheet = normalizeSheet(next.sheet);
  return next;
}

async function saveMeta(next) {
  await store.saveMeta(next);
}

function extOf(name) {
  const i = String(name).lastIndexOf(".");
  return i >= 0 ? String(name).slice(i + 1).toLowerCase() : "";
}

function decodeOrig(name) {
  try {
    return Buffer.from(name, "latin1").toString("utf8");
  } catch (_) {
    return name;
  }
}

function safeOriginalName(name) {
  return path.basename(String(name || "untitled")).replace(/[\u0000-\u001f]/g, "").slice(0, 240) || "untitled";
}

const storage = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, store.kind === "github" ? os.tmpdir() : UPLOAD_DIR),
  filename: (_req, file, cb) => {
    const ext = extOf(decodeOrig(file.originalname));
    cb(null, `${Date.now().toString(36)}-${crypto.randomBytes(8).toString("hex")}${ext ? "." + ext : ""}`);
  },
});

const upload = multer({
  storage,
  limits: { fileSize: MAX_FILE_BYTES, files: 40 },
  fileFilter: (_req, file, cb) => {
    const ext = extOf(decodeOrig(file.originalname));
    if (!ALLOWED_EXT.has(ext)) {
      cb(Object.assign(new Error("UNSUPPORTED_TYPE"), { status: 400 }));
      return;
    }
    cb(null, true);
  },
});

const app = express();
app.disable("x-powered-by");
if (IS_PROD) app.set("trust proxy", 1);
app.use(compression({ threshold: 512 }));
app.use(express.json({ limit: "2mb" }));
app.use((req, res, next) => {
  if (!req.path.startsWith("/api/")) return next();
  if (req.path === "/api/health" || req.path === "/api/auth" || req.path === "/api/login") return next();
  if (isAuthed(req)) return next();
  res.status(401).json({ error: "비밀번호가 필요합니다.", auth: true });
});
app.use((req, res, next) => {
  const p = String(req.path || "").toLowerCase();
  if (
    p.startsWith("/node_modules") ||
    p.startsWith("/data") ||
    p.startsWith("/uploads") ||
    p === "/server.js" ||
    p === "/store.js" ||
    p === "/preview.js" ||
    p.startsWith("/package") ||
    p.endsWith(".bat") ||
    p.endsWith(".json") ||
    p.endsWith(".yml") ||
    p.endsWith(".yaml") ||
    p === "/dockerfile" ||
    p === "/procfile" ||
    p === "/.dockerignore"
  ) {
    res.status(404).end();
    return;
  }
  next();
});
app.use(
  express.static(PUBLIC_DIR, {
    index: "index.html",
    etag: true,
    lastModified: true,
    maxAge: "2h",
    setHeaders(res, filePath) {
      if (/\.(html|js|css)$/.test(filePath)) res.setHeader("Cache-Control", "no-cache");
    },
  })
);

function counts() {
  const map = Object.fromEntries(MENUS.map((m) => [m.id, 0]));
  for (const f of meta.files) {
    if (map[f.menuId] != null) map[f.menuId] += 1;
  }
  return map;
}

function publicFile(f) {
  return { id: f.id, name: f.name, ext: f.ext, size: f.size, uploadedAt: f.uploadedAt, menuId: f.menuId };
}

function normalizeSheet(raw) {
  const src = raw && typeof raw === "object" ? raw : {};
  const rows = Math.min(200, Math.max(10, Number(src.rows) || 40));
  const cols = Math.min(40, Math.max(4, Number(src.cols) || 10));
  const cells = {};
  if (src.cells && typeof src.cells === "object" && !Array.isArray(src.cells)) {
    for (const [key, value] of Object.entries(src.cells)) {
      if (!/^\d+,\d+$/.test(key)) continue;
      const text = String(value == null ? "" : value);
      if (!text) continue;
      cells[key] = text.slice(0, 20000);
    }
  }
  return { rows, cols, cells };
}

function publicMenus() {
  const labels = meta.labels || {};
  return MENUS.map((m) => {
    const custom = String(labels[m.id] || "").replace(/\s+/g, " ").trim();
    return { ...m, label: custom || m.label };
  });
}

function findFile(id) {
  return meta.files.find((f) => f.id === id);
}

app.get("/api/health", (_req, res) => {
  res.setHeader("Cache-Control", "no-store");
  res.json({ ok: true, store: store.kind });
});

app.get("/api/auth", (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  res.json({ required: Boolean(TEAM_PASSWORD), ok: isAuthed(req) });
});

app.post("/api/login", (req, res) => {
  if (!TEAM_PASSWORD) {
    res.json({ ok: true });
    return;
  }
  if (!sameText((req.body && req.body.password) || "", TEAM_PASSWORD)) {
    res.status(401).json({ error: "비밀번호가 올바르지 않습니다." });
    return;
  }
  setAuthCookie(res, req);
  res.json({ ok: true });
});

app.get("/api/bootstrap", (_req, res) => {
  res.setHeader("Cache-Control", "no-store");
  const persist = store.kind === "github" || Boolean(process.env.DATA_DIR) || !IS_PROD;
  res.json({ menus: publicMenus(), counts: counts(), persist, store: store.kind });
});

app.patch("/api/menus/:id", async (req, res) => {
  const id = String(req.params.id || "");
  if (!MENU_IDS.has(id)) {
    res.status(400).json({ error: "메뉴를 찾을 수 없습니다." });
    return;
  }
  const label = String(req.body && req.body.label || "").replace(/\s+/g, " ").trim();
  if (!label || label.length > 80) {
    res.status(400).json({ error: "제목은 1~80자로 입력해 주세요." });
    return;
  }
  if (!meta.labels) meta.labels = {};
  meta.labels[id] = label;
  try {
    await saveMeta(meta);
  } catch (_) {
    res.status(500).json({ error: "저장에 실패했습니다." });
    return;
  }
  res.json({ ok: true, menus: publicMenus() });
});

app.get("/api/sheet", (_req, res) => {
  res.setHeader("Cache-Control", "no-store");
  res.json(normalizeSheet(meta.sheet));
});

app.put("/api/sheet", async (req, res) => {
  meta.sheet = normalizeSheet(req.body);
  try {
    await saveMeta(meta);
  } catch (_) {
    res.status(500).json({ error: "저장에 실패했습니다." });
    return;
  }
  res.json({ ok: true, sheet: meta.sheet });
});

app.get("/api/files", (req, res) => {
  const menuId = String(req.query.menu || "");
  if (!MENU_IDS.has(menuId)) {
    res.status(400).json({ error: "메뉴를 찾을 수 없습니다." });
    return;
  }
  const files = [];
  for (let i = meta.files.length - 1; i >= 0; i--) {
    const f = meta.files[i];
    if (f.menuId === menuId) files.push(publicFile(f));
  }
  res.setHeader("Cache-Control", "no-store");
  res.json({ files });
});

app.post("/api/upload", (req, res) => {
  upload.array("files", 40)(req, res, async (err) => {
    if (err) {
      const maxMb = Math.round(MAX_FILE_BYTES / 1024 / 1024);
      const status = err.status || (err.code === "LIMIT_FILE_SIZE" ? 413 : 400);
      const message =
        err.message === "UNSUPPORTED_TYPE"
          ? "ppt, pptx, pdf, hwp, hwpx, 메모장, 한글오피스 파일만 올릴 수 있습니다."
          : err.code === "LIMIT_FILE_SIZE"
            ? `파일이 너무 큽니다. (최대 ${maxMb}MB)`
            : "업로드에 실패했습니다.";
      res.status(status).json({ error: message });
      return;
    }

    const menuId = String(req.body.menuId || "");
    if (!MENU_IDS.has(menuId)) {
      for (const file of req.files || []) fs.unlink(file.path, () => {});
      res.status(400).json({ error: "메뉴를 선택해 주세요." });
      return;
    }

    const added = [];
    const now = new Date().toISOString();
    try {
      for (const file of req.files || []) {
        await store.putFile(path.basename(file.filename), file.path);
        const name = safeOriginalName(decodeOrig(file.originalname));
        const row = {
          id: crypto.randomBytes(12).toString("hex"),
          name,
          stored: path.basename(file.filename),
          ext: extOf(name),
          size: file.size,
          menuId,
          uploadedAt: now,
        };
        meta.files.push(row);
        added.push(publicFile(row));
      }
      await saveMeta(meta);
    } catch (_) {
      res.status(500).json({ error: "저장에 실패했습니다." });
      return;
    }

    res.json({ ok: true, files: added, counts: counts() });
  });
});

app.get("/api/preview/:id", async (req, res) => {
  const row = findFile(String(req.params.id || ""));
  if (!row) {
    res.status(404).json({ error: "파일을 찾을 수 없습니다." });
    return;
  }
  const kind = previewKind(row.ext);
  if (kind !== "pdf" && kind !== "text") {
    res.status(400).json({ error: "이 파일은 바로 미리볼 수 없습니다." });
    return;
  }
  try {
    const buf = await store.readFile(row.stored);
    if (!buf) {
      res.status(404).json({ error: "파일이 디스크에 없습니다." });
      return;
    }
    res.setHeader("Cache-Control", "private, max-age=120");
    res.setHeader("Content-Type", previewMime(row.ext));
    res.setHeader("Content-Disposition", "inline");
    res.setHeader("X-Content-Type-Options", "nosniff");
    res.end(buf);
  } catch (_) {
    res.status(404).json({ error: "파일이 디스크에 없습니다." });
  }
});

app.get("/api/slides/:id", async (req, res) => {
  const row = findFile(String(req.params.id || ""));
  if (!row) {
    res.status(404).json({ error: "파일을 찾을 수 없습니다." });
    return;
  }
  const kind = previewKind(row.ext);
  if (kind !== "slides" && kind !== "hwpx") {
    res.status(400).json({ error: "슬라이드 미리보기를 지원하지 않는 형식입니다." });
    return;
  }
  try {
    const buf = await store.readFile(row.stored);
    if (!buf) {
      res.status(404).json({ error: "파일이 디스크에 없습니다." });
      return;
    }
    const slides = await buildSlides(row.ext, buf);
    res.setHeader("Cache-Control", "private, max-age=120");
    res.json({ ok: true, name: row.name, slides });
  } catch (_) {
    res.status(500).json({ error: "미리보기를 만들지 못했습니다." });
  }
});

app.get("/api/download/:id", async (req, res) => {
  const row = findFile(String(req.params.id || ""));
  if (!row) {
    res.status(404).json({ error: "파일을 찾을 수 없습니다." });
    return;
  }
  try {
    const ok = await store.sendFile(res, row.stored, row.name);
    if (!ok) res.status(404).json({ error: "파일이 디스크에 없습니다." });
  } catch (_) {
    res.status(404).json({ error: "파일이 디스크에 없습니다." });
  }
});

app.delete("/api/files/:id", async (req, res) => {
  const id = String(req.params.id || "");
  const idx = meta.files.findIndex((f) => f.id === id);
  if (idx < 0) {
    res.status(404).json({ error: "파일을 찾을 수 없습니다." });
    return;
  }
  const [row] = meta.files.splice(idx, 1);
  try {
    await store.deleteFile(row.stored);
    await saveMeta(meta);
  } catch (_) {
    meta.files.splice(idx, 0, row);
    res.status(500).json({ error: "삭제에 실패했습니다." });
    return;
  }
  res.json({ ok: true, counts: counts() });
});

app.use((err, _req, res, _next) => {
  res.status(err.status || 500).json({ error: "서버 오류가 발생했습니다." });
});

function lanUrls(port) {
  const urls = [];
  for (const list of Object.values(os.networkInterfaces())) {
    for (const item of list || []) {
      if (item.family === "IPv4" && !item.internal) urls.push(`http://${item.address}:${port}`);
    }
  }
  return urls;
}

function shouldOpenBrowser() {
  if (process.env.NO_OPEN === "1" || IS_PROD) return false;
  return true;
}

function openBrowser(url) {
  if (!shouldOpenBrowser()) return;
  try {
    if (process.platform === "win32") {
      spawn("cmd", ["/c", "start", "", url], { detached: true, stdio: "ignore" }).unref();
    } else if (process.platform === "darwin") {
      spawn("open", [url], { detached: true, stdio: "ignore" }).unref();
    } else {
      spawn("xdg-open", [url], { detached: true, stdio: "ignore" }).unref();
    }
  } catch (_) {}
}

async function isOurServer(port) {
  try {
    const res = await fetch(`http://127.0.0.1:${port}/api/health`, { signal: AbortSignal.timeout(600) });
    return res.ok;
  } catch (_) {
    return false;
  }
}

function listen(port) {
  const server = app.listen(port, "0.0.0.0");
  server.on("listening", () => {
    const local = `http://localhost:${port}`;
    console.log("");
    console.log("  2028 대입 지도서 (다온) 첨부파일");
    console.log(`  이 컴퓨터: ${local}`);
    if (!IS_PROD) {
      for (const url of lanUrls(port)) console.log(`  다른 컴퓨터: ${url}`);
      console.log("  이 창을 닫으면 사이트가 꺼집니다.");
    } else {
      console.log(`  저장: ${store.kind === "github" ? "GitHub" : "디스크"}`);
    }
    console.log("");
    openBrowser(local);
  });
  server.on("error", async (err) => {
    if (err.code !== "EADDRINUSE") {
      console.error(err);
      process.exit(1);
    }
    if (await isOurServer(port)) {
      const local = `http://localhost:${port}`;
      console.log("이미 실행 중입니다. 브라우저를 엽니다.");
      openBrowser(local);
      process.exit(0);
    }
    if (port < START_PORT + 10) listen(port + 1);
    else {
      console.error("빈 포트를 찾지 못했습니다.");
      process.exit(1);
    }
  });
}

async function start() {
  try {
    meta = hydrateMeta(await store.loadMeta());
  } catch (err) {
    console.error(err && err.message ? err.message : err);
    process.exit(1);
  }
  if (IS_PROD && store.kind === "local" && !process.env.DATA_DIR) {
    console.log("  경고: 재시작하면 올린 파일이 사라집니다. GITHUB_TOKEN 과 GITHUB_REPO 를 설정하세요.");
  }
  listen(START_PORT);
}

start();
