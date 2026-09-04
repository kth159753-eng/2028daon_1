"use strict";

const { spawn } = require("child_process");
const crypto = require("crypto");
const fs = require("fs");
const os = require("os");
const path = require("path");

const OFFICE_EXTS = new Set(["ppt", "pptx", "pot", "potx", "pps", "ppsx", "hwp", "hwpx", "hwt"]);
const PPT_EXTS = new Set(["ppt", "pptx", "pot", "potx", "pps", "ppsx"]);
const HWP_EXTS = new Set(["hwp", "hwpx", "hwt"]);

const inFlight = new Map();

function exists(p) {
  try {
    return p && fs.existsSync(p);
  } catch (_) {
    return false;
  }
}

function findSoffice() {
  const list = [
    process.env.SOFFICE_PATH,
    "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
    "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
  ];
  return list.find((p) => exists(p)) || "";
}

function run(cmd, args, timeoutMs) {
  return new Promise((resolve, reject) => {
    const child = spawn(cmd, args, { windowsHide: true, stdio: ["ignore", "pipe", "pipe"] });
    let err = "";
    child.stderr.on("data", (d) => {
      err += d;
    });
    const timer = setTimeout(() => {
      try {
        child.kill();
      } catch (_) {}
      reject(new Error("변환 시간이 너무 깁니다."));
    }, timeoutMs);
    child.on("error", (e) => {
      clearTimeout(timer);
      reject(e);
    });
    child.on("close", (code) => {
      clearTimeout(timer);
      if (code === 0) resolve();
      else reject(new Error((err || "변환에 실패했습니다.").trim().slice(0, 300)));
    });
  });
}

function runPs(scriptPath, inputPath, outputPath) {
  return run(
    "powershell.exe",
    [
      "-NoProfile",
      "-STA",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      scriptPath,
      "-InputPath",
      inputPath,
      "-OutputPath",
      outputPath,
    ],
    120000
  );
}

async function convertLibreOffice(srcPath, outDir) {
  const soffice = findSoffice();
  if (!soffice) throw new Error("NO_LO");
  const profile = path.join(os.tmpdir(), "daon-lo-" + crypto.randomBytes(6).toString("hex"));
  const profileUrl = "file:///" + profile.replace(/\\/g, "/");
  try {
    await run(
      soffice,
      [
        "--headless",
        "--norestore",
        "--nolockcheck",
        "--nodefault",
        "--nofirststartwizard",
        `-env:UserInstallation=${profileUrl}`,
        "--convert-to",
        "pdf",
        "--outdir",
        outDir,
        srcPath,
      ],
      120000
    );
  } finally {
    fs.rm(profile, { recursive: true, force: true }, () => {});
  }
  const pdf = path.join(outDir, path.basename(srcPath, path.extname(srcPath)) + ".pdf");
  if (!exists(pdf)) throw new Error("PDF 파일을 만들지 못했습니다.");
  return pdf;
}

async function convertOffice(buf, ext, destPdf) {
  const work = fs.mkdtempSync(path.join(os.tmpdir(), "daon-prev-"));
  const src = path.join(work, "source." + ext);
  const dst = path.join(work, "out.pdf");
  fs.writeFileSync(src, buf);
  const pptScript = path.join(__dirname, "scripts", "ppt-to-pdf.ps1");
  const hwpScript = path.join(__dirname, "scripts", "hwp-to-pdf.ps1");
  const errors = [];

  const tryPs = async (script) => {
    await runPs(script, src, dst);
    if (!exists(dst)) throw new Error("PDF 파일을 만들지 못했습니다.");
    return dst;
  };

  try {
    if (PPT_EXTS.has(ext)) {
      if (exists(pptScript) && process.platform === "win32") {
        try {
          await tryPs(pptScript);
        } catch (e) {
          errors.push(e.message);
          const made = await convertLibreOffice(src, work);
          fs.copyFileSync(made, dst);
        }
      } else {
        const made = await convertLibreOffice(src, work);
        fs.copyFileSync(made, dst);
      }
    } else if (HWP_EXTS.has(ext)) {
      if (exists(hwpScript) && process.platform === "win32") {
        try {
          await tryPs(hwpScript);
        } catch (e) {
          errors.push(e.message);
          const made = await convertLibreOffice(src, work);
          fs.copyFileSync(made, dst);
        }
      } else {
        const made = await convertLibreOffice(src, work);
        fs.copyFileSync(made, dst);
      }
    } else {
      throw new Error("지원하지 않는 형식입니다.");
    }
    if (!exists(dst)) throw new Error(errors[0] || "PDF 파일을 만들지 못했습니다.");
    fs.mkdirSync(path.dirname(destPdf), { recursive: true });
    fs.copyFileSync(dst, destPdf);
  } finally {
    fs.rm(work, { recursive: true, force: true }, () => {});
  }
}

function cachePath(previewDir, row) {
  return path.join(previewDir, `${row.id}-${row.size}.pdf`);
}

async function ensurePdf(previewDir, row, readFile) {
  const dest = cachePath(previewDir, row);
  if (exists(dest) && fs.statSync(dest).size > 80) return dest;
  const key = row.id;
  if (inFlight.has(key)) return inFlight.get(key);
  const job = (async () => {
    const buf = await readFile(row.stored);
    if (!buf) throw new Error("파일이 디스크에 없습니다.");
    await convertOffice(buf, String(row.ext || "").toLowerCase(), dest);
    return dest;
  })().finally(() => inFlight.delete(key));
  inFlight.set(key, job);
  return job;
}

function removePdf(previewDir, row) {
  try {
    const dest = cachePath(previewDir, row);
    if (exists(dest)) fs.unlinkSync(dest);
  } catch (_) {}
}

function isOfficeExt(ext) {
  return OFFICE_EXTS.has(String(ext || "").toLowerCase());
}

module.exports = { ensurePdf, removePdf, isOfficeExt, OFFICE_EXTS };
