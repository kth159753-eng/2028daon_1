"use strict";

const JSZip = require("jszip");

const SLIDE_EXTS = new Set(["pptx", "ppsx", "potx"]);
const TEXT_EXTS = new Set(["txt", "text"]);

function decodeXml(s) {
  return String(s || "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function joinZip(fromFile, target) {
  const dir = fromFile.replace(/\\/g, "/").split("/").slice(0, -1);
  for (const part of String(target || "").replace(/\\/g, "/").split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") dir.pop();
    else dir.push(part);
  }
  return dir.join("/");
}

function slideNum(name) {
  const m = String(name).match(/slide(\d+)\.xml$/i);
  return m ? Number(m[1]) : 0;
}

function paragraphsFromSlide(xml) {
  const paras = [];
  const pRe = /<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g;
  let m;
  while ((m = pRe.exec(xml))) {
    let line = "";
    const tRe = /<a:t[^>]*>([^<]*)<\/a:t>|<a:br\s*\/>/g;
    let t;
    while ((t = tRe.exec(m[1]))) {
      if (t[0].startsWith("<a:br")) line += "\n";
      else line += decodeXml(t[1]);
    }
    const text = line.replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
    if (text) paras.push(text);
  }
  return paras;
}

function mimeFromNameOrBytes(name, buf) {
  const ext = String(name).split(".").pop().toLowerCase();
  if (ext === "png") return "image/png";
  if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
  if (ext === "gif") return "image/gif";
  if (ext === "webp") return "image/webp";
  if (buf && buf.length >= 4) {
    if (buf[0] === 0x89 && buf[1] === 0x50) return "image/png";
    if (buf[0] === 0xff && buf[1] === 0xd8) return "image/jpeg";
    if (buf[0] === 0x47 && buf[1] === 0x49) return "image/gif";
  }
  return "";
}

async function imagesFromRels(zip, slidePath, relsXml) {
  const images = [];
  const re = /Target="([^"]+)"/g;
  let m;
  while ((m = re.exec(relsXml))) {
    const target = m[1].split("?")[0];
    if (!/\.(png|jpe?g|gif|webp)$/i.test(target) && !/media\//i.test(target)) continue;
    const imgPath = joinZip(slidePath, target);
    const file = zip.file(imgPath);
    if (!file) continue;
    const buf = await file.async("nodebuffer");
    if (!buf || buf.length < 24 || buf.length > 4 * 1024 * 1024) continue;
    const mime = mimeFromNameOrBytes(imgPath, buf);
    if (!mime) continue;
    images.push(`data:${mime};base64,${buf.toString("base64")}`);
    if (images.length >= 6) break;
  }
  return images;
}

async function parsePptx(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const names = Object.keys(zip.files)
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/i.test(n))
    .sort((a, b) => slideNum(a) - slideNum(b))
    .slice(0, 80);
  const slides = [];
  for (const name of names) {
    const xml = await zip.file(name).async("string");
    const texts = paragraphsFromSlide(xml);
    let images = [];
    const relsName = name.replace(/slides\/([^/]+)$/i, "slides/_rels/$1.rels");
    const relsFile = zip.file(relsName);
    if (relsFile) {
      try {
        images = await imagesFromRels(zip, name, await relsFile.async("string"));
      } catch (_) {}
    }
    slides.push({ texts, images });
  }
  return slides;
}

async function parseHwpx(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const texts = [];
  const images = [];
  const prvText = zip.file("Preview/PrvText.txt") || zip.file("Preview/PrvText");
  if (prvText) {
    const raw = (await prvText.async("string")).replace(/^\uFEFF/, "").trim();
    if (raw) texts.push(raw);
  }
  const prvImg = zip.file("Preview/PrvImage") || zip.file("Preview/PrvImage.png") || zip.file("Preview/PrvImage.jpg");
  if (prvImg) {
    const buf = await prvImg.async("nodebuffer");
    const mime = mimeFromNameOrBytes(prvImg.name, buf);
    if (mime && buf.length < 6 * 1024 * 1024) images.push(`data:${mime};base64,${buf.toString("base64")}`);
  }
  if (!texts.length) {
    const sections = Object.keys(zip.files).filter((n) => /Contents\/section\d+\.xml$/i.test(n)).sort();
    for (const name of sections.slice(0, 8)) {
      const xml = await zip.file(name).async("string");
      const found = [...xml.matchAll(/<hp:t\b[^>]*>([^<]*)<\/hp:t>/g)].map((m) => decodeXml(m[1]).trim()).filter(Boolean);
      texts.push(...found);
      if (texts.length > 80) break;
    }
  }
  if (!texts.length && !images.length) return [];
  return [{ texts: texts.slice(0, 80), images }];
}

function previewKind(ext) {
  const e = String(ext || "").toLowerCase();
  if (e === "pdf") return "pdf";
  if (TEXT_EXTS.has(e)) return "text";
  if (SLIDE_EXTS.has(e) || e === "ppt" || e === "pps" || e === "pot" || e === "hwp" || e === "hwpx" || e === "hwt") {
    return "office";
  }
  return "none";
}

function previewMime(ext) {
  const e = String(ext || "").toLowerCase();
  if (e === "pdf") return "application/pdf";
  if (TEXT_EXTS.has(e)) return "text/plain; charset=utf-8";
  if (e === "pptx" || e === "ppsx" || e === "potx") {
    return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  }
  return "application/octet-stream";
}

async function buildSlides(ext, buffer) {
  const e = String(ext || "").toLowerCase();
  if (SLIDE_EXTS.has(e)) return parsePptx(buffer);
  if (e === "hwpx") return parseHwpx(buffer);
  return [];
}

module.exports = { previewKind, previewMime, buildSlides };
