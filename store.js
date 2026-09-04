"use strict";

const fs = require("fs");
const path = require("path");

function emptyMeta() {
  return { files: [], labels: {}, sheet: { rows: 40, cols: 10, cells: {} } };
}

function parseRepo(raw) {
  const text = String(raw || "").trim()
    .replace(/^https?:\/\/github\.com\//i, "")
    .replace(/\.git$/i, "")
    .replace(/\/+$/, "");
  const parts = text.split("/").filter(Boolean);
  if (parts.length < 2) return null;
  return { owner: parts[0], repo: parts[1] };
}

function createLocalStore({ dataDir, uploadDir, metaPath }) {
  fs.mkdirSync(uploadDir, { recursive: true });
  fs.mkdirSync(dataDir, { recursive: true });

  return {
    kind: "local",
    persist: true,
    maxFileBytes: 200 * 1024 * 1024,

    async loadMeta() {
      try {
        const raw = fs.readFileSync(metaPath, "utf8").replace(/^\uFEFF/, "");
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed.files)) return parsed;
      } catch (_) {}
      return emptyMeta();
    },

    async saveMeta(next) {
      const tmp = metaPath + ".tmp";
      fs.writeFileSync(tmp, JSON.stringify(next));
      try {
        fs.renameSync(tmp, metaPath);
      } catch (_) {
        fs.copyFileSync(tmp, metaPath);
        fs.unlinkSync(tmp);
      }
    },

    async putFile(_stored, _srcPath) {},

    localPath(stored) {
      const abs = path.resolve(uploadDir, path.basename(String(stored || "")));
      const root = path.resolve(uploadDir);
      if (abs !== root && !abs.startsWith(root + path.sep)) return null;
      return abs;
    },

    async deleteFile(stored) {
      const abs = this.localPath(stored);
      if (abs) fs.unlink(abs, () => {});
    },

    async readFile(stored) {
      const abs = this.localPath(stored);
      if (!abs || !fs.existsSync(abs)) return null;
      return fs.readFileSync(abs);
    },

    sendFile(res, stored, downloadName) {
      const abs = this.localPath(stored);
      if (!abs || !fs.existsSync(abs)) return false;
      res.download(abs, downloadName);
      return true;
    },
  };
}

function createGithubStore({ token, repo: repoRaw, branch }) {
  const parsed = parseRepo(repoRaw);
  if (!parsed) throw new Error("GITHUB_REPO 형식이 올바르지 않습니다. 예: owner/name");
  const { owner, repo } = parsed;
  const ref = branch || "storage";
  let metaSha = null;
  let ready = false;

  async function api(method, urlPath, body, headers) {
    const res = await fetch(`https://api.github.com${urlPath}`, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "daon-writing-files",
        ...(headers || {}),
      },
      body: body == null ? undefined : JSON.stringify(body),
    });
    return res;
  }

  async function json(method, urlPath, body) {
    const res = await api(method, urlPath, body);
    const data = await res.json().catch(() => ({}));
    return { ok: res.ok, status: res.status, data };
  }

  function contentsPath(filePath) {
    return `/repos/${owner}/${repo}/contents/${filePath.split("/").map(encodeURIComponent).join("/")}`;
  }

  async function ensureBranch() {
    if (ready) return;
    const exists = await api("GET", `/repos/${owner}/${repo}/git/ref/heads/${encodeURIComponent(ref)}`);
    if (exists.ok) {
      ready = true;
      return;
    }
    const info = await json("GET", `/repos/${owner}/${repo}`);
    if (!info.ok) {
      const err = new Error(info.data.message || "GitHub 저장소에 접근할 수 없습니다.");
      err.status = info.status;
      throw err;
    }
    const def = info.data.default_branch || "main";
    const head = await json("GET", `/repos/${owner}/${repo}/git/ref/heads/${encodeURIComponent(def)}`);
    if (!head.ok || !head.data.object || !head.data.object.sha) {
      throw new Error("기본 브랜치를 찾지 못했습니다.");
    }
    const created = await json("POST", `/repos/${owner}/${repo}/git/refs`, {
      ref: `refs/heads/${ref}`,
      sha: head.data.object.sha,
    });
    if (!created.ok && created.status !== 422) {
      throw new Error(created.data.message || "storage 브랜치를 만들지 못했습니다.");
    }
    ready = true;
  }

  async function getContents(filePath) {
    const res = await json("GET", `${contentsPath(filePath)}?ref=${encodeURIComponent(ref)}`);
    if (res.status === 404) return null;
    if (!res.ok) {
      const err = new Error(res.data.message || "GitHub에서 파일을 읽지 못했습니다.");
      err.status = res.status;
      throw err;
    }
    return res.data;
  }

  async function putContents(filePath, textOrB64, sha, message, alreadyB64) {
    const content = alreadyB64 ? textOrB64 : Buffer.from(textOrB64).toString("base64");
    const payload = { message, content, branch: ref };
    if (sha) payload.sha = sha;
    const res = await json("PUT", contentsPath(filePath), payload);
    if (!res.ok) {
      const err = new Error(res.data.message || "GitHub에 저장하지 못했습니다.");
      err.status = res.status;
      throw err;
    }
    return res.data;
  }

  return {
    kind: "github",
    persist: true,
    maxFileBytes: 80 * 1024 * 1024,

    async loadMeta() {
      await ensureBranch();
      const file = await getContents("storage/meta.json");
      if (!file || !file.content) {
        metaSha = file && file.sha ? file.sha : null;
        return emptyMeta();
      }
      metaSha = file.sha || null;
      const raw = Buffer.from(file.content.replace(/\n/g, ""), "base64").toString("utf8");
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed.files)) return parsed;
      return emptyMeta();
    },

    async saveMeta(next) {
      await ensureBranch();
      for (let i = 0; i < 4; i++) {
        try {
          const saved = await putContents("storage/meta.json", JSON.stringify(next), metaSha, "update daon files");
          metaSha = saved.content && saved.content.sha ? saved.content.sha : metaSha;
          return;
        } catch (err) {
          if (err.status === 409 || err.status === 422) {
            const file = await getContents("storage/meta.json");
            metaSha = file && file.sha ? file.sha : null;
            continue;
          }
          throw err;
        }
      }
      throw new Error("저장에 실패했습니다. 잠시 후 다시 시도해 주세요.");
    },

    async putFile(stored, srcPath) {
      await ensureBranch();
      const buf = fs.readFileSync(srcPath);
      const b64 = buf.toString("base64");
      await putContents(`storage/files/${path.basename(stored)}`, b64, null, `upload ${path.basename(stored)}`, true);
      fs.unlink(srcPath, () => {});
    },

    async deleteFile(stored) {
      await ensureBranch();
      const file = await getContents(`storage/files/${path.basename(stored)}`);
      if (!file || !file.sha) return;
      await json("DELETE", contentsPath(`storage/files/${path.basename(stored)}`), {
        message: `delete ${path.basename(stored)}`,
        sha: file.sha,
        branch: ref,
      });
    },

    async readFile(stored) {
      await ensureBranch();
      const apiRes = await api(
        "GET",
        `${contentsPath(`storage/files/${path.basename(stored)}`)}?ref=${encodeURIComponent(ref)}`,
        null,
        { Accept: "application/vnd.github.raw" }
      );
      if (!apiRes.ok) return null;
      return Buffer.from(await apiRes.arrayBuffer());
    },

    async sendFile(res, stored, downloadName) {
      const buf = await this.readFile(stored);
      if (!buf) return false;
      res.setHeader("Cache-Control", "private, max-age=120");
      res.setHeader("Content-Type", "application/octet-stream");
      res.setHeader("Content-Length", String(buf.length));
      res.attachment(downloadName);
      res.end(buf);
      return true;
    },
  };
}

function createStore(opts) {
  const token = String(opts.token || "").trim();
  const repo = String(opts.repo || "").trim();
  if (token && repo) {
    return createGithubStore({ token, repo, branch: opts.branch });
  }
  return createLocalStore(opts);
}

module.exports = { createStore, emptyMeta, parseRepo };
