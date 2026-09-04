(() => {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const menuList = $("menu-list");
  const btnUpload = $("btn-upload");
  const paneUpload = $("pane-upload");
  const paneFiles = $("pane-files");
  const filesTitle = $("files-title");
  const filesMeta = $("files-meta");
  const filesWrap = $("files-wrap");
  const paneSheet = $("pane-sheet");
  const sheetEl = $("sheet");
  const sheetScroll = $("sheet-scroll");
  const sheetEditor = $("sheet-editor");
  const sheetAddr = $("sheet-addr");
  const sheetTitle = $("sheet-title");
  const sheetFiles = $("sheet-files");
  const statusRows = $("status-rows");
  const statusSum = $("status-sum");
  const drop = $("drop");
  const fileInput = $("file-input");
  const modal = $("modal");
  const modalSub = $("modal-sub");
  const modalMenus = $("modal-menus");
  const pendingList = $("pending-list");
  const toastEl = $("toast");
  const search = $("search");
  const main = $("main");
  const progress = $("progress");
  const progressBar = $("progress-bar");
  const offline = $("offline");
  const login = $("login");
  const loginForm = $("login-form");
  const loginPass = $("login-pass");
  const loginErr = $("login-err");
  const navHint = $("nav-hint");
  const schedList = $("sched-list");
  const schedPop = $("sched-pop");
  const schedForm = $("sched-form");
  const schedDate = $("sched-date");
  const schedTime = $("sched-time");
  const schedText = $("sched-text");
  const schedDelete = $("sched-delete");
  const schedTitle = $("sched-pop-title");
  const onlineCountEl = $("online-count");

  const ICONS = {
    note: `<svg viewBox="0 0 20 20" width="16" height="16"><path fill="currentColor" d="M5 3.4A1.4 1.4 0 0 0 3.6 4.8v10.4A1.4 1.4 0 0 0 5 16.6h10a1.4 1.4 0 0 0 1.4-1.4V7.1L12.9 3.4H5Zm1.6 4.2h4.2a.7.7 0 1 1 0 1.4H6.6a.7.7 0 1 1 0-1.4Zm0 3h6.8a.7.7 0 1 1 0 1.4H6.6a.7.7 0 1 1 0-1.4Z"/></svg>`,
    edit: `<svg viewBox="0 0 20 20" width="16" height="16"><path fill="currentColor" d="M13.6 3.3a1.6 1.6 0 0 1 2.2 0l.9.9a1.6 1.6 0 0 1 0 2.2l-8.8 8.8H5.2v-2.7l8.4-8.4Zm-9 12.9h10.8a.7.7 0 1 1 0 1.4H4.6a.7.7 0 1 1 0-1.4Z"/></svg>`,
    folder: `<svg viewBox="0 0 20 20" width="16" height="16"><path fill="currentColor" d="M3.6 5.2A1.6 1.6 0 0 1 5.2 3.6h3.1l1.4 1.6h5.1A1.6 1.6 0 0 1 16.4 6.8v7.6a1.6 1.6 0 0 1-1.6 1.6H5.2A1.6 1.6 0 0 1 3.6 14.4V5.2Z"/></svg>`,
    doc: `<svg viewBox="0 0 20 20" width="16" height="16"><path fill="currentColor" d="M5 3.5A1.5 1.5 0 0 0 3.5 5v10A1.5 1.5 0 0 0 5 16.5h10A1.5 1.5 0 0 0 16.5 15V7.2L12.8 3.5H5Zm7.2.7 3.1 3.1H12.2a.5.5 0 0 1-.5-.5V4.2Z"/></svg>`,
  };

  const ALLOWED = new Set([
    "ppt", "pptx", "pot", "potx", "pps", "ppsx",
    "hwp", "hwpx", "hwt", "hml", "hwpml", "owpml",
    "cell", "show", "nxl", "nxls", "txt", "text", "pdf",
  ]);

  const KIND = {
    ppt: "ppt", pptx: "ppt", pot: "ppt", potx: "ppt", pps: "ppt", ppsx: "ppt",
    hwp: "hwp", hwpx: "hwp", hwt: "hwp", hml: "hwp", hwpml: "hwp", owpml: "hwp",
    txt: "txt", text: "txt",
    cell: "cell", nxl: "cell", nxls: "cell",
    show: "show",
    pdf: "pdf",
  };

  const cache = new Map();
  let menus = [];
  let counts = {};
  let latest = {};
  let view = "upload";
  let pending = [];
  let shownFiles = [];
  let toastTimer = 0;
  let uploading = false;

  const extOf = (name) => {
    const i = name.lastIndexOf(".");
    return i >= 0 ? name.slice(i + 1).toLowerCase() : "";
  };

  const formatSize = (n) => {
    if (n < 1024) return n + " B";
    if (n < 1048576) return (n / 1024).toFixed(n < 10240 ? 1 : 0) + " KB";
    return (n / 1048576).toFixed(n < 10485760 ? 1 : 0) + " MB";
  };

  const formatWhen = (iso) => {
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return "";
    return `${d.getFullYear()}. ${d.getMonth() + 1}. ${d.getDate()}.`;
  };

  const formatMenuWhen = (iso) => {
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return "";
    const h = d.getHours();
    const ap = h < 12 ? "오전" : "오후";
    const hour12 = h % 12 === 0 ? 12 : h % 12;
    return `(${d.getFullYear()}.${d.getMonth() + 1}.${d.getDate()}. ${ap} ${hour12}시 ${d.getMinutes()}분)`;
  };

  const newestFirst = (files) => files.slice().sort((a, b) => String(b.uploadedAt || "").localeCompare(String(a.uploadedAt || "")));
  const STATUS_SKIP = new Set(["memo", "merged", "refs"]);
  const formatStatusWhen = (iso) => {
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return "아직 없음";
    const h = d.getHours();
    const ap = h < 12 ? "오전" : "오후";
    const hour12 = h % 12 === 0 ? 12 : h % 12;
    const min = String(d.getMinutes()).padStart(2, "0");
    return `${d.getFullYear()}.${d.getMonth() + 1}.${d.getDate()}. ${ap} ${hour12}시 ${min}분`;
  };

  const applyStats = (data) => {
    if (data.counts) counts = data.counts;
    if (data.latest) latest = data.latest;
  };

  let events = [];
  let editingEventId = "";
  const sessionId = (() => {
    try {
      let id = sessionStorage.getItem("daon-sid");
      if (!id) {
        id = (crypto.randomUUID ? crypto.randomUUID() : Date.now() + "-" + Math.random()).replace(/[^a-zA-Z0-9_-]/g, "").slice(0, 32);
        if (id.length < 8) id = (id + "sessionid").slice(0, 16);
        sessionStorage.setItem("daon-sid", id);
      }
      return id;
    } catch (_) {
      return "s" + Date.now() + Math.random().toString(36).slice(2, 8);
    }
  })();

  const setOnline = (n) => {
    if (onlineCountEl) onlineCountEl.textContent = String(Math.max(1, Number(n) || 1));
  };

  const eventStamp = (ev) => ev.date + "T" + (ev.time || "00:00");

  const daysUntil = (ev) => {
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const d = new Date(ev.date + "T00:00:00");
    if (Number.isNaN(d.getTime())) return 999;
    return Math.round((d - start) / 86400000);
  };

  const formatEventWhen = (ev) => {
    const [y, m, d] = ev.date.split("-");
    const date = `${Number(m)}. ${Number(d)}.`;
    return ev.time ? `${date} ${ev.time}` : date;
  };

  const visibleEvents = () => events
    .filter((ev) => daysUntil(ev) >= 0)
    .slice()
    .sort((a, b) => eventStamp(a).localeCompare(eventStamp(b)));

  const renderEvents = () => {
    if (!schedList) return;
    const frag = document.createDocumentFragment();
    for (const ev of visibleEvents()) {
      const soon = daysUntil(ev) <= 3;
      const wrap = document.createElement("div");
      wrap.className = "sched-item" + (soon ? " is-soon" : "");
      wrap.innerHTML =
        `<button type="button" class="sched-main" data-sched="${ev.id}"></button>` +
        `<button type="button" class="sched-x" data-sched-del="${ev.id}" aria-label="일정 삭제">×</button>`;
      const mainBtn = wrap.querySelector(".sched-main");
      const when = document.createElement("span");
      when.className = "sched-when";
      when.textContent = formatEventWhen(ev);
      mainBtn.appendChild(when);
      if (ev.text) {
        const label = document.createElement("span");
        label.className = "sched-text";
        label.textContent = ev.text;
        mainBtn.appendChild(label);
      }
      frag.appendChild(wrap);
    }
    schedList.replaceChildren(frag);
  };

  const closeSched = () => {
    if (schedPop) schedPop.classList.add("is-hidden");
    editingEventId = "";
  };

  const openSched = (ev) => {
    editingEventId = ev && ev.id ? ev.id : "";
    schedTitle.textContent = editingEventId ? "일정 수정" : "일정 등록";
    schedDate.value = ev && ev.date ? ev.date : "";
    schedTime.value = ev && ev.time ? ev.time : "";
    schedText.value = ev && ev.text ? ev.text : "";
    schedDelete.classList.toggle("is-hidden", !editingEventId);
    schedPop.classList.remove("is-hidden");
    schedDate.focus();
  };

  const applyEvents = (list) => {
    if (!Array.isArray(list)) return;
    events = list;
    renderEvents();
  };

  const toast = (msg) => {
    toastEl.hidden = false;
    toastEl.textContent = msg;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { toastEl.hidden = true; }, 2600);
  };

  const setProgress = (ratio) => {
    const on = ratio != null;
    progress.classList.toggle("is-hidden", !on);
    progressBar.style.width = on ? Math.round(ratio * 100) + "%" : "0";
  };

  const colName = (index) => {
    let n = index + 1;
    let name = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      name = String.fromCharCode(65 + rem) + name;
      n = Math.floor((n - 1) / 26);
    }
    return name;
  };

  const cellKey = (r, c) => r + "," + c;

  const Sheet = {
    rows: 40,
    cols: 10,
    cells: {},
    sel: { r: 0, c: 0 },
    anchor: { r: 0, c: 0 },
    dragging: false,
    editing: false,
    editFrom: "",
    undo: [],
    redo: [],
    saveTimer: 0,
    loaded: false,

    key(r, c) { return cellKey(r, c); },
    get(r, c) { return this.cells[this.key(r, c)] || ""; },
    addr(r, c) { return colName(c) + (r + 1); },

    td(r, c) {
      return sheetEl.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
    },

    clamp(r, c) {
      return {
        r: Math.max(0, Math.min(this.rows - 1, r)),
        c: Math.max(0, Math.min(this.cols - 1, c)),
      };
    },

    range() {
      return {
        r1: Math.min(this.anchor.r, this.sel.r),
        c1: Math.min(this.anchor.c, this.sel.c),
        r2: Math.max(this.anchor.r, this.sel.r),
        c2: Math.max(this.anchor.c, this.sel.c),
      };
    },

    paintSel() {
      sheetEl.querySelectorAll("td.is-sel, td.is-range").forEach((el) => {
        el.classList.remove("is-sel", "is-range");
      });
      const { r1, c1, r2, c2 } = this.range();
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          const td = this.td(r, c);
          if (td) td.classList.add("is-range");
        }
      }
      const active = this.td(this.sel.r, this.sel.c);
      if (active) active.classList.add("is-sel");
      sheetAddr.textContent = r1 === r2 && c1 === c2
        ? this.addr(r1, c1)
        : this.addr(r1, c1) + ":" + this.addr(r2, c2);
    },

    render() {
      const head = document.createElement("thead");
      const hr = document.createElement("tr");
      hr.appendChild(Object.assign(document.createElement("th"), { className: "rh" }));
      for (let c = 0; c < this.cols; c++) {
        const th = document.createElement("th");
        th.textContent = colName(c);
        hr.appendChild(th);
      }
      head.appendChild(hr);

      const body = document.createElement("tbody");
      for (let r = 0; r < this.rows; r++) {
        const tr = document.createElement("tr");
        const rh = document.createElement("th");
        rh.className = "rh";
        rh.textContent = String(r + 1);
        tr.appendChild(rh);
        for (let c = 0; c < this.cols; c++) {
          const td = document.createElement("td");
          td.dataset.r = String(r);
          td.dataset.c = String(c);
          td.textContent = this.get(r, c);
          tr.appendChild(td);
        }
        body.appendChild(tr);
      }
      sheetEl.replaceChildren(head, body);
      this.paintSel();
    },

    grow(r, c) {
      let changed = false;
      if (r >= this.rows - 1 && this.rows < 200) {
        this.rows = Math.min(200, this.rows + 10);
        changed = true;
      }
      if (c >= this.cols - 1 && this.cols < 40) {
        this.cols = Math.min(40, this.cols + 3);
        changed = true;
      }
      if (changed) this.render();
    },

    scheduleSave() {
      clearTimeout(this.saveTimer);
      this.saveTimer = setTimeout(() => this.save(), 400);
    },

    async save() {
      try {
        await fetch("api/sheet", {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ rows: this.rows, cols: this.cols, cells: this.cells }),
        });
      } catch (_) {}
    },

    write(r, c, next) {
      if (next) this.cells[this.key(r, c)] = next;
      else delete this.cells[this.key(r, c)];
      const td = this.td(r, c);
      if (td && !this.editing) td.textContent = next;
    },

    apply(r, c, next, record) {
      const prev = this.get(r, c);
      if (prev === next) return false;
      this.write(r, c, next);
      if (record) {
        this.undo.push({ r, c, from: prev, to: next });
        if (this.undo.length > 200) this.undo.shift();
        this.redo.length = 0;
      }
      this.grow(r, c);
      this.scheduleSave();
      return true;
    },

    applyMany(pairs) {
      if (!pairs.length) return;
      let maxR = this.rows - 1;
      let maxC = this.cols - 1;
      for (const p of pairs) {
        maxR = Math.max(maxR, p.r);
        maxC = Math.max(maxC, p.c);
      }
      this.grow(maxR, maxC);
      const many = [];
      for (const p of pairs) {
        const from = this.get(p.r, p.c);
        if (from === p.to) continue;
        this.write(p.r, p.c, p.to);
        many.push({ r: p.r, c: p.c, from, to: p.to });
      }
      if (!many.length) return;
      this.undo.push({
        many,
        sel: { r: this.sel.r, c: this.sel.c },
        anchor: { r: this.anchor.r, c: this.anchor.c },
      });
      if (this.undo.length > 200) this.undo.shift();
      this.redo.length = 0;
      this.scheduleSave();
    },

    restoreSel(item) {
      if (item.sel) this.sel = { r: item.sel.r, c: item.sel.c };
      if (item.anchor) this.anchor = { r: item.anchor.r, c: item.anchor.c };
      else this.anchor = { r: this.sel.r, c: this.sel.c };
      this.paintSel();
    },

    undoLast() {
      const item = this.undo.pop();
      if (!item) return;
      this.redo.push(item);
      if (item.many) {
        for (const ch of item.many) this.write(ch.r, ch.c, ch.from);
        this.restoreSel(item);
        this.scheduleSave();
        return;
      }
      this.apply(item.r, item.c, item.from, false);
      this.select(item.r, item.c);
    },

    redoLast() {
      const item = this.redo.pop();
      if (!item) return;
      this.undo.push(item);
      if (item.many) {
        for (const ch of item.many) this.write(ch.r, ch.c, ch.to);
        this.restoreSel(item);
        this.scheduleSave();
        return;
      }
      this.apply(item.r, item.c, item.to, false);
      this.select(item.r, item.c);
    },

    copyText() {
      const { r1, c1, r2, c2 } = this.range();
      const lines = [];
      for (let r = r1; r <= r2; r++) {
        const cols = [];
        for (let c = c1; c <= c2; c++) {
          cols.push(this.get(r, c).replace(/\t/g, " ").replace(/\r?\n/g, "\n"));
        }
        lines.push(cols.join("\t"));
      }
      return lines.join("\n");
    },

    parseTsv(text) {
      const raw = String(text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
      const body = raw.endsWith("\n") ? raw.slice(0, -1) : raw;
      if (!body) return [[""]];
      return body.split("\n").map((line) => line.split("\t"));
    },

    paste(text) {
      const grid = this.parseTsv(text);
      const { r1, c1, r2, c2 } = this.range();
      const one = grid.length === 1 && grid[0].length === 1;
      const fill = one && (r1 !== r2 || c1 !== c2);
      const pairs = [];
      if (fill) {
        for (let r = r1; r <= r2; r++) {
          for (let c = c1; c <= c2; c++) pairs.push({ r, c, to: grid[0][0] });
        }
      } else {
        for (let i = 0; i < grid.length; i++) {
          for (let j = 0; j < grid[i].length; j++) {
            pairs.push({ r: r1 + i, c: c1 + j, to: grid[i][j] });
          }
        }
        const last = pairs[pairs.length - 1];
        this.sel = this.clamp(last.r, last.c);
        this.anchor = { r: r1, c: c1 };
      }
      this.applyMany(pairs);
      this.paintSel();
    },

    clearRange() {
      const { r1, c1, r2, c2 } = this.range();
      const pairs = [];
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) pairs.push({ r, c, to: "" });
      }
      this.applyMany(pairs);
    },

    placeEditor(td) {
      const left = td.offsetLeft;
      const top = td.offsetTop;
      sheetEditor.style.left = left + "px";
      sheetEditor.style.top = top + "px";
      sheetEditor.style.width = Math.max(td.offsetWidth, 140) + "px";
      sheetEditor.style.height = "auto";
      sheetEditor.style.minHeight = Math.max(td.offsetHeight, 28) + "px";
      sheetEditor.style.height = Math.max(td.offsetHeight, sheetEditor.scrollHeight) + "px";
    },

    fitEditor() {
      sheetEditor.style.height = "auto";
      sheetEditor.style.height = Math.max(28, sheetEditor.scrollHeight) + "px";
    },

    startEdit(selectAll) {
      const { r, c } = this.sel;
      const td = this.td(r, c);
      if (!td) return;
      this.editing = true;
      this.editFrom = this.get(r, c);
      td.textContent = "";
      sheetEditor.value = this.editFrom;
      sheetEditor.classList.remove("is-hidden");
      this.placeEditor(td);
      sheetEditor.focus();
      if (selectAll) sheetEditor.select();
      else sheetEditor.setSelectionRange(sheetEditor.value.length, sheetEditor.value.length);
      this.fitEditor();
    },

    commit(moveR, moveC) {
      if (!this.editing) return;
      const { r, c } = this.sel;
      const next = sheetEditor.value.replace(/\u00a0/g, " ");
      this.editing = false;
      sheetEditor.classList.add("is-hidden");
      this.apply(r, c, next, true);
      const td = this.td(r, c);
      if (td) td.textContent = this.get(r, c);
      if (moveR || moveC) this.select(r + (moveR || 0), c + (moveC || 0));
      else this.paintSel();
    },

    cancelEdit() {
      if (!this.editing) return;
      this.editing = false;
      sheetEditor.classList.add("is-hidden");
      const td = this.td(this.sel.r, this.sel.c);
      if (td) td.textContent = this.editFrom;
    },

    blur() {
      if (this.editing) this.commit(0, 0);
    },

    select(r, c, extend) {
      const next = this.clamp(r, c);
      this.sel.r = next.r;
      this.sel.c = next.c;
      if (!extend) {
        this.anchor.r = next.r;
        this.anchor.c = next.c;
      }
      this.paintSel();
      const td = this.td(this.sel.r, this.sel.c);
      if (td) td.scrollIntoView({ block: "nearest", inline: "nearest" });
    },

    selectAll() {
      this.anchor = { r: 0, c: 0 };
      this.sel = { r: this.rows - 1, c: this.cols - 1 };
      this.paintSel();
    },

    insertNewline() {
      const start = sheetEditor.selectionStart;
      const end = sheetEditor.selectionEnd;
      const value = sheetEditor.value;
      sheetEditor.value = value.slice(0, start) + "\n" + value.slice(end);
      const pos = start + 1;
      sheetEditor.setSelectionRange(pos, pos);
      this.fitEditor();
    },

    async load() {
      const res = await fetch("api/sheet", { cache: "no-store" });
      const data = await res.json();
      this.rows = data.rows || 40;
      this.cols = data.cols || 10;
      this.cells = data.cells || {};
      this.loaded = true;
      this.render();
    },
  };

  const setView = (next) => {
    view = next;
    const isUpload = next === "upload";
    const isSheet = next === "memo";
    paneUpload.classList.toggle("is-hidden", !isUpload);
    paneFiles.classList.toggle("is-hidden", isUpload || isSheet);
    paneSheet.classList.toggle("is-hidden", !isSheet);
    btnUpload.classList.toggle("is-active", isUpload);
    for (const btn of menuList.querySelectorAll("[data-menu]")) {
      btn.classList.toggle("is-active", btn.dataset.menu === next);
    }
    if (isUpload) search.value = "";
    if (!isSheet) Sheet.blur();
  };

  const renderStatus = () => {
    if (!statusRows || !statusSum) return;
    const rows = menus.filter((m) => !STATUS_SKIP.has(m.id));
    const nums = rows.map((m) => counts[m.id] || 0);
    const max = Math.max(1, ...nums);
    const done = nums.filter((n) => n > 0).length;
    const total = nums.reduce((a, b) => a + b, 0);
    statusSum.textContent = `${rows.length}명 중 ${done}명 제출 · 파일 ${total}개`;
    const frag = document.createDocumentFragment();
    for (const m of rows) {
      const n = counts[m.id] || 0;
      const when = n ? formatStatusWhen(latest[m.id]) : "아직 없음";
      const pct = n ? Math.max(10, Math.round((n / max) * 100)) : 0;
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "status-row" + (n ? " is-on" : "");
      btn.dataset.menu = m.id;
      btn.title = `${m.label} · ${n}개 · ${when}`;
      btn.innerHTML =
        `<span class="status-name"><i class="status-dot" aria-hidden="true"></i><span class="status-label"></span></span>` +
        `<span class="status-bar" aria-hidden="true"><i style="width:${pct}%"></i></span>` +
        `<span class="status-n">${n}개</span>` +
        `<span class="status-when"></span>`;
      btn.querySelector(".status-label").textContent = m.label;
      btn.querySelector(".status-when").textContent = when;
      frag.appendChild(btn);
    }
    statusRows.replaceChildren(frag);
  };

  const renderMenus = () => {
    renderStatus();
    const frag = document.createDocumentFragment();
    for (const m of menus) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "menu-item" + (view === m.id ? " is-active" : "");
      btn.dataset.menu = m.id;
      const n = counts[m.id] || 0;
      const when = formatMenuWhen(latest[m.id]);
      btn.innerHTML =
        `<span class="menu-ico" aria-hidden="true">${ICONS[m.icon] || ICONS.doc}</span>` +
        `<span class="menu-text"><span class="menu-label"></span>${when ? `<span class="menu-date"></span>` : ""}</span>` +
        (n ? `<span class="menu-count">${n}</span>` : "");
      btn.querySelector(".menu-label").textContent = m.label;
      const dateEl = btn.querySelector(".menu-date");
      if (dateEl) dateEl.textContent = when;
      btn.title = when
        ? `최근 파일 ${when} · 더블클릭하면 제목을 바꿀 수 있습니다`
        : "더블클릭하면 제목을 바꿀 수 있습니다";
      frag.appendChild(btn);
    }
    menuList.replaceChildren(frag);
  };

  const fileBadge = (ext) => {
    const kind = KIND[ext] || "etc";
    return `<div class="file-badge ${kind}">${(ext || "FILE").slice(0, 4).toUpperCase()}</div>`;
  };

  const renderSheetFiles = (files) => {
    files = newestFirst(files);
    if (!files.length) {
      sheetFiles.replaceChildren();
      return;
    }
    const frag = document.createDocumentFragment();
    for (const f of files) {
      const row = document.createElement("article");
      row.className = "file-row";
      row.innerHTML =
        fileBadge(f.ext) +
        `<div><div class="file-name"></div><div class="file-sub">${formatSize(f.size)}</div></div>` +
        `<a class="btn-dl" href="api/download/${encodeURIComponent(f.id)}">내려받기</a>` +
        `<button type="button" class="btn-del" data-del="${f.id}" aria-label="삭제">삭제</button>`;
      row.querySelector(".file-name").textContent = f.name;
      frag.appendChild(row);
    }
    sheetFiles.replaceChildren(frag);
  };

  const renderFiles = (menuId, files) => {
    files = newestFirst(files);
    shownFiles = files;
    const menu = menus.find((m) => m.id === menuId);
    if (menuId === "memo") {
      if (menu) sheetTitle.textContent = menu.label;
      renderSheetFiles(files);
      return;
    }
    const q = search.value.trim().toLowerCase();
    const visible = q ? files.filter((f) => f.name.toLowerCase().includes(q)) : files;
    filesTitle.textContent = menu ? menu.label : "";
    filesMeta.textContent = files.length
      ? `${files.length}개 파일 · 누구나 내려받을 수 있습니다`
      : "아직 올라온 파일이 없습니다";

    if (!visible.length) {
      filesWrap.innerHTML = `<div class="empty">${files.length ? "찾는 파일이 없습니다." : "이 메뉴에는 첨부파일이 없습니다.<br />왼쪽에서 업로드를 눌러 파일을 올려 주세요."}</div>`;
      return;
    }

    const frag = document.createDocumentFragment();
    for (const f of visible) {
      const row = document.createElement("article");
      row.className = "file-row";
      row.innerHTML =
        fileBadge(f.ext) +
        `<div><div class="file-name"></div><div class="file-sub">${formatSize(f.size)}</div></div>` +
        `<div class="file-when file-sub">${formatWhen(f.uploadedAt)}</div>` +
        `<a class="btn-dl" href="api/download/${encodeURIComponent(f.id)}">내려받기</a>` +
        `<button type="button" class="btn-del" data-del="${f.id}" aria-label="삭제">삭제</button>`;
      row.querySelector(".file-name").textContent = f.name;
      frag.appendChild(row);
    }
    filesWrap.replaceChildren(frag);
  };

  const loadFiles = async (menuId, force) => {
    if (!force && cache.has(menuId)) {
      renderFiles(menuId, cache.get(menuId));
      return;
    }
    const res = await fetch("api/files?menu=" + encodeURIComponent(menuId), { cache: "no-store" });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "목록을 불러오지 못했습니다.");
    cache.set(menuId, data.files);
    renderFiles(menuId, data.files);
  };

  const openMenu = async (menuId, force) => {
    setView(menuId);
    if (menuId === "memo") {
      const menu = menus.find((m) => m.id === menuId);
      if (menu) sheetTitle.textContent = menu.label;
      try {
        await Sheet.load();
      } catch (_) {
        Sheet.render();
      }
    } else {
      filesWrap.innerHTML = `<div class="empty">불러오는 중…</div>`;
    }
    try {
      await loadFiles(menuId, force);
    } catch (e) {
      if (menuId !== "memo") filesWrap.innerHTML = `<div class="empty">${e.message}</div>`;
    }
  };

  const closeModal = () => {
    if (uploading) return;
    pending = [];
    fileInput.value = "";
    modal.classList.add("is-hidden");
  };

  const openModal = (files) => {
    pending = files;
    modalSub.textContent = files.length === 1
      ? "아래 메뉴를 누르면 그 칸에 저장됩니다."
      : `${files.length}개 파일을 어느 메뉴에 둘까요?`;
    pendingList.replaceChildren();
    for (const file of files.slice(0, 8)) {
      const chip = document.createElement("span");
      chip.className = "chip";
      chip.textContent = `${file.name} · ${formatSize(file.size)}`;
      pendingList.appendChild(chip);
    }
    if (files.length > 8) {
      const more = document.createElement("span");
      more.className = "chip";
      more.textContent = `외 ${files.length - 8}개`;
      pendingList.appendChild(more);
    }
    const frag = document.createDocumentFragment();
    for (const m of menus) {
      const b = document.createElement("button");
      b.type = "button";
      b.dataset.pick = m.id;
      b.className = view === m.id ? "is-current" : "";
      b.textContent = m.label;
      frag.appendChild(b);
    }
    modalMenus.replaceChildren(frag);
    modal.classList.remove("is-hidden");
  };

  const filterFiles = (list) => {
    const ok = [];
    let bad = 0;
    for (const file of list) {
      if (ALLOWED.has(extOf(file.name))) ok.push(file);
      else bad += 1;
    }
    if (!ok.length) {
      toast("ppt, pptx, pdf, hwp, hwpx, 메모장, 한글오피스 파일만 올릴 수 있습니다.");
      return [];
    }
    if (bad) toast(`지원하지 않는 형식 ${bad}개는 제외했습니다`);
    return ok;
  };

  const postUpload = (menuId, files) => new Promise((resolve, reject) => {
    const body = new FormData();
    body.append("menuId", menuId);
    for (const file of files) body.append("files", file, file.name);
    const xhr = new XMLHttpRequest();
    xhr.open("POST", "api/upload");
    xhr.upload.onprogress = (e) => {
      if (e.lengthComputable) setProgress(e.loaded / e.total);
    };
    xhr.onload = () => {
      let data = {};
      try { data = JSON.parse(xhr.responseText); } catch (_) {}
      if (xhr.status >= 200 && xhr.status < 300) resolve(data);
      else reject(new Error(data.error || "업로드 실패"));
    };
    xhr.onerror = () => reject(new Error("네트워크 오류가 났습니다."));
    xhr.send(body);
  });

  const uploadTo = async (menuId) => {
    if (!pending.length || uploading) return;
    const files = pending;
    uploading = true;
    modal.classList.add("is-hidden");
    setProgress(0.05);
    try {
      const data = await postUpload(menuId, files);
      applyStats(data);
      cache.delete(menuId);
      renderMenus();
      toast(`${files.length}개 파일을 올렸습니다`);
      await openMenu(menuId, true);
    } catch (e) {
      toast(e.message);
    } finally {
      uploading = false;
      pending = [];
      fileInput.value = "";
      setProgress(null);
    }
  };

  const takeFiles = (fileList) => {
    if (uploading) return;
    const files = filterFiles(Array.from(fileList || []));
    if (files.length) openModal(files);
  };

  const deleteFile = async (id) => {
    if (!confirm("이 파일을 삭제할까요?")) return;
    try {
      const res = await fetch("api/files/" + encodeURIComponent(id), { method: "DELETE" });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "삭제 실패");
      applyStats(data);
      cache.delete(view);
      renderMenus();
      await openMenu(view, true);
      toast("삭제했습니다");
    } catch (e) {
      toast(e.message);
    }
  };

  const startRename = (btn) => {
    const id = btn.dataset.menu;
    const menu = menus.find((m) => m.id === id);
    if (!menu || btn.querySelector(".menu-edit")) return;
    const labelEl = btn.querySelector(".menu-label");
    const input = document.createElement("input");
    input.className = "menu-edit";
    input.value = menu.label;
    input.maxLength = 80;
    input.setAttribute("aria-label", "메뉴 제목");
    labelEl.replaceWith(input);
    input.focus();
    input.select();

    let cancelled = false;
    const finish = async (save) => {
      input.onblur = null;
      const next = input.value.replace(/\s+/g, " ").trim();
      if (!save || !next || next === menu.label) {
        renderMenus();
        return;
      }
      try {
        const res = await fetch("api/menus/" + encodeURIComponent(id), {
          method: "PATCH",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ label: next }),
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || "이름 변경 실패");
        menus = data.menus;
        renderMenus();
        if (view === id) {
          filesTitle.textContent = next;
          if (id === "memo") sheetTitle.textContent = next;
        }
        toast("메뉴 이름을 바꿨습니다");
      } catch (e) {
        toast(e.message);
        renderMenus();
      }
    };

    input.addEventListener("click", (e) => e.stopPropagation());
    input.addEventListener("mousedown", (e) => e.stopPropagation());
    input.addEventListener("dblclick", (e) => e.stopPropagation());
    input.addEventListener("keydown", (e) => {
      e.stopPropagation();
      if (e.key === "Enter") {
        e.preventDefault();
        input.blur();
      }
      if (e.key === "Escape") {
        cancelled = true;
        e.preventDefault();
        input.blur();
      }
    });
    input.addEventListener("blur", () => finish(!cancelled));
  };

  if (statusRows) {
    statusRows.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-menu]");
      if (btn) openMenu(btn.dataset.menu);
    });
  }

  menuList.addEventListener("click", (e) => {
    if (e.target.closest(".menu-edit")) return;
    const btn = e.target.closest("[data-menu]");
    if (btn) openMenu(btn.dataset.menu);
  });

  menuList.addEventListener("dblclick", (e) => {
    const btn = e.target.closest("[data-menu]");
    if (btn) {
      e.preventDefault();
      startRename(btn);
    }
  });

  const onFileDel = (e) => {
    const btn = e.target.closest("[data-del]");
    if (btn) deleteFile(btn.dataset.del);
  };
  filesWrap.addEventListener("click", onFileDel);
  sheetFiles.addEventListener("click", onFileDel);

  if ($("sched-add")) $("sched-add").addEventListener("click", () => openSched(null));
  if (schedPop) {
    schedPop.addEventListener("click", (e) => {
      if (e.target === schedPop) closeSched();
    });
  }
  if ($("sched-cancel")) $("sched-cancel").addEventListener("click", closeSched);
  if (schedDelete) {
    schedDelete.addEventListener("click", async () => {
      if (!editingEventId || !confirm("이 일정을 삭제할까요?")) return;
      try {
        const res = await fetch("api/events/" + encodeURIComponent(editingEventId), { method: "DELETE" });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || "삭제 실패");
        applyEvents(data.events);
        closeSched();
      } catch (e) {
        toast(e.message);
      }
    });
  }
  if (schedForm) {
    schedForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      const body = { date: schedDate.value, time: schedTime.value, text: schedText.value };
      try {
        const res = await fetch(editingEventId ? "api/events/" + encodeURIComponent(editingEventId) : "api/events", {
          method: editingEventId ? "PATCH" : "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(body),
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || "저장 실패");
        applyEvents(data.events);
        closeSched();
      } catch (err) {
        toast(err.message);
      }
    });
  }
  if (schedList) {
    schedList.addEventListener("click", async (e) => {
      const del = e.target.closest("[data-sched-del]");
      if (del) {
        e.preventDefault();
        if (!confirm("이 일정을 삭제할까요?")) return;
        try {
          const res = await fetch("api/events/" + encodeURIComponent(del.dataset.schedDel), { method: "DELETE" });
          const data = await res.json();
          if (!res.ok) throw new Error(data.error || "삭제 실패");
          applyEvents(data.events);
        } catch (err) {
          toast(err.message);
        }
        return;
      }
      const btn = e.target.closest("[data-sched]");
      if (!btn) return;
      const ev = events.find((item) => item.id === btn.dataset.sched);
      if (ev) openSched(ev);
    });
  }

  btnUpload.addEventListener("click", () => setView("upload"));
  $("btn-pick").addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", () => takeFiles(fileInput.files));
  $("modal-cancel").addEventListener("click", closeModal);
  modal.addEventListener("click", (e) => { if (e.target === modal) closeModal(); });
  modalMenus.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-pick]");
    if (btn) uploadTo(btn.dataset.pick);
  });
  search.addEventListener("input", () => {
    if (view !== "upload") renderFiles(view, shownFiles);
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      if (schedPop && !schedPop.classList.contains("is-hidden")) {
        e.preventDefault();
        closeSched();
        return;
      }
      if (Sheet.editing) {
        e.preventDefault();
        Sheet.cancelEdit();
        return;
      }
      closeModal();
    }
    if (view !== "memo" || modal.classList.contains("is-hidden") === false) return;
    if (e.target.closest(".menu-edit") || e.target === search) return;

    const meta = e.ctrlKey || e.metaKey;
    if (meta && e.key.toLowerCase() === "z" && !e.shiftKey) {
      if (Sheet.editing && document.activeElement === sheetEditor) return;
      e.preventDefault();
      Sheet.undoLast();
      return;
    }
    if (meta && (e.key.toLowerCase() === "y" || (e.key.toLowerCase() === "z" && e.shiftKey))) {
      e.preventDefault();
      if (Sheet.editing) Sheet.commit(0, 0);
      Sheet.redoLast();
      return;
    }

    if (Sheet.editing) return;
    if (e.isComposing || e.keyCode === 229) return;

    if (meta && e.key.toLowerCase() === "a") {
      e.preventDefault();
      Sheet.selectAll();
      return;
    }
    if (e.key === "Enter") {
      e.preventDefault();
      if (e.shiftKey) Sheet.select(Sheet.sel.r - 1, Sheet.sel.c);
      else Sheet.select(Sheet.sel.r + 1, Sheet.sel.c);
      return;
    }
    if (e.key === "Tab") {
      e.preventDefault();
      Sheet.select(Sheet.sel.r, Sheet.sel.c + (e.shiftKey ? -1 : 1));
      return;
    }
    if (e.key === "ArrowUp") { e.preventDefault(); Sheet.select(Sheet.sel.r - 1, Sheet.sel.c, e.shiftKey); return; }
    if (e.key === "ArrowDown") { e.preventDefault(); Sheet.select(Sheet.sel.r + 1, Sheet.sel.c, e.shiftKey); return; }
    if (e.key === "ArrowLeft") { e.preventDefault(); Sheet.select(Sheet.sel.r, Sheet.sel.c - 1, e.shiftKey); return; }
    if (e.key === "ArrowRight") { e.preventDefault(); Sheet.select(Sheet.sel.r, Sheet.sel.c + 1, e.shiftKey); return; }
    if (e.key === "F2") { e.preventDefault(); Sheet.startEdit(false); return; }
    if (e.key === "Delete" || e.key === "Backspace") {
      e.preventDefault();
      Sheet.clearRange();
      return;
    }
    if (e.key.length === 1 && !meta && !e.altKey) {
      e.preventDefault();
      Sheet.startEdit(true);
      sheetEditor.value = e.key;
      Sheet.fitEditor();
    }
  });

  sheetEl.addEventListener("mousedown", (e) => {
    if (e.button !== 0) return;
    const td = e.target.closest("td");
    if (!td) return;
    e.preventDefault();
    if (Sheet.editing) Sheet.commit(0, 0);
    Sheet.select(Number(td.dataset.r), Number(td.dataset.c), e.shiftKey);
    Sheet.dragging = true;
  });
  document.addEventListener("mousemove", (e) => {
    if (!Sheet.dragging) return;
    const hit = document.elementFromPoint(e.clientX, e.clientY);
    const td = hit && hit.closest && hit.closest("#sheet td");
    if (!td) return;
    Sheet.select(Number(td.dataset.r), Number(td.dataset.c), true);
  });
  document.addEventListener("mouseup", () => {
    Sheet.dragging = false;
  });
  document.addEventListener("copy", (e) => {
    if (view !== "memo" || Sheet.editing || modal.classList.contains("is-hidden") === false) return;
    e.preventDefault();
    e.clipboardData.setData("text/plain", Sheet.copyText());
  });
  document.addEventListener("cut", (e) => {
    if (view !== "memo" || Sheet.editing || modal.classList.contains("is-hidden") === false) return;
    e.preventDefault();
    e.clipboardData.setData("text/plain", Sheet.copyText());
    Sheet.clearRange();
  });
  document.addEventListener("paste", (e) => {
    if (view !== "memo" || Sheet.editing || modal.classList.contains("is-hidden") === false) return;
    e.preventDefault();
    Sheet.paste(e.clipboardData.getData("text/plain"));
  });

  sheetEl.addEventListener("dblclick", (e) => {
    if (e.target.closest("td")) Sheet.startEdit(false);
  });

  sheetEditor.addEventListener("keydown", (e) => {
    if (e.isComposing || e.keyCode === 229) return;
    if (e.key === "Enter") {
      e.preventDefault();
      if (e.shiftKey) Sheet.insertNewline();
      else Sheet.commit(1, 0);
      return;
    }
    if (e.key === "Tab") {
      e.preventDefault();
      Sheet.commit(0, e.shiftKey ? -1 : 1);
      return;
    }
    if (e.key === "Escape") {
      e.preventDefault();
      Sheet.cancelEdit();
    }
  });

  sheetEditor.addEventListener("input", () => Sheet.fitEditor());
  sheetEditor.addEventListener("blur", () => {
    if (Sheet.editing) Sheet.commit(0, 0);
  });

  const onDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
  };
  const hasFiles = (e) => Array.from(e.dataTransfer && e.dataTransfer.types || []).includes("Files");

  main.addEventListener("dragenter", (e) => {
    if (!hasFiles(e)) return;
    onDrag(e);
    main.classList.add("is-over");
    drop.classList.add("is-over");
  });
  main.addEventListener("dragover", (e) => {
    if (!hasFiles(e)) return;
    onDrag(e);
    e.dataTransfer.dropEffect = "copy";
    main.classList.add("is-over");
  });
  main.addEventListener("dragleave", (e) => {
    if (!main.contains(e.relatedTarget)) {
      main.classList.remove("is-over");
      drop.classList.remove("is-over");
    }
  });
  main.addEventListener("drop", (e) => {
    onDrag(e);
    main.classList.remove("is-over");
    drop.classList.remove("is-over");
    takeFiles(e.dataTransfer.files);
  });
  drop.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      fileInput.click();
    }
  });

  const showOffline = () => {
    const title = $("offline-title");
    const body = $("offline-body");
    if (location.protocol === "file:") {
      title.textContent = "시작.bat 을 더블클릭하세요";
      body.innerHTML = "HTML 파일을 직접 열면 첨부파일이 저장되지 않습니다.<br />프로젝트 폴더의 <strong>시작.bat</strong> 을 실행한 뒤, 열린 브라우저에서 사용하세요.";
    } else {
      title.textContent = "서버에 연결하지 못했습니다";
      body.textContent = "배포가 아직 끝나지 않았거나 서버가 꺼져 있습니다. 잠시 후 새로고침해 주세요.";
    }
    offline.classList.remove("is-hidden");
  };

  const startApp = (data) => {
    menus = data.menus;
    applyStats(data);
    applyEvents(data.events);
    if (data.online != null) setOnline(data.online);
    if (data.persist === false && navHint) {
      navHint.textContent = "무료 서버는 재시작하면 파일이 사라질 수 있습니다. GITHUB_TOKEN을 설정하세요.";
    }
    renderMenus();
    const beat = async () => {
      try {
        const res = await fetch("api/presence", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ id: sessionId }),
        });
        const next = await res.json();
        if (res.ok && next.online != null) setOnline(next.online);
      } catch (_) {}
    };
    beat();
    setInterval(() => {
      if (!document.hidden) beat();
    }, 5000);
    setInterval(async () => {
      if (document.hidden || menuList.querySelector(".menu-edit")) return;
      try {
        const poll = await fetch("api/bootstrap", { cache: "no-store" });
        if (!poll.ok) return;
        const next = await poll.json();
        menus = next.menus;
        applyStats(next);
        applyEvents(next.events);
        if (next.online != null) setOnline(next.online);
        renderMenus();
        if (view !== "upload") {
          const menu = menus.find((m) => m.id === view);
          if (menu) {
            filesTitle.textContent = menu.label;
            if (view === "memo") sheetTitle.textContent = menu.label;
          }
        }
      } catch (_) {}
    }, 8000);
  };

  if (loginForm) {
    loginForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      loginErr.classList.add("is-hidden");
      try {
        const res = await fetch("api/login", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ password: loginPass.value }),
        });
        const data = await res.json().catch(() => ({}));
        if (!res.ok) throw new Error(data.error || "로그인 실패");
        login.classList.add("is-hidden");
        const boot = await fetch("api/bootstrap", { cache: "no-store" });
        if (!boot.ok) throw new Error();
        startApp(await boot.json());
      } catch (err) {
        loginErr.textContent = err.message || "비밀번호가 올바르지 않습니다.";
        loginErr.classList.remove("is-hidden");
      }
    });
  }

  (async () => {
    try {
      const authRes = await fetch("api/auth", { cache: "no-store" });
      if (authRes.ok) {
        const auth = await authRes.json();
        if (auth.required && !auth.ok) {
          login.classList.remove("is-hidden");
          loginPass.focus();
          return;
        }
      }
      const res = await fetch("api/bootstrap", { cache: "no-store" });
      if (res.status === 401) {
        login.classList.remove("is-hidden");
        loginPass.focus();
        return;
      }
      if (!res.ok) throw new Error();
      startApp(await res.json());
    } catch (_) {
      showOffline();
    }
  })();
})();
