/**
 * HR 工具：名單解析、抽籤動畫、分組視覺化
 */

/** @param {string} text */
function parseNamesRaw(text) {
  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const names = [];
  for (const line of lines) {
    const firstCell = line.split(/,|\t|;/)[0].trim();
    if (firstCell) names.push(firstCell);
  }
  return names;
}

/** @param {string[]} arr */
function dedupePreserveOrder(arr) {
  const seen = new Set();
  const out = [];
  for (const x of arr) {
    if (!seen.has(x)) {
      seen.add(x);
      out.push(x);
    }
  }
  return out;
}

/**
 * 解析單列 CSV（支援雙引號欄位內含逗號）。
 * @param {string} line
 * @returns {string[]}
 */
function parseCsvLine(line) {
  const result = [];
  let i = 0;
  let cur = "";
  let inQuotes = false;
  const len = line.length;
  while (i < len) {
    const c = line[i];
    if (inQuotes) {
      if (c === '"') {
        if (i + 1 < len && line[i + 1] === '"') {
          cur += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      cur += c;
      i++;
    } else {
      if (c === '"') {
        inQuotes = true;
        i++;
        continue;
      }
      if (c === ",") {
        result.push(cur.trim());
        cur = "";
        i++;
        continue;
      }
      cur += c;
      i++;
    }
  }
  result.push(cur.trim());
  return result;
}

/** @param {string} s */
function normalizeHeaderToken(s) {
  return String(s)
    .replace(/^\uFEFF/, "")
    .replace(/\u3000/g, " ")
    .trim()
    .replace(/\s+/g, "");
}

/**
 * @param {string[]} headers
 * @param {string[]} keys 欲比對的標題字（已去空白）
 */
function findColumnIndex(headers, keys) {
  const cells = headers.map((h) => normalizeHeaderToken(h));
  for (const key of keys) {
    const k = normalizeHeaderToken(key);
    const idx = cells.findIndex((c) => c === k);
    if (idx >= 0) return idx;
  }
  return -1;
}

/**
 * 從 CSV 匯入姓名：優先尋找標題列中的「姓名」欄；若有「組別」「次序」則一併讀取並排序。
 * 若找不到「姓名」標題，改為每列第一欄（與舊版相容）。
 * @param {string} csvText
 * @returns {{ names: string[], detail: string }}
 */
function importNamesFromCsv(csvText) {
  const raw = String(csvText).replace(/^\uFEFF/, "");
  const lines = raw.split(/\r?\n/);
  /** @type {string[][]} */
  const rows = [];
  for (const line of lines) {
    if (line.trim() === "") continue;
    rows.push(parseCsvLine(line));
  }

  if (rows.length === 0) {
    return { names: [], detail: "檔案沒有可讀取的列。" };
  }

  const nameKeys = ["姓名", "名字"];
  const groupKeys = ["組別"];
  const orderKeys = ["次序"];

  let headerRow = -1;
  let nameCol = -1;
  let groupCol = -1;
  let orderCol = -1;

  for (let r = 0; r < rows.length; r++) {
    const nc = findColumnIndex(rows[r], nameKeys);
    if (nc >= 0) {
      headerRow = r;
      nameCol = nc;
      groupCol = findColumnIndex(rows[r], groupKeys);
      orderCol = findColumnIndex(rows[r], orderKeys);
      break;
    }
  }

  if (headerRow < 0) {
    const names = [];
    for (const row of rows) {
      if (row.length === 0) continue;
      const cell = String(row[0] ?? "")
        .trim()
        .replace(/^["']|["']$/g, "");
      if (cell) names.push(cell);
    }
    const detail =
      names.length > 0
        ? "未偵測到「姓名」標題列，已依每列第一欄匯入（與舊版相容）。"
        : "無法從 CSV 讀取姓名。";
    return { names, detail };
  }

  /** @type {{ name: string; g: number; o: number; _idx: number }[]} */
  const records = [];
  for (let i = headerRow + 1; i < rows.length; i++) {
    const row = rows[i];
    if (nameCol >= row.length) continue;
    const name = String(row[nameCol] ?? "")
      .trim()
      .replace(/^["']|["']$/g, "");
    if (!name) continue;

    let g = 0;
    let o = 0;
    if (groupCol >= 0 && groupCol < row.length) {
      const gn = parseInt(String(row[groupCol]).trim(), 10);
      g = Number.isFinite(gn) ? gn : 0;
    }
    if (orderCol >= 0 && orderCol < row.length) {
      const on = parseInt(String(row[orderCol]).trim(), 10);
      o = Number.isFinite(on) ? on : 0;
    }
    records.push({ name, g, o, _idx: records.length });
  }

  const hasGroup = groupCol >= 0;
  const hasOrder = orderCol >= 0;
  if (hasGroup && hasOrder) {
    records.sort((a, b) => a.g - b.g || a.o - b.o || a._idx - b._idx);
  } else if (hasGroup) {
    records.sort((a, b) => a.g - b.g || a._idx - b._idx);
  } else if (hasOrder) {
    records.sort((a, b) => a.o - b.o || a._idx - b._idx);
  }

  const names = records.map((r) => r.name);

  let detail = `已依「姓名」欄匯入（偵測到標題列）。`;
  if (hasGroup && hasOrder) {
    detail += " 已依「組別」「次序」排序。";
  } else if (hasGroup) {
    detail += " 已依「組別」排序。";
  } else if (hasOrder) {
    detail += " 已依「次序」排序。";
  } else {
    detail += " 未含「組別／次序」欄，維持原列順序。";
  }

  return { names, detail };
}

/** @param {string[]} names */
function duplicateFlags(names) {
  const counts = new Map();
  for (const n of names) counts.set(n, (counts.get(n) || 0) + 1);
  return names.map((n) => counts.get(n) > 1);
}

function shuffleArray(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function chunk(array, size) {
  if (size < 1) return [];
  const chunks = [];
  for (let i = 0; i < array.length; i += size) {
    chunks.push(array.slice(i, i + size));
  }
  return chunks;
}

/** @param {string} s */
function csvEscapeCell(s) {
  if (/[",\r\n]/.test(s)) return `"${String(s).replace(/"/g, '""')}"`;
  return String(s);
}

/** @param {string[][]} groups */
function groupsToCsv(groups) {
  const lines = ["組別,次序,姓名"];
  groups.forEach((g, gi) => {
    g.forEach((name, idx) => {
      lines.push(`${gi + 1},${idx + 1},${csvEscapeCell(name)}`);
    });
  });
  return lines.join("\r\n");
}

function downloadTextFile(filename, text, mime) {
  const blob = new Blob([text], { type: mime || "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.rel = "noopener";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function todayStamp() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

const DEMO_SAMPLE = `王小明
李小華
張三
陳美玲
林志豪
王小明
黃雅婷
林怡君`;

// —— State ——
let namesMaster = [];
/** 不重複模式下的剩餘池 */
let poolUnique = [];

/** @type {string[]} */
let historyWinners = [];

/** @type {string[][] | null} */
let lastGroups = null;

// —— DOM ——
const $textarea = document.getElementById("name-textarea");
const $csvFile = document.getElementById("csv-file");
const $clearList = document.getElementById("clear-list");
const $btnDemo = document.getElementById("btn-demo");
const $nameCount = document.getElementById("name-count");
const $nameCountExtra = document.getElementById("name-count-extra");
const $parseNote = document.getElementById("parse-note");
const $namePreviewWrap = document.getElementById("name-preview-wrap");
const $namePreview = document.getElementById("name-preview");
const $btnDedupe = document.getElementById("btn-dedupe");

const $tabBtns = document.querySelectorAll(".tabs__btn");
const $tabLottery = document.getElementById("tab-lottery");
const $tabGroups = document.getElementById("tab-groups");

const $drawCount = document.getElementById("draw-count");
const $btnDraw = document.getElementById("btn-draw");
const $btnResetPool = document.getElementById("btn-reset-pool");
const $lotteryDisplay = document.getElementById("lottery-display");
const $lotteryMeta = document.getElementById("lottery-meta");
const $winnersList = document.getElementById("winners-list");

const $groupSize = document.getElementById("group-size");
const $shuffleGroups = document.getElementById("shuffle-groups");
const $btnGroup = document.getElementById("btn-group");
const $btnDownloadGroups = document.getElementById("btn-download-groups");
const $groupsVisual = document.getElementById("groups-visual");

function getRawNamesFromText() {
  return parseNamesRaw($textarea.value);
}

function getNamesFromUI() {
  return dedupePreserveOrder(getRawNamesFromText());
}

function invalidateGroupsExport() {
  lastGroups = null;
  $btnDownloadGroups.disabled = true;
}

function syncFromTextarea() {
  namesMaster = getNamesFromUI();
  poolUnique = [...namesMaster];
  historyWinners = [];
  renderNameCount();
  renderNamePreview();
  renderWinners();
  clearLotteryDisplayIdle();
  invalidateGroupsExport();
}

function renderNameCount() {
  const raw = getRawNamesFromText();
  const unique = namesMaster.length;
  const flags = duplicateFlags(raw);

  $nameCount.textContent = String(unique);

  if (raw.length === 0) {
    $nameCountExtra.hidden = true;
    $nameCountExtra.textContent = "";
  } else if (raw.length !== unique) {
    const dupRows = flags.filter(Boolean).length;
    $nameCountExtra.hidden = false;
    $nameCountExtra.textContent = `（原始 ${raw.length} 筆，其中 ${dupRows} 筆為重複姓名）`;
  } else {
    $nameCountExtra.hidden = true;
    $nameCountExtra.textContent = "";
  }
}

function renderNamePreview() {
  const raw = getRawNamesFromText();
  const flags = duplicateFlags(raw);

  if (raw.length === 0) {
    $namePreviewWrap.hidden = true;
    $namePreview.innerHTML = "";
    $btnDedupe.hidden = true;
    return;
  }

  $namePreviewWrap.hidden = false;
  const hasDup = flags.some(Boolean);
  $btnDedupe.hidden = !hasDup;

  $namePreview.innerHTML = raw
    .map((name, i) => {
      const dup = flags[i];
      const cls = dup ? "name-preview__row name-preview__row--dup" : "name-preview__row";
      const badge = dup
        ? '<span class="name-preview__badge" title="此姓名在名單中出現多次">重複</span>'
        : "";
      return `<li class="${cls}"><span class="name-preview__idx">${i + 1}</span><span class="name-preview__name">${escapeHtml(
        name
      )}</span>${badge}</li>`;
    })
    .join("");
}

function syncPoolFromMaster() {
  poolUnique = [...namesMaster];
}

function getDrawMode() {
  const el = document.querySelector('input[name="draw-mode"]:checked');
  return el && el.value === "repeat" ? "repeat" : "unique";
}

function clearLotteryDisplayIdle() {
  if (namesMaster.length === 0) {
    $lotteryDisplay.innerHTML = '<span class="lottery-display__idle">請先貼上或上傳名單</span>';
  } else {
    $lotteryDisplay.innerHTML = '<span class="lottery-display__idle">準備就緒</span>';
  }
  $lotteryDisplay.classList.remove("is-spinning");
}

function renderWinners() {
  if (historyWinners.length === 0) {
    $winnersList.innerHTML = "";
    return;
  }
  const tags = historyWinners
    .map((n, i) => `<span class="tag tag--new" style="animation-delay:${Math.min(i * 0.03, 0.5)}s">${escapeHtml(n)}</span>`)
    .join("");
  $winnersList.innerHTML = `<p class="winners__title">本次抽中</p><div class="winners__tags">${tags}</div>`;
}

function escapeHtml(s) {
  const div = document.createElement("div");
  div.textContent = s;
  return div.innerHTML;
}

let spinToken = 0;

/**
 * @param {string[]} poolForDisplay
 * @param {number} durationMs
 * @param {() => void} onDone
 */
function runSpinAnimation(poolForDisplay, durationMs, onDone) {
  const myToken = ++spinToken;
  $lotteryDisplay.classList.add("is-spinning");
  const names = poolForDisplay.length ? poolForDisplay : namesMaster;
  const start = Date.now();
  const tick = () => {
    if (myToken !== spinToken) return;
    const elapsed = Date.now() - start;
    const pick = names[Math.floor(Math.random() * names.length)];
    $lotteryDisplay.innerHTML = `<div class="lottery-display__name">${escapeHtml(pick)}</div>`;
    if (elapsed >= durationMs) {
      onDone();
      return;
    }
    const left = durationMs - elapsed;
    const delay = Math.max(30, Math.min(100, left / 20));
    setTimeout(tick, delay);
  };
  tick();
}

/**
 * @param {string[]} pool
 * @param {number} n
 */
function drawSample(pool, n) {
  if (pool.length === 0 || n <= 0) return [];
  const take = Math.min(n, pool.length);
  const shuffled = shuffleArray(pool);
  return shuffled.slice(0, take);
}

$textarea.addEventListener("input", () => {
  syncFromTextarea();
});

$clearList.addEventListener("click", () => {
  $textarea.value = "";
  syncFromTextarea();
  $parseNote.textContent = "";
});

$btnDemo.addEventListener("click", () => {
  $textarea.value = DEMO_SAMPLE;
  syncFromTextarea();
  $parseNote.textContent = "已載入模擬名單（含一筆重複姓名，可試用「預覽名單」與移除重複）。";
});

$btnDedupe.addEventListener("click", () => {
  const unique = dedupePreserveOrder(getRawNamesFromText());
  $textarea.value = unique.join("\n");
  syncFromTextarea();
  $parseNote.textContent = "已移除重複姓名，僅保留各姓名首次出現的列。";
});

$csvFile.addEventListener("change", async (e) => {
  const file = e.target.files && e.target.files[0];
  if (!file) return;
  try {
    const text = await file.text();
    const { names: raw, detail } = importNamesFromCsv(text);
    $textarea.value = raw.join("\n");
    syncFromTextarea();
    const uniq = dedupePreserveOrder(raw);
    const dupHint = raw.length !== uniq.length ? ` 偵測到重複（原始 ${raw.length} 筆，不重複 ${uniq.length} 人）。` : "";
    $parseNote.textContent = `${detail} 共 ${raw.length} 筆。${dupHint}`;
  } catch {
    $parseNote.textContent = "無法讀取檔案。";
  }
  e.target.value = "";
});

$tabBtns.forEach((btn) => {
  btn.addEventListener("click", () => {
    const tab = btn.getAttribute("data-tab");
    $tabBtns.forEach((b) => {
      b.classList.toggle("is-active", b === btn);
      b.setAttribute("aria-selected", b === btn ? "true" : "false");
    });
    if (tab === "lottery") {
      $tabLottery.hidden = false;
      $tabLottery.classList.add("is-visible");
      $tabGroups.hidden = true;
    } else {
      $tabLottery.hidden = true;
      $tabLottery.classList.remove("is-visible");
      $tabGroups.hidden = false;
    }
  });
});

$btnResetPool.addEventListener("click", () => {
  syncPoolFromMaster();
  historyWinners = [];
  renderWinners();
  clearLotteryDisplayIdle();
  $lotteryMeta.textContent = "已重置獎池，所有人可重新被抽中。";
});

$btnDraw.addEventListener("click", () => {
  const mode = getDrawMode();
  const n = Math.max(1, parseInt(String($drawCount.value), 10) || 1);

  if (namesMaster.length === 0) {
    $lotteryMeta.textContent = "請先建立名單。";
    return;
  }

  if (mode === "unique") {
    if (poolUnique.length === 0) {
      $lotteryMeta.textContent = "已無人可抽，請按「重置獎池」或改為可重複模式。";
      return;
    }
    if (n > poolUnique.length) {
      $lotteryMeta.textContent = `剩餘僅 ${poolUnique.length} 人，已調整為抽出 ${poolUnique.length} 人。`;
    }
  }

  const poolForSpin = mode === "unique" ? poolUnique : namesMaster;
  const actualN = mode === "unique" ? Math.min(n, poolUnique.length) : n;

  $btnDraw.disabled = true;
  $lotteryMeta.textContent = "抽籤中…";

  const spinPool = poolForSpin.length ? poolForSpin : namesMaster;
  runSpinAnimation(spinPool, 2200, () => {
    let picked;
    if (mode === "unique") {
      picked = drawSample(poolUnique, actualN);
      for (const p of picked) {
        poolUnique = poolUnique.filter((x) => x !== p);
      }
    } else {
      picked = [];
      for (let i = 0; i < actualN; i++) {
        picked.push(namesMaster[Math.floor(Math.random() * namesMaster.length)]);
      }
    }

    historyWinners = picked;
    if (mode === "unique") {
      $lotteryDisplay.classList.remove("is-spinning");
      if (picked.length === 1) {
        $lotteryDisplay.innerHTML = `<div class="lottery-display__name">${escapeHtml(picked[0])}</div>`;
      } else {
        $lotteryDisplay.innerHTML = `<div class="lottery-display__names">${picked.map((x) => escapeHtml(x)).join(" · ")}</div>`;
      }
      $lotteryMeta.textContent = `剩餘可抽：${poolUnique.length} 人`;
    } else {
      $lotteryDisplay.classList.remove("is-spinning");
      if (picked.length === 1) {
        $lotteryDisplay.innerHTML = `<div class="lottery-display__name">${escapeHtml(picked[0])}</div>`;
      } else {
        $lotteryDisplay.innerHTML = `<div class="lottery-display__names">${picked.map((x) => escapeHtml(x)).join(" · ")}</div>`;
      }
      $lotteryMeta.textContent = "可重複模式：總人數不變";
    }
    renderWinners();
    $btnDraw.disabled = false;
  });
});

$btnGroup.addEventListener("click", () => {
  const size = Math.max(1, parseInt(String($groupSize.value), 10) || 4);
  if (namesMaster.length === 0) {
    $groupsVisual.innerHTML = '<p class="group-card__empty">請先建立名單。</p>';
    lastGroups = null;
    $btnDownloadGroups.disabled = true;
    return;
  }

  let list = [...namesMaster];
  if ($shuffleGroups.checked) list = shuffleArray(list);

  const groups = chunk(list, size);
  lastGroups = groups;
  $btnDownloadGroups.disabled = false;

  $groupsVisual.innerHTML = groups
    .map((g, i) => {
      const items = g.map((name) => `<li>${escapeHtml(name)}</li>`).join("");
      return `
        <article class="group-card">
          <div class="group-card__head">
            <h3 class="group-card__label">第 ${i + 1} 組</h3>
            <span class="group-card__count">${g.length} 人</span>
          </div>
          <ul class="group-card__list">${items}</ul>
        </article>
      `;
    })
    .join("");
});

$btnDownloadGroups.addEventListener("click", () => {
  if (!lastGroups || lastGroups.length === 0) return;
  const csv = groupsToCsv(lastGroups);
  const bom = "\uFEFF";
  downloadTextFile(`分組結果_${todayStamp()}.csv`, bom + csv, "text/csv;charset=utf-8");
});

// 初始
syncFromTextarea();
