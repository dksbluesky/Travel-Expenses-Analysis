// 消費分析儀表板 — 100% 前端執行，檔案不會上傳到任何伺服器。
// 讀取 PDF 轉 Excel 工具輸出的 "Raw Sheet"：日期 / 店家/場所 / 品名 / 說明 / 數量 / 單價 / 金額 [/ NT$ / 金額]

const COL = {
  date: "日期",
  store: "店家/場所",
  item: "品名",
  desc: "說明",
  qty: "數量",
  price: "單價",
  amount: "金額",
  ntd: "NT$ / 金額",
};

const DEFAULT_RULES_TEXT = `餐飲: 麥當勞,肯德基,摩斯,漢堡王,星巴克,路易莎,cama,85度C,八方雲集,鼎泰豐,早餐,午餐,晚餐,火鍋,餐廳,小吃,飲料,咖啡,茶,便當,牛肉麵,燒烤,食堂,美食,可頌,麵包,壽司,拉麵,牛排,簡餐,熱炒,KFC,Starbucks,McDonald
購物: 全聯,家樂福,大潤發,愛買,7-11,7-ELEVEN,7-Eleven,全家,FamilyMart,OK超商,萊爾富,美廉社,屈臣氏,康是美,寶雅,百貨,SOGO,新光三越,momo,PChome,蝦皮,賣場,超市,量販,文具,書局,誠品
居家: IKEA,特力屋,特力家,家具,燈飾,五金,生活百貨,寢具
交通: 加油,中油,台灣中油,全國加油站,停車,高鐵,台鐵,客運,計程車,Uber,捷運,ETC,機車,汽車,停車場,加氣
娛樂: 電影,威秀,影城,KTV,遊樂園,健身房,運動中心,展覽,門票,樂園`;

let categoryRules = parseRulesText(DEFAULT_RULES_TEXT);
let allRows = [];      // flat item-level rows across all uploaded files
let loadedFiles = [];  // {name, rowCount}

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

function parseRulesText(text) {
  return text.split("\n").map(l => l.trim()).filter(Boolean).map(line => {
    const idx = line.indexOf(":");
    if (idx === -1) return null;
    const name = line.slice(0, idx).trim();
    const keywords = line.slice(idx + 1).split(",").map(k => k.trim().toLowerCase()).filter(Boolean);
    return { name, keywords };
  }).filter(Boolean);
}

function categorize(store, item) {
  const haystack = `${store || ""} ${item || ""}`.toLowerCase();
  for (const rule of categoryRules) {
    if (rule.keywords.some(kw => kw && haystack.includes(kw))) return rule.name;
  }
  return "其他";
}

function parseAmount(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return isNaN(v) ? 0 : v;
  const cleaned = String(v).replace(/[,\s]/g, "").replace(/[^\d.\-]/g, "");
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  const s = String(v).trim();
  const m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (!m) return null;
  const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  return isNaN(d) ? null : d;
}

function fmtMoney(n) {
  return "$" + Math.round(n).toLocaleString("zh-TW");
}

function findHeaderRow(rows) {
  // Raw Sheet header is normally row 0, but build_excel() shifts it to row 1
  // when a currency rate-note occupies A1 — check both.
  for (let i = 0; i < Math.min(rows.length, 3); i++) {
    const row = rows[i] || [];
    if (row.some(c => String(c).trim() === COL.date)) return i;
  }
  return -1;
}

function extractRowsFromSheet(sheet, fileName) {
  const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const headerIdx = findHeaderRow(grid);
  if (headerIdx === -1) return { rows: [], error: `找不到表頭「${COL.date}」，請確認此檔案含有「Raw Sheet」工作表` };

  const headers = grid[headerIdx].map(h => String(h).trim());
  const colIndex = {};
  Object.entries(COL).forEach(([key, label]) => {
    colIndex[key] = headers.indexOf(label);
  });
  if (colIndex.date === -1 || colIndex.amount === -1) {
    return { rows: [], error: "表頭欄位不完整，無法解析" };
  }

  const out = [];
  for (let i = headerIdx + 1; i < grid.length; i++) {
    const r = grid[i];
    if (!r || r.every(c => c === "" || c === undefined)) continue;
    const dateRaw = r[colIndex.date];
    if (String(dateRaw).startsWith("小計") || String(dateRaw).startsWith("總計")) continue;

    const store = colIndex.store !== -1 ? String(r[colIndex.store] ?? "").trim() : "";
    const item = colIndex.item !== -1 ? String(r[colIndex.item] ?? "").trim() : "";
    if (!store && !item) continue;

    const ntdVal = colIndex.ntd !== -1 ? r[colIndex.ntd] : "";
    const baseAmount = r[colIndex.amount];
    const amount = parseAmount(ntdVal !== "" && ntdVal !== undefined ? ntdVal : baseAmount);
    const dateObj = parseDate(dateRaw);

    out.push({
      dateStr: dateObj ? dateObj.toISOString().slice(0, 10) : (String(dateRaw).trim() || "未知日期"),
      dateObj,
      store: store || "未知店家",
      item,
      desc: colIndex.desc !== -1 ? String(r[colIndex.desc] ?? "").trim() : "",
      qty: colIndex.qty !== -1 ? r[colIndex.qty] : "",
      amount,
      category: categorize(store, item),
      sourceFile: fileName,
    });
  }
  return { rows: out, error: null };
}

async function handleFiles(fileList) {
  const files = Array.from(fileList).filter(f => /\.xlsx?$/i.test(f.name));
  if (!files.length) return;

  for (const file of files) {
    if (loadedFiles.some(f => f.name === file.name)) continue;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: false });
      const sheet = wb.Sheets["Raw Sheet"] || wb.Sheets[wb.SheetNames[0]];
      const { rows, error } = extractRowsFromSheet(sheet, file.name);
      if (error) {
        setStatus(`⚠️ ${file.name}：${error}`, true);
        continue;
      }
      allRows.push(...rows);
      loadedFiles.push({ name: file.name, rowCount: rows.length });
    } catch (e) {
      setStatus(`⚠️ ${file.name} 讀取失敗：${e.message}`, true);
    }
  }
  renderFileList();
  if (allRows.length) {
    $("#dashboard").classList.remove("hidden");
    renderAll();
    setStatus(`已載入 ${loadedFiles.length} 個檔案，共 ${allRows.length} 筆消費記錄`);
  }
}

function removeFile(name) {
  loadedFiles = loadedFiles.filter(f => f.name !== name);
  allRows = allRows.filter(r => r.sourceFile !== name);
  renderFileList();
  if (allRows.length) {
    renderAll();
  } else {
    $("#dashboard").classList.add("hidden");
  }
}

function setStatus(msg, isError) {
  const el = $("#uploadStatus");
  el.textContent = msg;
  el.className = "status" + (isError ? " error" : "");
}

function renderFileList() {
  const ul = $("#fileList");
  ul.innerHTML = "";
  loadedFiles.forEach(f => {
    const li = document.createElement("li");
    li.innerHTML = `<span>📄 ${f.name}（${f.rowCount} 筆）</span>`;
    const btn = document.createElement("button");
    btn.textContent = "移除";
    btn.onclick = () => removeFile(f.name);
    li.appendChild(btn);
    ul.appendChild(li);
  });
}

// ── KPI ────────────────────────────────────────────
function visitKey(r) { return `${r.dateStr}|${r.store}`; }

function renderKPI() {
  const total = allRows.reduce((s, r) => s + r.amount, 0);
  const visits = new Set(allRows.map(visitKey));
  const stores = new Set(allRows.map(r => r.store));
  const dates = allRows.map(r => r.dateObj).filter(Boolean).sort((a, b) => a - b);
  const dayCount = new Set(allRows.map(r => r.dateStr)).size || 1;

  $("#kpiTotal").textContent = fmtMoney(total);
  $("#kpiVisits").textContent = visits.size + " 次";
  $("#kpiAvg").textContent = fmtMoney(visits.size ? total / visits.size : 0);
  $("#kpiStores").textContent = stores.size + " 家";
  $("#kpiDaily").textContent = fmtMoney(total / dayCount);
  $("#kpiRange").textContent = dates.length
    ? `${dates[0].toISOString().slice(0, 10)} ~ ${dates[dates.length - 1].toISOString().slice(0, 10)}`
    : "未知";
}

// ── 店家分析 ────────────────────────────────────────
let storeChart, storeMetric = "amount";

function storeStats() {
  const map = new Map();
  allRows.forEach(r => {
    if (!map.has(r.store)) map.set(r.store, { store: r.store, category: r.category, count: 0, amount: 0, lastDate: r.dateStr, visits: new Set() });
    const s = map.get(r.store);
    s.amount += r.amount;
    s.visits.add(visitKey(r));
    if (r.dateStr > s.lastDate) s.lastDate = r.dateStr;
  });
  map.forEach(s => { s.count = s.visits.size; });
  return Array.from(map.values()).sort((a, b) => b.amount - a.amount);
}

function splitStoreLabel(store) {
  // Bilingual AI-mode names follow a strict "原文 / 中文翻譯" format (see
  // pdf-convert-to-excel-protable's extraction prompt) — split into two
  // tick lines instead of one long string, so a vertical bar chart doesn't
  // need to truncate or rotate the label.
  const idx = store.indexOf(" / ");
  return idx === -1 ? [store] : [store.slice(0, idx), store.slice(idx + 3)];
}

function renderStores() {
  const stats = storeStats();
  const sorted = [...stats].sort((a, b) => storeMetric === "amount" ? b.amount - a.amount : b.count - a.count);
  const top = sorted.slice(0, 10);

  const ctx = $("#storeChart");
  const data = {
    labels: top.map(s => splitStoreLabel(s.store)),
    datasets: [{
      label: storeMetric === "amount" ? "消費金額" : "消費次數",
      data: top.map(s => storeMetric === "amount" ? s.amount : s.count),
      backgroundColor: "#1ca9e6",
      borderRadius: 4,
    }],
  };
  if (storeChart) storeChart.destroy();
  storeChart = new Chart(ctx, {
    type: "bar",
    data,
    options: {
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { maxRotation: 0, minRotation: 0, autoSkip: false } },
        y: { beginAtZero: true },
      },
    },
  });

  const tbody = $("#storeTable tbody");
  tbody.innerHTML = "";
  stats.forEach(s => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${s.store}</td><td>${s.category}</td><td>${s.count}</td><td>${fmtMoney(s.amount)}</td><td>${fmtMoney(s.amount / s.count)}</td><td>${s.lastDate}</td>`;
    tbody.appendChild(tr);
  });
}

// ── 分類分析 ────────────────────────────────────────
let categoryChart;
const CATEGORY_COLORS = { 餐飲: "#f0833f", 購物: "#1ca9e6", 居家: "#1aab6e", 交通: "#5d6b78", 娛樂: "#f0b429", 其他: "#9b9b9b" };

function categoryColor(name) {
  return CATEGORY_COLORS[name] || "#" + ((Array.from(name).reduce((h, c) => h + c.codePointAt(0), 0) * 999) % 0xffffff).toString(16).padStart(6, "0");
}

function categoryStats() {
  const map = new Map();
  allRows.forEach(r => {
    if (!map.has(r.category)) map.set(r.category, { category: r.category, amount: 0, count: 0 });
    const c = map.get(r.category);
    c.amount += r.amount;
    c.count += 1;
  });
  return Array.from(map.values()).sort((a, b) => b.amount - a.amount);
}

function renderCategories() {
  const stats = categoryStats();
  const ctx = $("#categoryChart");
  if (categoryChart) categoryChart.destroy();
  categoryChart = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels: stats.map(s => s.category),
      datasets: [{ data: stats.map(s => s.amount), backgroundColor: stats.map(s => categoryColor(s.category)) }],
    },
    options: { plugins: { legend: { position: "bottom" } } },
  });

  const total = stats.reduce((s, c) => s + c.amount, 0) || 1;
  const list = $("#categoryList");
  list.innerHTML = "";
  stats.forEach(s => {
    const row = document.createElement("div");
    row.className = "category-row";
    row.innerHTML = `<span class="cat-dot" style="background:${categoryColor(s.category)}"></span>
      <span class="cat-name">${s.category}（${s.count} 筆）</span>
      <span class="cat-amount">${fmtMoney(s.amount)}（${(s.amount / total * 100).toFixed(1)}%）</span>`;
    list.appendChild(row);
  });
}

// ── 時間分析 ────────────────────────────────────────
let monthChart, weekdayChart;
const WEEKDAY_LABELS = ["週日", "週一", "週二", "週三", "週四", "週五", "週六"];

function renderTime() {
  const monthMap = new Map();
  const weekdayTotals = new Array(7).fill(0);

  allRows.forEach(r => {
    if (r.dateObj) {
      const ym = r.dateStr.slice(0, 7);
      monthMap.set(ym, (monthMap.get(ym) || 0) + r.amount);
      weekdayTotals[r.dateObj.getUTCDay()] += r.amount;
    }
  });

  const months = Array.from(monthMap.keys()).sort();
  if (monthChart) monthChart.destroy();
  monthChart = new Chart($("#monthChart"), {
    type: "bar",
    data: { labels: months, datasets: [{ label: "月消費金額", data: months.map(m => monthMap.get(m)), backgroundColor: "#1ca9e6", borderRadius: 4 }] },
    options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } },
  });

  if (weekdayChart) weekdayChart.destroy();
  weekdayChart = new Chart($("#weekdayChart"), {
    type: "bar",
    data: { labels: WEEKDAY_LABELS, datasets: [{ label: "星期消費金額", data: weekdayTotals, backgroundColor: "#f0833f", borderRadius: 4 }] },
    options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } },
  });
}

// ── 明細列表 ────────────────────────────────────────
function renderDetail(filterText) {
  const tbody = $("#detailTable tbody");
  tbody.innerHTML = "";
  const filter = (filterText || "").trim().toLowerCase();
  const rows = [...allRows].sort((a, b) => b.dateStr.localeCompare(a.dateStr));
  rows.forEach(r => {
    if (filter && !(`${r.store} ${r.item}`.toLowerCase().includes(filter))) return;
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${r.dateStr}</td><td>${r.store}</td><td>${r.item}</td><td>${r.category}</td><td>${r.qty}</td><td>${fmtMoney(r.amount)}</td>`;
    tbody.appendChild(tr);
  });
}

function renderAll() {
  allRows.forEach(r => { r.category = categorize(r.store, r.item); });
  renderKPI();
  renderStores();
  renderCategories();
  renderTime();
  renderDetail($("#detailSearch").value);
}

// ── UI wiring ───────────────────────────────────────
function init() {
  $("#categoryRules").value = DEFAULT_RULES_TEXT;

  const dropZone = $("#dropZone");
  const fileInput = $("#fileInput");
  fileInput.addEventListener("change", e => handleFiles(e.target.files));
  ["dragenter", "dragover"].forEach(evt =>
    dropZone.addEventListener(evt, e => { e.preventDefault(); dropZone.classList.add("dragover"); }));
  ["dragleave", "drop"].forEach(evt =>
    dropZone.addEventListener(evt, e => { e.preventDefault(); dropZone.classList.remove("dragover"); }));
  dropZone.addEventListener("drop", e => handleFiles(e.dataTransfer.files));

  $$(".tab-btn").forEach(btn => btn.addEventListener("click", () => {
    $$(".tab-btn").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    $$(".tab-panel").forEach(p => p.classList.add("hidden"));
    $(`#tab-${btn.dataset.tab}`).classList.remove("hidden");
  }));

  $$(".toggle-btn").forEach(btn => btn.addEventListener("click", () => {
    $$(".toggle-btn").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    storeMetric = btn.dataset.metric;
    renderStores();
  }));

  $("#detailSearch").addEventListener("input", e => renderDetail(e.target.value));

  $("#saveRulesBtn").addEventListener("click", () => {
    categoryRules = parseRulesText($("#categoryRules").value);
    if (allRows.length) renderAll();
    const status = $("#rulesStatus");
    status.textContent = "已套用";
    setTimeout(() => { status.textContent = ""; }, 1500);
  });
}

document.addEventListener("DOMContentLoaded", init);
