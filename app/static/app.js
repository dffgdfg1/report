"use strict";
// 飞书会话过期拦截：任何 API 返回 401 login_required 时跳回飞书登录
(function () {
  const _fetch = window.fetch.bind(window);
  window.fetch = async function (...args) {
    const resp = await _fetch(...args);
    try {
      const url = typeof args[0] === "string" ? args[0] : (args[0] && args[0].url) || "";
      if (resp.status === 401 && url.indexOf("/api/") >= 0) {
        const j = await resp.clone().json().catch(() => ({}));
        if (j && j.login_required) { window.location = "/feishu/login"; }
      }
    } catch (e) {}
    return resp;
  };
})();
// ============ 状态 ============
let META = { types: [], presets: {}, img_w: 4.8, img_h: 6.4 };
let ENV = { feishu: false, user: "", is_local: true };  // 运行环境：飞书免登/是否本机
let state = { name: "", info: {}, tests: [] };
let SCHEMES = [];   // 已保存的测试方案名列表
let STD = {};       // 标准库：{测试项目: {车厂: {standard, condition, requirement}}}
let DEV = [];       // 设备库：[{name, model, mgmt_no, cal_date, cal_end, ...}]
let saveTimer = null;
// 记录最近一次点击/悬停的图片上传区，粘贴截图(Ctrl+V)时图片进这里
let PASTE_TARGET = null;

// ===== 撤销/重做（Ctrl+Z / Ctrl+Y）=====
let undoStack = [];        // 历史状态快照(JSON字符串)
let redoStack = [];        // 重做栈
let lastCommitted = null;  // 最近一次已入历史的状态(JSON)，用于判断是否有新改动
let historyTimer = null;   // 编辑停顿后提交历史的防抖计时器
const UNDO_LIMIT = 60;     // 最多保留多少步

// 与后端 safe_name 完全一致：/ 替换为全角斜杠 ／，其他非法字符 → _。
// 用于构造 /api/image 的路径，避免项目名里的斜杠(如 ME/WTD、YJ/SYBG)导致取图 404。
function safeName(s) {
  s = (s || "").trim().replace(/\//g, "／");
  s = s.replace(/[^0-9A-Za-z_\-一-鿿／]/g, "_");
  return s || "未命名";
}
function imageUrl(file) {
  return `/api/image/${encodeURIComponent(safeName(state.name))}/${encodeURIComponent(file)}`;
}
function escAttr(s) {
  return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

async function reloadSchemes() {
  try { SCHEMES = await readJSON(await fetch("/api/schemes")); } catch (e) { SCHEMES = []; }
}

async function reloadStandards() {
  try { STD = await readJSON(await fetch("/api/standards")); } catch (e) { STD = {}; }
}

async function reloadDevices() {
  try { DEV = await readJSON(await fetch("/api/devices")); } catch (e) { DEV = []; }
}

// 模糊匹配：子序列匹配，"振测"命中"振动测试"，"振动"也命中。忽略大小写。
function fuzzyMatch(name, filter) {
  const s = (name || "").toLowerCase();
  const f = (filter || "").toLowerCase().replace(/\s+/g, "");
  if (!f) return true;
  if (s.includes(f)) return true;        // 连续子串优先
  let i = 0;
  for (const ch of s) { if (ch === f[i]) i++; if (i === f.length) return true; }
  return false;                          // 顺序子序列命中
}

const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const el = (tag, cls, html) => { const e = document.createElement(tag); if (cls) e.className = cls; if (html != null) e.innerHTML = html; return e; };
function toast(msg) { const t = $("#toast"); t.textContent = msg; t.classList.add("show"); clearTimeout(t._t); t._t = setTimeout(() => t.classList.remove("show"), 1800); }
function status(msg) { $("#status").textContent = msg; }

// 安全解析响应：非 JSON（如后端错误页）时给出可读提示，而不是崩溃
async function readJSON(r) {
  const text = await r.text();
  try { return JSON.parse(text); }
  catch (e) {
    throw new Error(`服务器返回异常（HTTP ${r.status}）。可能服务器需要重启：请关闭那个黑色命令行窗口，重新双击「启动.bat」。`);
  }
}

// 常用默认值（新建/空字段时自动填入，可随时手动修改）
const DEFAULT_INFO = {
  // report_no 由后端按“YJ/SYBG-今天日期+序号”算，见 fillNextReportNo()
  client_name: "深圳佑驾创新科技有限公司",
  client_addr: "深圳市福田区沙头街道上沙社区滨河大道9285号中洲滨海商业中心二期1栋A座二十五层",
  maker_name: "深圳佑驾创新科技有限公司",
  maker_addr: "深圳市福田区沙头街道上沙社区滨河大道9285号中洲滨海商业中心二期1栋A座二十五层",
  sample_way: "客户送样",
  lab_name: "深圳佑驾创新科技股份有限公司实验中心",
  test_items: "参考客户要求",
  test_basis: "参考客户要求",
};

// 报告编号为空时，向后端要“今天的下一个编号”并填入
async function fillNextReportNo() {
  if (state.info.report_no) return;  // 已有则不覆盖
  try {
    const j = await readJSON(await fetch("/api/next_report_no"));
    if (j.report_no && !state.info.report_no) {
      state.info.report_no = j.report_no;
      renderInfo();
    }
  } catch (e) { /* 拿不到就留空，用户可手填 */ }
}

// 用默认值补齐 info 中为空的字段（不覆盖已填内容）
function applyDefaults(info) {
  info = info || {};
  for (const k in DEFAULT_INFO) {
    if (info[k] == null || info[k] === "") info[k] = DEFAULT_INFO[k];
  }
  return info;
}

// 报告信息字段定义 [key, 标签, 类型]
const INFO_FIELDS = [
  ["report_no", "报告编号", "text"],
  ["sample_name", "样品名称", "text"],
  ["sample_no", "样品零件号", "text"],
  ["sample_model", "样品型号", "text"],
  ["verify_phase", "验证阶段", "text"],
  ["client_name", "委托方名称", "text"],
  ["client_addr", "委托方地址", "text"],
  ["maker_name", "制造商名称", "text"],
  ["maker_addr", "制造商地址", "text"],
  ["sample_qty", "样品数量", "text"],
  ["rated_volt", "额定电压", "text"],
  ["sample_way", "来样方式", "text"],
  ["recv_date", "收样日期", "date"],
  ["commission_no", "委托单号", "text"],
  ["test_date_range", "检测日期", "daterange"],
  ["lab_name", "检测单位", "text"],
  ["test_items", "检测项目", "text"],
  ["test_basis", "检测依据", "text"],
  ["remark", "备注", "text"],
];

// ============ 渲染：报告信息 ============
function renderInfo() {
  const body = $("#infoBody");
  body.innerHTML = "";
  const grid = el("div", "grid");
  INFO_FIELDS.forEach(([key, label, type]) => {
    if (type === "date" || type === "daterange") { grid.appendChild(fieldInput(state.info, key, label, type)); return; }
    const f = el("div", "field");
    f.innerHTML = `<label>${label}</label>`;
    const inp = el("input");
    inp.value = state.info[key] || "";
    inp.oninput = () => {
      state.info[key] = inp.value;
      // 委托单号 / 样品型号 变化时，自动重算项目名称
      if (key === "commission_no" || key === "sample_model") autoName();
      scheduleSave();
    };
    f.appendChild(inp);
    grid.appendChild(f);
  });
  body.appendChild(grid);
}

// 项目名称 = 委托单号 + 样品型号 + “试验报告”（自动生成，只读）
function autoName() {
  const info = state.info || {};
  const c = (info.commission_no || "").trim();
  const s = (info.sample_model || "").trim();
  // 委托单号与样品型号之间用“-”连接；某一项为空时不留多余的横杠
  const head = [c, s].filter(Boolean).join("-");
  state.name = head ? `${head}试验报告` : "";
  const pn = $("#pname");
  if (pn) pn.value = state.name;
}
// ============ 渲染：测试项目列表 ============
function renderTests() {
  const list = $("#testList");
  list.innerHTML = "";
  state.tests.forEach((t, i) => list.appendChild(renderTest(t, i)));
}

function fieldInput(obj, key, label, type) {
  if (type === "date") return dateField(obj, key, label);
  if (type === "daterange") return dateRangeField(obj, key, label);
  const f = el("div", "field");
  f.innerHTML = `<label>${label}</label>`;
  const inp = el(type === "textarea" ? "textarea" : "input");
  inp.value = obj[key] || "";
  inp.oninput = () => { obj[key] = inp.value; scheduleSave(); };
  f.appendChild(inp);
  return f;
}

// 解析样机编号：支持 "7#-9#"、"7-9"、"1#~3#"、单个 "8#"/"8"。
// 返回编号数组如 ["7#","8#","9#"]；无法识别返回 null。
function parseSampleNos(str) {
  const s = (str || "").trim();
  if (!s) return [];
  const range = s.match(/^(\d+)#?\s*[-~到至]\s*(\d+)#?$/);
  if (range) {
    let a = parseInt(range[1], 10), b = parseInt(range[2], 10);
    if (a > b) { const tmp = a; a = b; b = tmp; }
    const out = [];
    for (let i = a; i <= b; i++) out.push(i + "#");
    return out;
  }
  const single = s.match(/^(\d+)#?$/);
  if (single) return [parseInt(single[1], 10) + "#"];
  return null;  // 格式不认识（如含字母），不自动改样品
}

// 按样机编号同步"试验结论"样品行：编号按范围填好，结果保留原有(按行)或继承首行
function syncSamplesFromSampleNo(t) {
  const nos = parseSampleNos(t.sample_no);
  if (nos === null || nos.length === 0) return false;  // 无法解析则不动样品
  if (!t.samples) t.samples = [];
  const base = t.samples.length > 0 ? t.samples[0].result : DEFAULT_SAMPLE_RESULT;
  t.samples = nos.map((no, i) => {
    const ex = t.samples[i];
    return {
      no,
      result: ex ? ex.result : base,
      conclusion: ex ? ex.conclusion : "合格",
    };
  });
  return true;
}

// 样机编号字段：输入时只存值(不重渲染,保持焦点)，失焦时按范围同步样品编号
function sampleNoField(t, idx) {
  const f = el("div", "field");
  f.innerHTML = `<label>样机编号（如 1#-2#）</label>`;
  const inp = el("input");
  inp.value = t.sample_no || "";
  inp.oninput = () => { t.sample_no = inp.value; scheduleSave(); };
  inp.onblur = () => {
    if (syncSamplesFromSampleNo(t)) { renderTest_replace(idx); scheduleSave(); }
  };
  f.appendChild(inp);
  return f;
}

// ===== 日期工具：存储用 2026.07.28；<input type=date> 用 2026-07-28 =====
function toISO(s) {  // 任意 年?月?日 -> YYYY-MM-DD（供 date 控件）
  const m = (s || "").match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  if (!m) return "";
  return `${m[1]}-${String(m[2]).padStart(2, "0")}-${String(m[3]).padStart(2, "0")}`;
}
function toDot(iso) {  // YYYY-MM-DD -> YYYY.MM.DD（存储/报告用）
  const m = (iso || "").match(/(\d{4})-(\d{2})-(\d{2})/);
  return m ? `${m[1]}.${m[2]}.${m[3]}` : "";
}
function splitRange(s) {  // "2025.11.17-2026.11.16" 或 "…至…" -> [起, 止]
  const parts = (s || "").split(/\s*(?:至|~|—|-)\s*/).filter(Boolean);
  // 因日期本身含 - 会被误拆，改用正则提取两段完整日期
  const dates = (s || "").match(/\d{4}\D+\d{1,2}\D+\d{1,2}/g) || [];
  return [dates[0] || "", dates[1] || ""];
}

// 单个日期选择器（点击选择，存 YYYY.MM.DD）
function dateField(obj, key, label) {
  const f = el("div", "field");
  f.innerHTML = `<label>${label}</label>`;
  const inp = el("input"); inp.type = "date";
  inp.value = toISO(obj[key]);
  inp.oninput = () => { obj[key] = toDot(inp.value); scheduleSave(); };
  f.appendChild(inp);
  return f;
}

// 起止区间：两个日期选择器，存 "起-止"（保持原有惯例）
function dateRangeField(obj, key, label, sep) {
  sep = sep || "-";
  const f = el("div", "field");
  f.innerHTML = `<label>${label}</label>`;
  const row = el("div", "date-range");
  const [a0, b0] = splitRange(obj[key]);
  const a = el("input"); a.type = "date"; a.value = toISO(a0);
  const b = el("input"); b.type = "date"; b.value = toISO(b0);
  const sync = () => {
    const s = toDot(a.value), e = toDot(b.value);
    obj[key] = (s || e) ? `${s}${sep}${e}` : "";
    scheduleSave();
  };
  a.oninput = sync; b.oninput = sync;
  row.appendChild(a); row.appendChild(el("span", "date-sep", "至")); row.appendChild(b);
  f.appendChild(row);
  return f;
}

// 开始时间：点选日期；同步写入 test_date（测试日期=开始时间）
function startDateField(t, idx) {
  const f = el("div", "field");
  f.innerHTML = `<label>开始时间（＝测试日期）</label>`;
  const inp = el("input"); inp.type = "date";
  inp.value = toISO(t.start_date);
  if (t.end_date) inp.max = toISO(t.end_date);  // 开始不得晚于完成
  inp.oninput = () => {
    t.start_date = toDot(inp.value);
    t.test_date = t.start_date;  // 自动生成测试日期
    // 若完成时间早于新开始时间，清掉完成时间避免出现倒序
    if (t.end_date && toISO(t.end_date) < inp.value) { t.end_date = ""; toast("完成时间早于开始时间，已清空，请重选"); }
    renderTest_replace(idx);  // 刷新以更新完成时间的可选下限
  };
  f.appendChild(inp);
  return f;
}

// 完成时间：下限锁定为开始时间，选更早的日期直接被日历禁用
function endDateField(t, idx) {
  const f = el("div", "field");
  f.innerHTML = `<label>完成时间</label>`;
  const inp = el("input"); inp.type = "date";
  inp.value = toISO(t.end_date);
  if (t.start_date) inp.min = toISO(t.start_date);  // 不得早于开始
  inp.oninput = () => {
    // 双保险：手动改值若早于开始时间则拒绝
    if (t.start_date && inp.value && inp.value < toISO(t.start_date)) {
      toast("完成时间不能早于开始时间"); inp.value = toISO(t.end_date); return;
    }
    t.end_date = toDot(inp.value); scheduleSave();
  };
  f.appendChild(inp);
  return f;
}

// 标准号/条款号：输入框 + “参考客户大纲”快捷按钮
function standardField(t) {
  const f = el("div", "field");
  f.innerHTML = `<label>标准号/条款号</label>`;
  const row = el("div", "std-row");
  // 用 textarea：标准号/条款号常有多行，保留录入时的换行与空格，原样写进 Word
  const inp = el("textarea"); inp.value = t.standard || ""; inp.rows = 3; inp.className = "std-input";
  inp.placeholder = "可多行录入，换行/空格会原样保留到报告里";
  inp.oninput = () => { t.standard = inp.value; scheduleSave(); };
  const btn = el("button", "btn-mini", "参考客户大纲");
  btn.title = "点击填入“参考客户大纲”";
  btn.onclick = () => { t.standard = "参考客户大纲"; inp.value = t.standard; scheduleSave(); };
  row.appendChild(inp); row.appendChild(btn);
  f.appendChild(row);
  return f;
}

function renderTest(t, idx) {
  const card = el("div", "card test-card");
  const head = el("div", "head");
  head.innerHTML = `<span>测试项目 ${idx + 1}：<b>${t.title || "未命名"}</b></span>`;
  head.setAttribute("data-toggle", "");
  const saveScheme = el("button", "btn-mini", "保存为方案");
  saveScheme.style.marginLeft = "auto";
  saveScheme.onclick = (e) => { e.stopPropagation(); saveTestScheme(t); };
  head.appendChild(saveScheme);
  const del = el("button", "btn-del btn-mini", "删除");
  del.style.marginLeft = "6px";
  del.onclick = (e) => { e.stopPropagation(); if (confirm("删除该测试项目？")) { state.tests.splice(idx, 1); renderTests(); scheduleSave(); } };
  head.appendChild(del);
  card.appendChild(head);

  const body = el("div", "body");

  // —— 基本 / 汇总字段 ——
  const g1 = el("div", "grid");
  // 测试类型：可搜索下拉（内置预设 + 已保存方案）
  g1.appendChild(renderTypePicker(t, idx));
  g1.appendChild(fieldInput(t, "title", "测试项目名称", "text"));
  g1.appendChild(sampleNoField(t, idx));
  g1.appendChild(standardField(t));
  g1.appendChild(startDateField(t, idx));  // 开始时间：同步测试日期
  g1.appendChild(endDateField(t, idx));     // 完成时间：不得早于开始时间
  g1.appendChild(fieldInput(t, "overall_result", "试验结果（汇总）", "text"));
  body.appendChild(g1);

  // —— 试验描述 ——
  body.appendChild(el("div", "subhead", "试验描述"));
  // 标准库：按测试项目名匹配车厂，选中自动填条件
  body.appendChild(renderOemPicker(t, idx));
  const g2 = el("div", "grid");
  g2.appendChild(fieldInput(t, "env", "环境条件", "text"));
  // 测试日期已隐藏：自动 = 开始时间（见 startDateField）
  g2.appendChild(fieldInput(t, "condition", "试验条件", "textarea"));
  g2.appendChild(fieldInput(t, "requirement", "试验要求", "textarea"));
  body.appendChild(g2);
  // 试验条件配图
  body.appendChild(renderConditionImages(t));

  // —— 试验设备 ——
  body.appendChild(renderEquip(t));
  // —— 试验结论 ——
  body.appendChild(renderSamples(t));
  // —— 试验图片 ——
  body.appendChild(renderImageSection(t, idx));

  card.appendChild(body);
  return card;
}

function renderTest_replace(idx) {
  const list = $("#testList");
  const cards = list.children;
  const fresh = renderTest(state.tests[idx], idx);
  list.replaceChild(fresh, cards[idx]);
  scheduleSave();
}

// ============ 测试类型可搜索下拉（内置预设 + 已保存方案） ============
function renderTypePicker(t, idx) {
  const tf = el("div", "field combo");
  tf.innerHTML = `<label>测试类型（可搜索：内置预设 / 已保存方案）</label>`;
  const box = el("div", "combo-box");
  const inp = el("input");
  inp.placeholder = "点此搜索或选择…";
  inp.value = t.title || "";
  const menu = el("div", "combo-menu");
  menu.style.display = "none";
  box.appendChild(inp); box.appendChild(menu);
  tf.appendChild(box);

  // 候选项：预设(标 预设) + 方案(标 方案) + 自定义
  function buildItems(filter) {
    const f = (filter || "").trim().toLowerCase();
    const items = [];
    const seen = new Set();
    META.types.forEach(n => { items.push({ name: n, kind: "preset" }); seen.add(n); });
    SCHEMES.forEach(n => { if (!seen.has(n)) { items.push({ name: n, kind: "scheme" }); seen.add(n); } });
    Object.keys(STD).forEach(n => { if (!seen.has(n)) { items.push({ name: n, kind: "standard" }); seen.add(n); } });
    return items.filter(it => !f || fuzzyMatch(it.name, f));
  }

  function openMenu() { renderMenu(inp.value); menu.style.display = "block"; }
  function closeMenu() { setTimeout(() => { menu.style.display = "none"; }, 150); }

  function renderMenu(filter) {
    menu.innerHTML = "";
    const items = buildItems(filter);
    if (!items.length) {
      const d = el("div", "combo-empty", "无匹配。回车即用当前文字作为自定义名称。");
      menu.appendChild(d);
    }
    items.forEach(it => {
      const row = el("div", "combo-item");
      const tag = it.kind === "preset" ? "预设" : (it.kind === "standard" ? "标准库" : "方案");
      row.innerHTML = `<span class="combo-tag combo-${it.kind}">${tag}</span><span>${it.name}</span>`;
      if (it.kind === "scheme") {
        const dx = el("span", "combo-del", "×"); dx.title = "删除此方案";
        dx.onmousedown = async (e) => { e.preventDefault(); e.stopPropagation();
          if (!confirm(`删除方案「${it.name}」？`)) return;
          await fetch("/api/scheme/delete", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: it.name }) });
          await reloadSchemes(); renderMenu(inp.value);
        };
        row.appendChild(dx);
      }
      row.onmousedown = (e) => { e.preventDefault(); chooseItem(it); };
      menu.appendChild(row);
    });
  }

  async function chooseItem(it) {
    if (it.kind === "preset") {
      applyPreset(t, it.name);
    } else if (it.kind === "scheme") {
      await applyScheme(t, it.name);
    } else if (it.kind === "standard") {
      // 标准库项目：设为标题；若只有一个车厂直接套用，多个则提示去下方选
      t.title = it.name;
      const oems = Object.keys(STD[it.name] || {});
      if (oems.length === 1) {
        const e = STD[it.name][oems[0]];
        t.oem = oems[0];
        // 套用标准：先清空条件类内容（含试验条件配图）再带入
        t.standard = e.standard != null ? e.standard : "";
        t.condition = e.condition != null ? e.condition : "";
        t.requirement = e.requirement != null ? e.requirement : "";
        if (e.env != null) t.env = e.env;
        t.condition_images = [];
        toast(`已套用「${it.name} · ${oems[0]}」的条件 ✓`);
      } else if (oems.length > 1) {
        toast(`「${it.name}」有 ${oems.length} 个车厂，请在下方“按车厂”里选`);
      }
    }
    renderTest_replace(idx); scheduleSave();
  }

  inp.onfocus = openMenu;
  inp.onblur = closeMenu;
  inp.oninput = () => { renderMenu(inp.value); };
  // 回车：若正好等于某项则套用，否则作为自定义标题
  inp.onkeydown = (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      const items = buildItems(inp.value);
      const exact = items.find(it => it.name === inp.value.trim());
      if (exact) { chooseItem(exact); }
      else { t.title = inp.value.trim(); menu.style.display = "none"; renderTest_replace(idx); scheduleSave(); }
    }
  };
  return tf;
}

// ============ 标准库：按测试项目匹配车厂，选中自动填条件 ============
function renderOemPicker(t, idx) {
  const wrap = el("div", "oem-picker");
  // 用测试项目名去标准库找；找不到就按 title 模糊匹配 key
  const item = (t.title || "").trim();
  let oems = STD[item] ? Object.keys(STD[item]) : null;
  // 若精确名没命中，尝试忽略大小写/空白的匹配
  let matchedKey = item;
  if (!oems) {
    const k = Object.keys(STD).find(x => x.trim() === item);
    if (k) { oems = Object.keys(STD[k]); matchedKey = k; }
  }
  const label = el("label", null, "从标准库套用条件（按车厂）");
  wrap.appendChild(label);
  if (!oems || !oems.length) {
    const hint = el("span", "oem-hint",
      item ? `标准库里暂无「${item}」的条件。可点右上角“标准库”导入，或手动填下方。`
           : "先在上方填「测试项目名称」，这里会列出该项目各车厂的条件。");
    wrap.appendChild(hint);
    return wrap;
  }
  const row = el("div", "oem-row");
  const sel = el("select", "oem-sel");
  sel.appendChild(el("option", null, "— 选择车厂 —"));
  oems.sort().forEach(o => { const op = el("option", null, o); op.value = o; if (o === t.oem) op.selected = true; sel.appendChild(op); });
  row.appendChild(sel);
  const apply = el("button", "btn-mini", "套用");
  apply.onclick = async () => {
    const oem = sel.value;
    if (!oem || !STD[matchedKey] || !STD[matchedKey][oem]) { toast("请先选车厂"); return; }
    const e = STD[matchedKey][oem];
    t.oem = oem;
    // 套用标准 = 一次干净的覆盖：先清空本测试之前填的条件类内容（含试验条件配图），避免旧内容/旧附图叠加残留
    t.standard = e.standard != null ? e.standard : "";
    t.condition = e.condition != null ? e.condition : "";
    t.requirement = e.requirement != null ? e.requirement : "";
    if (e.env != null) t.env = e.env;
    t.condition_images = [];
    // 标准若带附图（PSD谱/曲线/表格截图等），复制进本项目的"试验条件配图"
    const nImgs = (e.images || []).length;
    if (nImgs) {
      if (!state.name) { alert("该标准带了附图，但项目名称还是空的（它由「委托单号 + 样品型号」自动生成）。\n请先完善委托单号和样品型号，再套用。"); }
      else {
        try {
          const j = await readJSON(await fetch("/api/standards/image/apply", {
            method: "POST", headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ project: state.name, item: matchedKey, oem }) }));
          if (j.ok) {
            // 上面已清空 condition_images，直接带入本标准的附图
            t.condition_images = (j.images || []);
          }
        } catch (err) { /* 图失败不阻断条件套用 */ }
      }
    }
    renderTest_replace(idx);
    scheduleSave();
    toast(`已套用「${matchedKey} · ${oem}」的条件${nImgs ? " + " + nImgs + " 张附图" : ""} ✓`);
  };
  row.appendChild(apply);
  wrap.appendChild(row);
  return wrap;
}

// 从服务器取方案并套用到测试项目（保留已上传图片）
async function applyScheme(t, sname) {
  try {
    const j = await readJSON(await fetch("/api/scheme/get?name=" + encodeURIComponent(sname)));
    if (!j.ok) { alert("读取方案失败：" + (j.error || "")); return; }
    const s = j.scheme || {};
    ["title", "standard", "env", "condition", "requirement", "overall_result"].forEach(k => { if (s[k] != null) t[k] = s[k]; });
    if (s.equipment) t.equipment = JSON.parse(JSON.stringify(s.equipment));
    if (s.image_group_titles) {
      const old = t.image_groups || [];
      t.image_groups = s.image_group_titles.map((title, i) => ({ title, images: (old[i] && old[i].images) ? old[i].images : [] }));
    }
    toast(`已套用方案「${sname}」`);
  } catch (e) { alert(e.message); }
}

// ============ 测试方案库（保存/导入可复用的测试配置） ============
async function saveTestScheme(t) {
  const def = t.title || "";
  const sname = prompt("把这个测试项目存成方案，下次可直接导入。\n\n方案名称：", def);
  if (!sname) return;
  try {
    const r = await fetch("/api/scheme/save", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ scheme_name: sname.trim(), test: t }),
    });
    const j = await readJSON(r);
    if (j.ok) { await reloadSchemes(); toast(`方案「${j.name}」已保存，可在"测试类型"里搜索选用 ✓`); }
    else alert("保存失败：" + (j.error || "未知"));
  } catch (e) { alert(e.message); }
}

function applyPreset(t, type) {
  const p = META.presets[type] || {};
  t.title = type;
  ["standard", "env", "condition", "requirement"].forEach(k => { if (!t[k]) t[k] = p[k] || ""; });
  if (!t.equipment || !t.equipment.length) t.equipment = JSON.parse(JSON.stringify(p.equipment || []));
  if (!t.image_groups || !t.image_groups.length) {
    t.image_groups = (p.image_group_titles || ["图片"]).map(title => ({ title, images: [] }));
  }
}

// —— 试验设备表 ——
function renderEquip(t) {
  if (!t.equipment) t.equipment = [];
  const wrap = el("div");
  wrap.appendChild(el("div", "subhead", "试验设备"));
  const tbl = el("table", "rows");
  tbl.innerHTML = `<tr><th style="width:32px">#</th><th>仪器设备名称</th><th>型号</th><th>管理编号</th><th>校准有效期</th><th style="width:40px"></th></tr>`;
  t.equipment.forEach((e, i) => {
    const tr = el("tr");
    ["name", "model", "mgmt_no", "cal_valid"].forEach((k, j) => {
      const td = el("td");
      if (j === 0) { const idx = el("td", null, String(i + 1)); tr.appendChild(idx); }
      if (k === "name") {
        // 名称格：可搜索的设备选择（按名称/型号/编号搜，编号唯一） + 可手填
        const box = el("div", "equip-name");
        const inp = el("input", "equip-search"); inp.value = e.name || "";
        inp.placeholder = "输入名称/型号/编号搜索，或手填";
        inp.setAttribute("autocomplete", "off");
        const dd = el("div", "equip-dd"); dd.style.display = "none";
        const fillFrom = (d) => {
          e.name = d.name; e.model = d.model || ""; e.mgmt_no = d.mgmt_no || "";
          const s = d.cal_date || "", en = d.cal_end || "";
          e.cal_valid = (s || en) ? `${s}-${en}` : "";
          rerenderTest(t); scheduleSave();
        };
        const renderDD = () => {
          const kw = inp.value.trim();
          dd.innerHTML = "";
          if (!DEV.length) { dd.style.display = "none"; return; }
          const hits = DEV.filter(d => fuzzyMatch(`${d.name} ${d.model} ${d.mgmt_no}`, kw)).slice(0, 12);
          if (!hits.length) { dd.style.display = "none"; return; }
          hits.forEach(d => {
            const it = el("div", "equip-dd-item");
            it.innerHTML = `<b>${escAttr(d.name)}</b>` +
              (d.model ? ` <span class="m">${escAttr(d.model)}</span>` : "") +
              (d.mgmt_no ? ` <span class="no">${escAttr(d.mgmt_no)}</span>` : "");
            it.onmousedown = (ev) => { ev.preventDefault(); fillFrom(d); };
            dd.appendChild(it);
          });
          dd.style.display = "block";
        };
        inp.oninput = () => { e.name = inp.value; scheduleSave(); renderDD(); };
        inp.onfocus = renderDD;
        inp.onblur = () => { setTimeout(() => { dd.style.display = "none"; }, 150); };
        box.appendChild(inp); box.appendChild(dd);
        td.appendChild(box);
        tr.appendChild(td);
        return;
      }
      if (k === "cal_valid") {
        // 校准有效期：两个日期选择器，存 "起-止"
        const box = el("div", "date-range");
        const [a0, b0] = splitRange(e[k]);
        const a = el("input"); a.type = "date"; a.value = toISO(a0);
        const b = el("input"); b.type = "date"; b.value = toISO(b0);
        const sync = () => { const s = toDot(a.value), en = toDot(b.value); e[k] = (s || en) ? `${s}-${en}` : ""; scheduleSave(); };
        a.oninput = sync; b.oninput = sync;
        box.appendChild(a); box.appendChild(el("span", "date-sep", "至")); box.appendChild(b);
        td.appendChild(box);
      } else {
        const inp = el("input"); inp.value = e[k] || "";
        inp.oninput = () => { e[k] = inp.value; scheduleSave(); };
        td.appendChild(inp);
      }
      tr.appendChild(td);
    });
    const tdx = el("td"); const b = el("button", "btn-del btn-mini", "×");
    b.onclick = () => { t.equipment.splice(i, 1); rerenderTest(t); };
    tdx.appendChild(b); tr.appendChild(tdx);
    tbl.appendChild(tr);
  });
  wrap.appendChild(tbl);
  const add = el("button", "btn-mini", "＋ 加一行设备");
  add.onclick = () => { t.equipment.push({ name: "", model: "", mgmt_no: "", cal_valid: "" }); rerenderTest(t); };
  const line = el("div", "rowline"); line.appendChild(add); wrap.appendChild(line);
  return wrap;
}

// 试验结果默认文案（新样品行自动填充，可手改）
const DEFAULT_SAMPLE_RESULT =
  "试验前，样品外观正常，CAN报文正常滚动，对应摄像头正常出图；\n" +
  "试验中，CAN报文正常滚动，对应摄像头正常出图；\n" +
  "试验后，样品外观正常，CAN报文正常滚动，对应摄像头正常出图；\n" +
  "符合试验要求。";

// —— 试验结论（按样品） ——
function renderSamples(t) {
  if (!t.samples) t.samples = [];
  // 样品为空但已填样机编号时，按编号范围自动生成样品行
  if (t.samples.length === 0) syncSamplesFromSampleNo(t);
  const wrap = el("div");
  wrap.appendChild(el("div", "subhead", "试验结论（按样品）"));
  const tbl = el("table", "rows");
  tbl.innerHTML = `<tr><th style="width:70px">样品编号</th><th>试验结果</th><th style="width:70px">结论</th><th style="width:40px"></th></tr>`;
  t.samples.forEach((s, i) => {
    if (!s.result || !s.result.trim()) s.result = DEFAULT_SAMPLE_RESULT;  // 结果空则自动填默认文案（含旧数据）
    const tr = el("tr");
    ["no", "result", "conclusion"].forEach(k => {
      const td = el("td");
      // 试验结果可能是多行文字，用 textarea 以保留换行；编号/结论仍用单行 input
      const inp = el(k === "result" ? "textarea" : "input");
      if (k === "result") inp.rows = 4;
      inp.value = s[k] || "";
      inp.oninput = () => { s[k] = inp.value; scheduleSave(); };
      td.appendChild(inp); tr.appendChild(td);
    });
    const tdx = el("td"); const b = el("button", "btn-del btn-mini", "×");
    b.onclick = () => { t.samples.splice(i, 1); rerenderTest(t); };
    tdx.appendChild(b); tr.appendChild(tdx);
    tbl.appendChild(tr);
  });
  wrap.appendChild(tbl);
  const add = el("button", "btn-mini", "＋ 加一个样品");
  add.onclick = () => {
    // 新样品继承第一个样品的试验结果（如果有的话），否则用默认值
    const baseResult = t.samples.length > 0 ? t.samples[0].result : DEFAULT_SAMPLE_RESULT;
    // 样品编号自动递增：如果第一个是 "8#"，新的就是 "9#"
    let newNo = "";
    if (t.samples.length > 0) {
      const firstNo = (t.samples[0].no || "").trim();
      const match = firstNo.match(/^(\d+)#?$/);  // 匹配 "8#" 或 "8"
      if (match) {
        const num = parseInt(match[1], 10);
        newNo = (num + t.samples.length) + "#";
      }
    }
    t.samples.push({ no: newNo, result: baseResult, conclusion: "合格" });
    rerenderTest(t);
  };
  const line = el("div", "rowline"); line.appendChild(add); wrap.appendChild(line);
  return wrap;
}

function rerenderTest(t) {
  const idx = state.tests.indexOf(t);
  if (idx >= 0) renderTest_replace(idx);
}
// ============ 渲染：试验条件配图 ============
function renderConditionImages(t) {
  if (!t.condition_images) t.condition_images = [];
  const wrap = el("div");
  wrap.appendChild(el("div", "subhead", "试验条件配图（可选，插在“试验条件”下方）"));
  const dz = el("div", "dropzone", "<div class='dz-icon'>🖼️</div><b>把试验条件相关图片拖到这里</b><br>或点击此区域选择，也可截图后按 Ctrl+V 粘贴<br><span style='font-size:12px;color:#a0a8bb'>仅支持图片，可多选</span>");
  const fi = el("input"); fi.type = "file"; fi.accept = "image/*"; fi.multiple = true; fi.style.display = "none";
  // 鼠标移入即设为粘贴目标（无需点击）；点击仍然只是打开选文件对话框
  dz.onmouseenter = () => { PASTE_TARGET = { t, arr: t.condition_images, dz }; setPasteHint(dz); };
  dz.onclick = () => fi.click();
  fi.onchange = () => handleFiles(fi.files, t, t.condition_images);
  dz.ondragover = (e) => { e.preventDefault(); dz.classList.add("over"); };
  dz.ondragleave = () => dz.classList.remove("over");
  dz.ondrop = (e) => { e.preventDefault(); dz.classList.remove("over"); handleFiles(e.dataTransfer.files, t, t.condition_images); };
  wrap.appendChild(dz);
  const thumbs = el("div", "thumbs");
  t.condition_images.forEach((im, ii) => {
    const th2 = el("div", "thumb");
    const ib = el("div", "imgbox");
    const src = imageUrl(im.file);
    const img = el("img"); img.src = src; img.loading = "lazy";
    ib.title = "点击放大查看"; ib.onclick = () => openLightbox(src);
    ib.appendChild(img); th2.appendChild(ib);
    const cap = el("input", "cap"); cap.placeholder = "图注（可选）"; cap.value = im.caption || "";
    cap.oninput = () => { im.caption = cap.value; scheduleSave(); };
    th2.appendChild(cap);
    const x = el("span", "x", "×"); x.title = "删除";
    x.onclick = () => { t.condition_images.splice(ii, 1); rerenderTest(t); };
    th2.appendChild(x);
    thumbs.appendChild(th2);
  });
  wrap.appendChild(thumbs);
  return wrap;
}

// ============ 渲染：试验图片 ============
function renderImageSection(t, tidx) {
  if (!t.image_groups) t.image_groups = [{ title: "试验前图片", images: [] }];
  const wrap = el("div");
  const sh = el("div", "subhead", `试验图片（2列·填满整行·每页约6张）`);
  wrap.appendChild(sh);
  t.image_groups.forEach((g, gi) => wrap.appendChild(renderImageGroup(t, g, gi)));
  const add = el("button", "btn-mini", "＋ 加一个图片分组");
  add.onclick = () => { t.image_groups.push({ title: "新分组", images: [] }); rerenderTest(t); };
  const line = el("div", "rowline"); line.appendChild(add); wrap.appendChild(line);
  return wrap;
}

function renderImageGroup(t, g, gi) {
  const box = el("div", "imggroup");
  // 分组标题（可编辑）
  const th = el("div", "rowline");
  const ti = el("input"); ti.value = g.title || ""; ti.style.fontWeight = "600"; ti.style.flex = "0 0 200px";
  ti.oninput = () => { g.title = ti.value; scheduleSave(); };
  th.appendChild(el("span", null, "分组标题："));
  th.appendChild(ti);
  const dg = el("button", "btn-del btn-mini", "删除分组");
  dg.onclick = () => { t.image_groups.splice(gi, 1); rerenderTest(t); };
  th.appendChild(dg);
  box.appendChild(th);

  // 拖放区
  const dz = el("div", "dropzone", "<div class='dz-icon'>📷</div><b>把照片拖到这里</b><br>或点击此区域选择，也可截图后按 Ctrl+V 粘贴（可多选）");
  const fileInput = el("input"); fileInput.type = "file"; fileInput.accept = "image/*"; fileInput.multiple = true; fileInput.style.display = "none";
  // 鼠标移入即设为粘贴目标（无需点击）；点击仍然只是打开选文件对话框
  dz.onmouseenter = () => { if (!g.images) g.images = []; PASTE_TARGET = { t, arr: g.images, dz }; setPasteHint(dz); };
  dz.onclick = () => fileInput.click();
  fileInput.onchange = () => { if (!g.images) g.images = []; handleFiles(fileInput.files, t, g.images); };
  dz.ondragover = (e) => { e.preventDefault(); dz.classList.add("over"); };
  dz.ondragleave = () => dz.classList.remove("over");
  dz.ondrop = (e) => { e.preventDefault(); dz.classList.remove("over"); if (!g.images) g.images = []; handleFiles(e.dataTransfer.files, t, g.images); };
  box.appendChild(dz);
  box.appendChild(fileInput);

  // 缩略图
  const thumbs = el("div", "thumbs");
  (g.images || []).forEach((im, ii) => {
    const th2 = el("div", "thumb");
    const ib = el("div", "imgbox");
    const src = imageUrl(im.file);
    const img = el("img"); img.src = src; img.loading = "lazy";
    ib.title = "点击放大查看"; ib.onclick = () => openLightbox(src);
    ib.appendChild(img); th2.appendChild(ib);
    const cap = el("input", "cap"); cap.placeholder = "图注（可选）"; cap.value = im.caption || "";
    cap.oninput = () => { im.caption = cap.value; scheduleSave(); };
    th2.appendChild(cap);
    const x = el("span", "x", "×"); x.title = "删除";
    x.onclick = () => { g.images.splice(ii, 1); rerenderTest(t); };
    th2.appendChild(x);
    thumbs.appendChild(th2);
  });
  box.appendChild(thumbs);
  return box;
}
// 点击缩略图 → 全屏查看原图
function openLightbox(src) {
  const lb = el("div", "lightbox");
  const big = el("img"); big.src = src;
  const close = el("span", "lb-close", "×");
  lb.appendChild(big); lb.appendChild(close);
  lb.onclick = () => document.body.removeChild(lb);
  const onEsc = (e) => { if (e.key === "Escape" && lb.parentNode) { document.body.removeChild(lb); document.removeEventListener("keydown", onEsc); } };
  document.addEventListener("keydown", onEsc);
  document.body.appendChild(lb);
}

// ============ 上传 ============
function isImageFile(f) {
  if (f.type && f.type.startsWith("image/")) return true;
  return /\.(jpe?g|png|bmp|gif|webp)$/i.test(f.name || "");
}

async function handleFiles(fileList, t, targetArr) {
  if (!state.name) { alert("请先在左上角填写项目名称"); return; }
  const all = [...fileList];
  const files = all.filter(isImageFile);
  const rejected = all.filter(f => !isImageFile(f));
  if (rejected.length) {
    alert("以下文件不是图片，已跳过（报告里只能放静态图片，不能放视频）：\n\n" +
      rejected.map(f => "· " + f.name).join("\n") +
      "\n\n如果要放视频里的画面，请先在手机/电脑上截图，再上传截图。");
  }
  if (!files.length) { status(""); return; }
  await ensureSaved();  // 确保项目目录已建

  // 逐张上传，避免一次请求过大导致连接中断；单张失败自动重试
  let done = 0;
  const failedFiles = [];  // {name, reason}
  $("#btnGen").disabled = true;
  for (const f of files) {
    done++;
    status(`上传中 ${done}/${files.length} …`);
    let ok = false, lastReason = "";
    for (let attempt = 1; attempt <= 3 && !ok; attempt++) {
      if (attempt > 1) { status(`上传中 ${done}/${files.length}（第 ${attempt} 次重试）…`); await sleep(400 * attempt); }
      const fd = new FormData();
      fd.append("project", state.name);
      fd.append("files", f);
      try {
        const r = await fetch("/api/upload", { method: "POST", body: fd });
        const j = await readJSON(r);
        if (j.ok && j.files && j.files.length) {
          j.files.forEach(x => targetArr.push({ file: x.file, orig: x.orig, caption: "" }));
          rerenderTest(t);   // 每张即时显示
          ok = true;
        } else {
          lastReason = (j.skipped && j.skipped[0] && j.skipped[0].reason) || j.error || "服务器未接收";
        }
      } catch (e) {
        lastReason = String(e).includes("Failed to fetch") ? "与后台连接中断" : String(e);
      }
    }
    if (!ok) failedFiles.push({ name: f.name, reason: lastReason });
  }
  $("#btnGen").disabled = false;
  scheduleSave();
  const failed = failedFiles.length;
  if (failed) {
    status(`完成：成功 ${files.length - failed} 张，失败 ${failed} 张`);
    alert(`有 ${failed} 张上传失败（已自动重试 3 次）：\n\n` +
      failedFiles.map(x => `· ${x.name}\n   原因：${x.reason}`).join("\n") +
      `\n\n可稍后重新拖入这几张再试。`);
  } else {
    status(`已添加 ${files.length} 张 ✓`);
  }
}

function sleep(ms) { return new Promise(res => setTimeout(res, ms)); }
// ============ 保存 / 载入 / 生成 ============
function scheduleSave() {
  clearTimeout(saveTimer);
  status(""); // 清空"已保存✓"，避免误导用户以为当前编辑内容已保存
  saveTimer = setTimeout(doSave, 900);
  scheduleHistory();  // 编辑停顿后把当前状态提交到撤销历史
}

// ===== 撤销/重做实现 =====
// 重置历史：载入/新建项目时调用，把当前状态设为基线
function resetHistory() {
  undoStack = [];
  redoStack = [];
  lastCommitted = JSON.stringify(state);
  clearTimeout(historyTimer);
}

// 编辑停顿 600ms 后提交一次历史；连续打字会合并成一步
function scheduleHistory() {
  clearTimeout(historyTimer);
  historyTimer = setTimeout(commitHistory, 600);
}

// 把「上一个已提交状态」压入撤销栈（仅当有实际变化时）
function commitHistory() {
  clearTimeout(historyTimer);
  const cur = JSON.stringify(state);
  if (lastCommitted === null) { lastCommitted = cur; return; }
  if (cur === lastCommitted) return;  // 无变化不记录
  undoStack.push(lastCommitted);
  if (undoStack.length > UNDO_LIMIT) undoStack.shift();
  redoStack = [];          // 新操作使重做失效
  lastCommitted = cur;
}

// 用一个状态快照(JSON)整体替换当前 state 并重渲染
function applyStateSnapshot(json) {
  const snap = JSON.parse(json);
  state = { name: snap.name || "", info: snap.info || {}, tests: snap.tests || [] };
  $("#pname").value = state.name || "";
  renderInfo(); renderTests();
  // 静默保存到后端(不走 scheduleSave，避免污染历史)
  clearTimeout(saveTimer);
  saveTimer = setTimeout(doSave, 900);
}

function doUndo() {
  commitHistory();  // 先把未提交的改动落进历史
  if (!undoStack.length) { toast("没有可撤销的操作"); return; }
  redoStack.push(JSON.stringify(state));
  const prev = undoStack.pop();
  lastCommitted = prev;
  applyStateSnapshot(prev);
  toast("已撤销");
}

function doRedo() {
  if (!redoStack.length) { toast("没有可重做的操作"); return; }
  undoStack.push(JSON.stringify(state));
  const next = redoStack.pop();
  lastCommitted = next;
  applyStateSnapshot(next);
  toast("已重做");
}

// 绑定 Ctrl+Z 撤销 / Ctrl+Y(或 Ctrl+Shift+Z) 重做
function bindUndoRedo() {
  document.addEventListener("keydown", (e) => {
    if (!(e.ctrlKey || e.metaKey)) return;
    const k = e.key.toLowerCase();
    if (k !== "z" && k !== "y") return;
    // 焦点在输入框/文本域内：交给浏览器原生撤销(oninput 会同步 state)，不拦截
    const ae = document.activeElement;
    if (ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA" || ae.isContentEditable)) return;
    // 弹窗(标准库/设备库)打开时也不接管，避免干扰
    if ($("#stdMgr").style.display === "flex" || $("#devMgr").style.display === "flex") return;
    e.preventDefault();
    if (k === "y" || (k === "z" && e.shiftKey)) doRedo();
    else doUndo();
  });
}

async function doSave() {
  if (!state.name) return;
  try {
    const r = await fetch("/api/save", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(state) });
    const j = await readJSON(r);
    if (j.ok) status("已保存 ✓");
  } catch (e) { status("保存失败"); }
}

async function ensureSaved() {
  if (!state.name) return;
  await doSave();
}

async function loadProject(name) {
  const r = await fetch("/api/load?name=" + encodeURIComponent(name));
  const j = await readJSON(r);
  state = { name: j.name || name, info: applyDefaults(j.info || {}), tests: j.tests || [] };
  $("#pname").value = state.name;
  renderInfo(); renderTests();
  resetHistory();  // 载入的项目作为撤销基线
  status("已载入");
}

async function openDialog() {
  const r = await fetch("/api/projects");
  const names = await readJSON(r);
  if (!names.length) { alert("暂无已保存的项目"); return; }
  const pick = prompt("打开哪个项目？\n\n" + names.map((n, i) => `${i + 1}. ${n}`).join("\n") + "\n\n输入序号或名称：");
  if (!pick) return;
  let name = names[parseInt(pick) - 1] || (names.includes(pick) ? pick : null);
  if (!name) { alert("未找到"); return; }
  await loadProject(name);
}

function newProject() {
  if (state.tests.length && !confirm("新建将清空当前编辑（未保存内容会丢失），继续？")) return;
  state = { name: "", info: applyDefaults({}), tests: [] };
  $("#pname").value = "";
  renderInfo(); renderTests();
  fillNextReportNo();  // 自动填今天的下一个报告编号
  resetHistory();  // 新项目作为撤销基线
  status("新项目");
}

async function serverAlive() {
  try { const r = await fetch("/api/health", { cache: "no-store" }); return r.ok; }
  catch (e) { return false; }
}

// 下载文件：飞书 webview 用 window.open 交给系统处理，普通浏览器直接跳转
function downloadFile(file) {
  const url = "/api/download?file=" + encodeURIComponent(file);
  if (ENV.feishu) { window.open(url, "_blank"); return; }
  window.location = url;
}

// 生成成功后：弹窗展示结果 + 下载/打开操作
function showResultActions(file, size, noun) {
  noun = noun || "报告";
  const mb = (size / 1048576).toFixed(1);
  const mask = $("#resultModal"), fileBox = $("#resultFile"), actions = $("#resultActions");
  if (!mask || !fileBox || !actions) { // 兜底：无弹窗元素时退回下载确认
    if (confirm(`${noun}已生成（${mb} MB）。是否下载？`)) window.location = "/api/download?file=" + encodeURIComponent(file);
    return;
  }
  fileBox.innerHTML = `📄 ${escAttr(file)}<span class="fsize">（${mb} MB）</span>`;
  actions.innerHTML = "";
  const titleReset = mask.querySelector(".modal-head b");
  if (titleReset) titleReset.textContent = "✓ " + noun + "已生成";
  const close = () => { mask.style.display = "none"; };
  const mkBtn = (label, cls, fn) => { const b = el("button", cls); b.textContent = label; b.onclick = fn; return b; };
  // 本机访问才可能让服务器打开 WPS/文件夹；飞书或远程浏览器一律走下载
  const hostLocal = ["localhost", "127.0.0.1", "::1", ""].includes(location.hostname);
  const isLocal = ENV.is_local && hostLocal && !ENV.feishu;
  const titleEl = mask.querySelector(".modal-head b");

  // 下载完成后把弹窗切换成已下载状态：仅保留“完成”，不再提示打开
  const showDownloaded = () => {
    if (titleEl) titleEl.textContent = "✓ " + noun + "已下载";
    actions.innerHTML = "";
    actions.appendChild(mkBtn("完成", "btn-primary", close));
    const tip = el("div", "result-tip", "已开始下载，请去下载文件夹查看。");
    actions.appendChild(tip);
  };

  const dl = () => { downloadFile(file); showDownloaded(); };
  // 主按钮：下载报告（点完切到“已下载”状态）
  actions.appendChild(mkBtn("下载" + noun, "btn-primary", dl));
  // 仅本机(开发机)访问时才提供“打开所在文件夹”，远程访问不显示
  if (isLocal) {
    actions.appendChild(mkBtn("打开所在文件夹", "btn-ghost", async () => {
      try { await fetch("/api/open_folder", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ file }) }); }
      catch (e) { alert(e.message); }
    }));
  }
  // 关闭：点×、点遮罩空白处、Esc
  $("#resultClose").onclick = close;
  mask.onclick = (e) => { if (e.target === mask) close(); };
  document.addEventListener("keydown", function esc(e) {
    if (e.key === "Escape") { close(); document.removeEventListener("keydown", esc); }
  });
  mask.style.display = "flex";
}

function showMultiFileResults(files, noun) {
  noun = noun || "文件";
  const mask = $("#resultModal"), fileBox = $("#resultFile"), actions = $("#resultActions");
  if (!mask || !fileBox || !actions) {
    alert(`${noun}已生成 ${files.length} 个文件。请前往输出目录查看。`);
    return;
  }
  
  // Build file list HTML with checkboxes
  let html = `<div style="margin-bottom:10px;">
    <label style="cursor:pointer;">
      <input type="checkbox" id="selectAllFiles" checked style="margin-right:5px;">
      <b>全选 / 取消全选</b>
    </label>
  </div>`;
  html += `<div style="max-height:200px;overflow-y:auto;border:1px solid #ddd;border-radius:3px;padding:5px;">`;
  files.forEach((f, idx) => {
    const mb = (f.size / 1048576).toFixed(1);
    html += `<div style="margin:5px 0;padding:5px;background:#f5f5f5;border-radius:3px;">`;
    html += `<label style="cursor:pointer;display:block;">`;
    html += `<input type="checkbox" class="file-checkbox" data-file="${escAttr(f.file)}" checked style="margin-right:8px;">`;
    html += `📄 ${escAttr(f.file)} <span class="fsize">(${mb} MB)</span>`;
    html += `</label>`;
    html += `</div>`;
  });
  html += `</div>`;
  fileBox.innerHTML = html;
  
  // Setup select all checkbox handler
  const selectAll = document.getElementById('selectAllFiles');
  const checkboxes = document.querySelectorAll('.file-checkbox');
  
  selectAll.addEventListener('change', () => {
    checkboxes.forEach(cb => cb.checked = selectAll.checked);
  });
  
  // Update select all when individual checkboxes change
  checkboxes.forEach(cb => {
    cb.addEventListener('change', () => {
      const allChecked = [...checkboxes].every(c => c.checked);
      const noneChecked = [...checkboxes].every(c => !c.checked);
      selectAll.checked = allChecked;
      selectAll.indeterminate = !allChecked && !noneChecked;
    });
  });
  
  const titleReset = mask.querySelector(".modal-head b");
  if (titleReset) titleReset.textContent = `✓ ${noun}已生成 (${files.length}个文件)`;
  
  actions.innerHTML = "";
  const close = () => { mask.style.display = "none"; };
  const mkBtn = (label, cls, fn) => { const b = el("button", cls); b.textContent = label; b.onclick = fn; return b; };
  
  // Download selected button
  const downloadBtn = mkBtn(`下载选中`, "btn-primary", () => {
    const selected = [...document.querySelectorAll('.file-checkbox:checked')];
    if (selected.length === 0) {
      alert('请至少选择一个文件');
      return;
    }
    selected.forEach(cb => {
      const fileName = cb.getAttribute('data-file');
      const link = document.createElement('a');
      link.href = "/api/download?file=" + encodeURIComponent(fileName);
      link.download = fileName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    });
    toast(`正在下载 ${selected.length} 个文件`);
  });
  
  actions.appendChild(downloadBtn);
  actions.appendChild(mkBtn("关闭", "btn-secondary", close));
  
  mask.style.display = "flex";
}


// 生成前必填校验：返回问题列表
function validateBeforeGenerate() {
  const problems = [];
  if (!state.info.sample_name) problems.push("首页缺“样品名称”");
  if (!state.info.report_no) problems.push("首页缺“报告编号”");
  state.tests.forEach((t, i) => {
    const tag = `第 ${i + 1} 个测试（${t.title || "未命名"}）`;
    if (!t.title) problems.push(`${tag}：缺“测试项目名称”`);
    if (!t.samples || !t.samples.length) problems.push(`${tag}：缺“试验结论”行`);
    else {
      const empty = t.samples.some(s => !s.conclusion);
      if (empty) problems.push(`${tag}：有样品未填“结论(合格/不合格)”`);
    }
    const hasImg = (t.image_groups || []).some(g => (g.images || []).length);
    if (!hasImg) problems.push(`${tag}：还没有上传任何试验图片`);
  });
  return problems;
}

async function generate() {
  if (!state.name) { alert("请先填写项目名称"); return; }
  if (!state.tests.length) { alert("请至少添加一个测试项目"); return; }
  // 必填校验：有问题时列出，让用户确认是否仍要生成
  const problems = validateBeforeGenerate();
  if (problems.length) {
    const msg = "生成前发现以下问题：\n\n" + problems.map(p => "· " + p).join("\n") +
      "\n\n点“确定”仍然生成（会留空/缺内容），点“取消”返回补全。";
    if (!confirm(msg)) { status("已取消，请补全后再生成"); return; }
  }
  if (!(await serverAlive())) {
    alert("连接不上后台服务。\n\n请检查那个黑色命令行窗口是否还开着：\n如果已关闭，请重新双击「启动.bat」，然后刷新本页面再试。");
    status("服务器未运行"); return;
  }
  await doSave();
  const rb = $("#resultBar"); if (rb) rb.style.display = "none";  // 清掉上次结果条
  status("正在生成报告（图片较多时请稍候）…");
  $("#btnGen").disabled = true;
  try {
    const r = await fetch("/api/generate", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(state) });
    const j = await readJSON(r);
    if (!j.ok) { alert("生成失败：" + (j.error || "")); status("生成失败"); return; }
    status("生成完成 ✓");
    showResultActions(j.file, j.size);
  } catch (e) {
    if (String(e).includes("Failed to fetch")) {
      alert("生成过程中与后台断开了连接。\n\n多半是那个黑色命令行窗口被关闭了。\n请重新双击「启动.bat」，刷新页面后再生成。\n（你的进度已保存，不会丢失。）");
      status("连接中断");
    } else {
      alert("生成出错：" + e); status("");
    }
  }
  finally { $("#btnGen").disabled = false; }
}
async function generateRaw() {
  if (!state.name) { alert("请先填写项目名称"); return; }
  if (!state.tests.length) { alert("请至少添加一个测试项目"); return; }
  if (!(await serverAlive())) {
    alert("连接不上后台服务。\n\n请检查那个黑色命令行窗口是否还开着：\n如果已关闭，请重新双击「启动.bat」，然后刷新本页面再试。");
    status("服务器未运行"); return;
  }
  await doSave();
  const rb = $("#resultBar"); if (rb) rb.style.display = "none";
  status("正在生成原始记录…");
  $("#btnGenRaw").disabled = true;
  try {
    const r = await fetch("/api/generate_raw", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(state) });
    const j = await readJSON(r);
    if (!j.ok) { alert("生成失败：" + (j.error || "")); status("生成失败"); return; }
    status("原始记录生成完成 ✓");
    // Handle single or multiple files
    if (j.files && j.files.length > 1) {
      status(`原始记录生成完成 ✓ (${j.count}个文件)`);
      showMultiFileResults(j.files, "原始记录");
    } else if (j.files && j.files.length === 1) {
      // Single file returned as array
      showResultActions(j.files[0].file, j.files[0].size, "原始记录");
    } else {
      // Legacy single file format
      showResultActions(j.file, j.size, "原始记录");
    }
  } catch (e) {
    if (String(e).includes("Failed to fetch")) {
      alert("生成过程中与后台断开了连接。\n\n多半是那个黑色命令行窗口被关闭了。\n请重新双击「启动.bat」，刷新页面后再生成。\n（你的进度已保存，不会丢失。）");
      status("连接中断");
    } else {
      alert("生成出错：" + e); status("");
    }
  }
  finally { $("#btnGenRaw").disabled = false; }
}
// ============ 初始化 ============
function addTest() {
  const t = { title: "", sample_no: "", standard: "", start_date: "", end_date: "", overall_result: "合格",
    sample_name: state.info.sample_name || "", env: "18℃-28℃、25%RH-75%RH", test_date: "", condition: "", requirement: "",
    equipment: [], samples: [{ no: "1#", result: DEFAULT_SAMPLE_RESULT, conclusion: "合格" }],
    image_groups: [{ title: "试验前图片", images: [] }, { title: "试验中图片", images: [] }, { title: "试验后图片", images: [] }] };
  if (META.types.length) applyPreset(t, META.types[0]);
  state.tests.push(t);
  renderTests();
  scheduleSave();
}

function bindToggles() {
  document.addEventListener("click", (e) => {
    const h = e.target.closest("[data-toggle]");
    if (h && !e.target.closest("button")) h.parentElement.classList.toggle("collapsed");
  });
}

// ============ 标准库管理面板（增删改查 + 导入/模板/清空） ============
let stdEditingKey = null;   // 正在行内编辑的 "item oem"；"__new__" 表示新增行

function stdFlatRows() {
  // 展平为 [{item, oem, standard, condition, requirement}]，按 项目→车厂 排序
  const rows = [];
  Object.keys(STD).forEach(item => {
    Object.keys(STD[item]).forEach(oem => {
      const e = STD[item][oem] || {};
      rows.push({ item, oem, standard: e.standard || "", condition: e.condition || "", requirement: e.requirement || "" });
    });
  });
  rows.sort((a, b) => (a.item + " " + a.oem).localeCompare(b.item + " " + b.oem, "zh"));
  return rows;
}

const SEP = " ";                 // 键分隔符
let stdNewInItem = null;              // 给某测试项目组新增车厂时记录该项目名
let stdImgOpenKey = null;             // 展开附图管理的那条 "item SEP oem"

function renderStdTable() {
  const tb = $("#stdTbody"); if (!tb) return;
  tb.innerHTML = "";
  const filter = ($("#stdSearch").value || "").trim();
  const items = Object.keys(STD).sort((a, b) => a.localeCompare(b, "zh"));
  const totalPairs = items.reduce((n, it) => n + Object.keys(STD[it]).length, 0);
  $("#stdCount").textContent = "共 " + items.length + " 个测试项目、" + totalPairs + " 条(项目/车厂)";

  let shown = 0;
  if (stdEditingKey === "__new__") tb.appendChild(stdFormRow({ item: "", oem: "", standard: "", condition: "", requirement: "" }, "new"));

  items.forEach(item => {
    const oems = Object.keys(STD[item]).sort((a, b) => a.localeCompare(b, "zh"));
    const itemHit = !filter || fuzzyMatch(item, filter);
    const visOems = oems.filter(oem => {
      if (itemHit) return true;
      const e = STD[item][oem] || {};
      return [oem, e.standard, e.condition, e.requirement].some(v => fuzzyMatch(v, filter));
    });
    const addingHere = stdNewInItem === item;
    if (!visOems.length && !addingHere) return;
    shown++;
    tb.appendChild(stdGroupHeader(item, oems.length));
    visOems.forEach(oem => {
      const e = STD[item][oem] || {};
      const r = { item, oem, standard: e.standard || "", condition: e.condition || "", requirement: e.requirement || "", images: e.images || [] };
      const key = item + SEP + oem;
      tb.appendChild(stdEditingKey === key ? stdFormRow(r, "edit") : stdViewRow(r));
      if (stdImgOpenKey === key) tb.appendChild(stdImgRow(r));
    });
    if (addingHere) tb.appendChild(stdFormRow({ item, oem: "", standard: "", condition: "", requirement: "" }, "addoem"));
  });

  const empty = $("#stdEmpty");
  if (!shown && stdEditingKey !== "__new__") {
    empty.style.display = "block";
    empty.textContent = items.length ? "没有匹配的结果，换个词试试。" : "标准库还是空的。点「新增一条」或「导入 Excel」开始。";
  } else { empty.style.display = "none"; }
}

function stdGroupHeader(item, count) {
  const tr = el("tr", "std-grp");
  const td = el("td"); td.colSpan = 6;
  const bar = el("div", "std-grp-bar");
  bar.appendChild(el("span", "std-grp-name", item));
  bar.appendChild(el("span", "std-grp-cnt", count + " 个车厂"));
  bar.appendChild(el("div", "spacer"));
  const addb = el("button", "std-grp-btn", "＋ 加车厂");
  addb.onclick = () => { stdEditingKey = null; stdNewInItem = item; renderStdTable(); };
  const delb = el("button", "std-grp-btn del", "删整项");
  delb.onclick = () => stdDeleteItem(item, count);
  bar.appendChild(addb); bar.appendChild(delb);
  td.appendChild(bar); tr.appendChild(td);
  return tr;
}

function stdViewRow(r) {
  const tr = el("tr", "std-data");
  const td = (txt, cls) => { const c = el("td", cls); c.textContent = txt; return c; };
  tr.appendChild(td("", "std-indent"));
  tr.appendChild(td(r.oem, "std-oem"));
  tr.appendChild(td(r.standard, "wrap"));
  tr.appendChild(td(r.condition, "wrap"));
  tr.appendChild(td(r.requirement, "wrap"));
  const act = el("td");
  const box = el("div", "std-act");
  const key = r.item + SEP + r.oem;
  const nImg = (r.images || []).length;
  const ib = el("button", nImg ? "hasimg" : null, nImg ? `图 ${nImg}` : "图");
  ib.title = "管理该标准的附图（PSD谱/曲线/表格截图等）";
  ib.onclick = () => { stdImgOpenKey = (stdImgOpenKey === key) ? null : key; renderStdTable(); };
  const eb = el("button", null, "编辑"); eb.onclick = () => { stdNewInItem = null; stdEditingKey = key; renderStdTable(); };
  const db = el("button", "del", "删除"); db.onclick = () => stdDelete(r);
  box.appendChild(ib); box.appendChild(eb); box.appendChild(db); act.appendChild(box); tr.appendChild(act);
  return tr;
}

// 附图管理行：缩略图 + 上传区，跨整行显示
function stdImgRow(r) {
  const tr = el("tr", "std-imgrow");
  const td = el("td"); td.colSpan = 6;
  const wrap = el("div", "std-img-wrap");
  const head = el("div", "std-img-head");
  head.appendChild(el("span", "std-img-title", `「${r.item} · ${r.oem}」的附图`));
  head.appendChild(el("span", "std-img-hint", "套用此标准时，这些图会自动插到报告「试验条件」下方"));
  wrap.appendChild(head);

  const thumbs = el("div", "std-img-thumbs");
  (r.images || []).forEach(im => {
    const cell = el("div", "std-img-thumb");
    const box = el("div", "std-img-box");
    const src = `/api/standards/image/${encodeURIComponent(im.file)}`;
    const img = el("img"); img.src = src; img.loading = "lazy";
    box.title = "点击放大"; box.onclick = () => openLightbox(src);
    box.appendChild(img); cell.appendChild(box);
    const cap = el("input", "std-img-cap"); cap.placeholder = "图注（可选）"; cap.value = im.caption || "";
    cap.onchange = () => stdImgCaption(r, im.file, cap.value);
    cell.appendChild(cap);
    const del = el("span", "std-img-del", "×"); del.title = "删除这张图";
    del.onclick = () => stdImgDelete(r, im.file);
    cell.appendChild(del);
    thumbs.appendChild(cell);
  });
  wrap.appendChild(thumbs);

  const dz = el("div", "std-img-dz", "＋ 点击上传附图（PSD谱 / 曲线 / 表格截图等，可多选）");
  const fi = el("input"); fi.type = "file"; fi.accept = "image/*"; fi.multiple = true; fi.style.display = "none";
  // 支持 Ctrl+V 粘贴截图
  if (!r._stdImages) r._stdImages = [];  // 临时数组用于粘贴目标
  dz.onmouseenter = () => { PASTE_TARGET = { t: null, arr: r._stdImages, dz, isStd: true, stdRow: r }; setPasteHint(dz); };
  dz.onclick = () => fi.click();
  fi.onchange = () => stdImgUpload(r, fi.files);
  wrap.appendChild(dz); wrap.appendChild(fi);

  td.appendChild(wrap); tr.appendChild(td);
  return tr;
}

async function stdImgUpload(r, fileList) {
  const files = [...fileList]; if (!files.length) return;
  const fd = new FormData();
  fd.append("item", r.item); fd.append("oem", r.oem);
  files.forEach(f => fd.append("files", f));
  status("正在上传附图…");
  try {
    const j = await readJSON(await fetch("/api/standards/image/upload", { method: "POST", body: fd }));
    if (!j.ok) { alert("上传失败：" + (j.error || "")); status("上传失败"); return; }
    await reloadStandards(); renderStdTable();
    status(`附图已添加，共 ${j.count} 张 ✓`);
  } catch (e) { alert(e.message); status("上传失败"); }
}

async function stdImgDelete(r, file) {
  if (!confirm("确定删除这张附图？")) return;
  try {
    const j = await readJSON(await fetch("/api/standards/image/delete", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ item: r.item, oem: r.oem, file }) }));
    if (!j.ok) { alert("删除失败：" + (j.error || "")); return; }
    await reloadStandards(); renderStdTable(); status("已删除附图 ✓");
  } catch (e) { alert(e.message); }
}

async function stdImgCaption(r, file, caption) {
  try {
    await fetch("/api/standards/image/caption", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ item: r.item, oem: r.oem, file, caption }) });
    if (STD[r.item] && STD[r.item][r.oem]) {
      (STD[r.item][r.oem].images || []).forEach(im => { if (im.file === file) im.caption = caption; });
    }
  } catch (e) { /* 静默 */ }
}

function stdFormRow(r, mode) {
  const tr = el("tr", "std-form");
  const mk = (val, ph, locked, multiline) => {
    const c = el("td");
    // 试验条件/要求常是多行：用 textarea，否则单行 input 会把粘进来的换行压成空格
    const i = el(multiline ? "textarea" : "input"); i.value = val || ""; if (ph) i.placeholder = ph;
    if (multiline) { i.rows = 4; i.className = "std-cell-ta"; }
    if (locked) { i.readOnly = true; i.classList.add("locked"); }
    c.appendChild(i); return { c, i };
  };
  const lockItem = mode === "addoem";
  const fItem = mk(r.item, "测试项目", lockItem), fOem = mk(r.oem, "车厂"),
        fStd = mk(r.standard, "标准号/条款号", false, true), fCond = mk(r.condition, "试验条件", false, true), fReq = mk(r.requirement, "试验要求", false, true);
  [fItem, fOem, fStd, fCond, fReq].forEach(f => tr.appendChild(f.c));
  const act = el("td"); const box = el("div", "std-act");
  const sb = el("button", "save", "保存");
  const oldKey = mode === "edit" ? { old_item: r.item, old_oem: r.oem } : { old_item: "", old_oem: "" };
  sb.onclick = () => stdSave(oldKey, {
    item: fItem.i.value, oem: fOem.i.value, standard: fStd.i.value, condition: fCond.i.value, requirement: fReq.i.value });
  const cb = el("button", null, "取消"); cb.onclick = () => { stdEditingKey = null; stdNewInItem = null; renderStdTable(); };
  box.appendChild(sb); box.appendChild(cb); act.appendChild(box); tr.appendChild(act);
  setTimeout(() => (lockItem ? fOem.i : fItem.i).focus(), 0);
  return tr;
}

async function stdDeleteItem(item, count) {
  if (!confirm("确定删除整个测试项目「" + item + "」及其 " + count + " 个车厂条件？此操作不可撤销。")) return;
  try {
    const j = await readJSON(await fetch("/api/standards/delete", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ item }) }));
    if (!j.ok) { alert("删除失败：" + (j.error || "")); return; }
    stdEditingKey = null; stdNewInItem = null;
    await reloadStandards(); renderStdTable(); renderTests();
    status("已删除整个项目 ✓");
  } catch (e) { alert(e.message); }
}

async function stdSave(oldKey, val) {
  if (!val.item.trim() || !val.oem.trim()) { alert("「测试项目」和「车厂」不能为空。"); return; }
  try {
    const j = await readJSON(await fetch("/api/standards/upsert", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify(Object.assign({}, oldKey, val)) }));
    if (!j.ok) { alert("保存失败：" + (j.error || "")); return; }
    stdEditingKey = null;
    await reloadStandards(); renderStdTable(); renderTests();
    status(j.created ? "已新增 1 条 ✓" : "已更新 ✓");
  } catch (e) { alert(e.message); }
}

async function stdDelete(r) {
  if (!confirm(`确定删除这条？\n\n测试项目：${r.item}\n车厂：${r.oem}`)) return;
  try {
    const j = await readJSON(await fetch("/api/standards/delete", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ item: r.item, oem: r.oem }) }));
    if (!j.ok) { alert("删除失败：" + (j.error || "")); return; }
    if (stdEditingKey === r.item + " " + r.oem) stdEditingKey = null;
    await reloadStandards(); renderStdTable(); renderTests();
    status("已删除 ✓");
  } catch (e) { alert(e.message); }
}

function openStdMgr() {
  stdEditingKey = null;
  $("#stdSearch").value = "";
  $("#stdMgr").style.display = "flex";
  renderStdTable();
}
function closeStdMgr() { $("#stdMgr").style.display = "none"; }

function bindStdMgr() {
  const btn = $("#btnStdMgr"); if (!btn) return;
  btn.onclick = openStdMgr;
  $("#stdMgrClose").onclick = closeStdMgr;
  $("#stdMgr").onclick = (e) => { if (e.target === $("#stdMgr")) closeStdMgr(); };
  $("#stdSearch").oninput = () => renderStdTable();
  $("#stdAdd").onclick = () => { stdEditingKey = "__new__"; renderStdTable(); };
  $("#stdTplBtn").onclick = () => { window.location.href = "/api/standards/template"; };
  $("#stdExportBtn").onclick = () => {
    if (!Object.keys(STD).length) { alert("标准库是空的，没有可导出的内容。"); return; }
    window.location.href = "/api/standards/export";
  };
  // 清空按钮已移除(防误删)，不再绑定
  // 导入 Excel 复用隐藏的 #stdFile
  const fi = $("#stdFile");
  $("#stdImportBtn").onclick = () => fi.click();
  fi.onchange = async () => {
    const f = fi.files && fi.files[0]; fi.value = "";
    if (!f) return;
    status("正在导入标准库…");
    try {
      const fd = new FormData(); fd.append("file", f);
      const j = await readJSON(await fetch("/api/standards/import", { method: "POST", body: fd }));
      if (!j.ok) { alert("导入失败：" + (j.error || "")); status("导入失败"); return; }
      await reloadStandards(); renderStdTable(); renderTests();
      status("标准库已更新 ✓");
      let imgLine = "";
      if (j.images) imgLine = `\n附图：导入 ${j.images} 张` + (j.images_orphan ? `（另有 ${j.images_orphan} 张没对上标准行，已忽略——请把图片放进对应标准那一行的「附图」列格子里）` : "");
      else if (j.images_orphan) imgLine = `\n附图：有 ${j.images_orphan} 张图没对上任何标准行，已忽略（请把图片放进对应标准那一行的「附图」列格子里）。`;
      alert(`导入完成：\n新增 ${j.added} 条，覆盖 ${j.updated} 条，跳过 ${j.skipped} 行。${imgLine}\n\n当前库共 ${j.items} 个测试项目、${j.pairs} 个(项目×车厂)条件。`);
    } catch (e) { alert(e.message); status("导入失败"); }
  };
}

// ============ 设备库（增删改查） ============
let devEditKey = null;   // 正在编辑的设备 mgmt_no；"__new__" 表示新增行

function devDataRow(d) {
  const tr = el("tr");
  const cell = (cls, txt) => { const td = el("td", cls); td.textContent = txt || ""; return td; };
  tr.appendChild(cell(null, d.name));
  tr.appendChild(cell(null, d.model));
  tr.appendChild(cell(null, d.mgmt_no));
  tr.appendChild(cell(null, d.cal_date));
  tr.appendChild(cell("dev-end", d.cal_end));
  tr.appendChild(cell("dev-tt", d.test_type));
  const op = el("td", "dev-act");
  const edit = el("button", "btn-mini", "编辑");
  edit.onclick = () => { devEditKey = d.mgmt_no || ("|" + d.name + "|" + d.model); renderDevTable(); };
  const del = el("button", "btn-del btn-mini", "删除");
  del.onclick = async () => {
    if (!confirm(`确定删除「${d.name}${d.mgmt_no ? " " + d.mgmt_no : ""}」？`)) return;
    const body = d.mgmt_no ? { mgmt_no: d.mgmt_no } : { name: d.name, model: d.model, factory_no: d.factory_no };
    const j = await readJSON(await fetch("/api/devices/delete", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) }));
    if (j.ok) { await reloadDevices(); renderDevTable(); renderTests(); status("已删除设备 ✓"); }
  };
  op.appendChild(edit); op.appendChild(del);
  tr.appendChild(op);
  return tr;
}

function devFormRow(d, isNew) {
  d = d || { name: "", model: "", mgmt_no: "", cal_date: "", test_type: "", factory_no: "", maker: "", remark: "" };
  const tr = el("tr", "dev-form");
  const mk = (val, ph) => { const i = el("input"); i.value = val || ""; if (ph) i.placeholder = ph; return i; };
  const nameI = mk(d.name, "设备名称*");
  const modelI = mk(d.model, "型号（中文会自动去掉）");
  const mgmtI = mk(d.mgmt_no, "管理编号");
  const calI = el("input"); calI.type = "date"; calI.value = toISO(d.cal_date);
  const endCell = el("td", "dev-end"); endCell.textContent = d.cal_end || "自动";
  calI.oninput = () => { const en = calFromISO(calI.value); endCell.textContent = en || "自动"; };
  const ttI = mk(d.test_type, "测试类型（可空）");
  [nameI, modelI, mgmtI].forEach((i, k) => { const td = el("td"); td.appendChild(i); tr.appendChild(td); });
  const calTd = el("td"); calTd.appendChild(calI); tr.appendChild(calTd);
  tr.appendChild(endCell);
  const ttTd = el("td"); ttTd.appendChild(ttI); tr.appendChild(ttTd);
  const op = el("td", "dev-act");
  const save = el("button", "btn-primary btn-mini", "保存");
  save.onclick = async () => {
    const name = nameI.value.trim();
    if (!name) { alert("设备名称不能为空"); return; }
    const body = {
      old_mgmt_no: isNew ? "" : (d.mgmt_no || ""),
      name, model: modelI.value, mgmt_no: mgmtI.value.trim(),
      cal_date: toDot(calI.value), test_type: ttI.value,
      factory_no: d.factory_no || "", maker: d.maker || "", remark: d.remark || "",
    };
    const j = await readJSON(await fetch("/api/devices/upsert", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) }));
    if (!j.ok) { alert(j.error || "保存失败"); return; }
    devEditKey = null;
    await reloadDevices(); renderDevTable(); renderTests();
    status(j.created ? "已新增设备 ✓" : "已更新设备 ✓");
  };
  const cancel = el("button", "btn-mini", "取消");
  cancel.onclick = () => { devEditKey = null; renderDevTable(); };
  op.appendChild(save); op.appendChild(cancel);
  tr.appendChild(op);
  return tr;
}
// YYYY-MM-DD -> 校准有效期止(点分)，供表单里即时预览
function calFromISO(iso) {
  const m = (iso || "").match(/(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return "";
  let y = +m[1]; const mo = +m[2], da = +m[3];
  const d = new Date(y + 1, mo - 1, da);
  d.setDate(d.getDate() - 1);
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}.${p(d.getMonth() + 1)}.${p(d.getDate())}`;
}

function renderDevTable() {
  const tb = $("#devTbody"); if (!tb) return;
  tb.innerHTML = "";
  const kw = ($("#devSearch").value || "").trim();
  const list = DEV.filter(d => !kw || fuzzyMatch(d.name + " " + d.model + " " + d.mgmt_no + " " + (d.test_type || ""), kw));
  $("#devCount").textContent = `共 ${DEV.length} 台设备` + (kw ? `，匹配 ${list.length} 台` : "");
  const empty = $("#devEmpty");
  if (devEditKey === "__new__") tb.appendChild(devFormRow(null, true));
  if (!DEV.length && devEditKey !== "__new__") {
    empty.style.display = "block";
    empty.textContent = "设备库还是空的。点「导入 CSV」选实验室设备清单，或点「＋ 新增设备」手动加。";
    return;
  }
  empty.style.display = (list.length || devEditKey === "__new__") ? "none" : "block";
  if (!list.length && devEditKey !== "__new__") empty.textContent = "没有匹配的设备。";
  list.forEach(d => {
    const key = d.mgmt_no || ("|" + d.name + "|" + d.model);
    tb.appendChild(devEditKey === key ? devFormRow(d, false) : devDataRow(d));
  });
}
function openDevMgr() { $("#devSearch").value = ""; devEditKey = null; $("#devMgr").style.display = "flex"; renderDevTable(); }
function closeDevMgr() { $("#devMgr").style.display = "none"; }
function bindDevMgr() {
  const btn = $("#btnDevMgr"); if (!btn) return;
  btn.onclick = openDevMgr;
  $("#devMgrClose").onclick = closeDevMgr;
  $("#devMgr").onclick = (e) => { if (e.target === $("#devMgr")) closeDevMgr(); };
  $("#devSearch").oninput = () => renderDevTable();
  $("#devAdd").onclick = () => { devEditKey = "__new__"; renderDevTable(); };
  $("#devExportBtn").onclick = () => {
    if (!DEV.length) { alert("设备库是空的，没有可导出的内容。"); return; }
    window.location.href = "/api/devices/export";
  };
  // 清空按钮已移除(防误删)，不再绑定
  const fi = $("#devFile");
  $("#devImportBtn").onclick = () => fi.click();
  fi.onchange = async () => {
    const f = fi.files && fi.files[0]; fi.value = "";
    if (!f) return;
    status("正在导入设备库…");
    try {
      const fd = new FormData(); fd.append("file", f);
      const j = await readJSON(await fetch("/api/devices/import", { method: "POST", body: fd }));
      if (!j.ok) { alert("导入失败：" + (j.error || "")); status("导入失败"); return; }
      await reloadDevices(); renderDevTable(); renderTests();
      status("设备库已更新 ✓");
      alert(`导入完成：设备库共 ${j.count} 台设备。\n（「校准有效期止」已按 校准日期+1年-1天 自动算好；型号里的中文已自动去掉）`);
    } catch (e) { alert(e.message); status("导入失败"); }
  };
}

// ============ 导入试验申请单 PDF，自动回填首页信息 ============
const FORM_FIELD_LABELS = {
  sample_name: "样品名称", sample_no: "样品零件号", sample_model: "样品型号",
  sample_qty: "样品数量",
  verify_phase: "验证阶段", client_name: "委托方名称", client_addr: "委托方地址",
  maker_name: "制造商名称", maker_addr: "制造商地址", commission_no: "委托单号（申请编号）",
  test_items: "检测项目", test_basis: "检测依据",
};
function bindImportForm() {
  const btn = $("#btnImportForm"), fi = $("#formFile");
  if (!btn || !fi) return;
  btn.onclick = () => fi.click();
  fi.onchange = async () => {
    const f = fi.files && fi.files[0]; fi.value = "";
    if (!f) return;
    status("正在识别申请单…");
    try {
      const fd = new FormData(); fd.append("file", f);
      const j = await readJSON(await fetch("/api/import_form", { method: "POST", body: fd }));
      if (!j.ok) { alert("导入失败：" + (j.error || "")); status("导入失败"); return; }
      const fields = j.fields || {};
      // 预览将回填的字段，让用户确认（会覆盖已填的同名字段）
      const preview = Object.keys(fields).map(k => `· ${FORM_FIELD_LABELS[k] || k}：${fields[k]}`).join("\n");
      if (!confirm(`识别到以下信息，将填入首页（覆盖同名已填内容）：\n\n${preview}\n\n确定填入？`)) { status("已取消"); return; }
      Object.assign(state.info, fields);
      autoName();          // 委托单号/样品型号变了，刷新项目名称
      renderInfo();        // 重绘首页
      scheduleSave();
      status(`已从申请单填入 ${Object.keys(fields).length} 项 ✓`);
    } catch (e) { alert(e.message); status("导入失败"); }
  };
}

async function init() {
  // 5 个初始请求互不依赖，并行拉取；比串行快数倍（尤其经隧道时）
  await Promise.all([
    (async () => { try { ENV = await (await fetch("/api/env")).json(); } catch (e) {} })(),
    (async () => { try { META = await (await fetch("/api/meta")).json(); } catch (e) {} })(),
    reloadSchemes(),
    reloadStandards(),
    reloadDevices(),
  ]);
  state.info = applyDefaults(state.info);
  renderInfo(); renderTests();
  fillNextReportNo();  // 首次进入自动填今天的下一个报告编号
  bindStdMgr();
  bindDevMgr();
  bindImportForm();
  // 项目名称自动生成，设为只读，避免手改导致图片目录脱钩
  $("#pname").readOnly = true;
  $("#pname").placeholder = "自动生成＝委托单号-样品型号+试验报告";
  $("#pname").title = "项目名称由“委托单号 - 样品型号 + 试验报告”自动生成，不可手改";
  $("#btnSave").onclick = doSave;
  $("#btnGen").onclick = generate;
  $("#btnGenRaw").onclick = generateRaw;
  $("#btnNew").onclick = newProject;
  $("#btnOpen").onclick = openDialog;
  $("#btnAddTest").onclick = addTest;
  bindToggles();
  bindPasteImages();
  bindUndoRedo();
  resetHistory();  // 首屏状态作为撤销基线
  status("就绪");
}

// 高亮当前粘贴目标框，让用户看清截图会粘到哪
function setPasteHint(dz) {
  document.querySelectorAll(".dropzone.paste-active").forEach(e => e.classList.remove("paste-active"));
  if (dz) dz.classList.add("paste-active");
}
// 截图后 Ctrl+V：把剪贴板里的图片放进鼠标当前所在的上传区
function bindPasteImages() {
  document.addEventListener("paste", (e) => {
    const items = (e.clipboardData && e.clipboardData.items) || [];
    const imgs = [];
    for (const it of items) {
      if (it.kind === "file" && it.type && it.type.startsWith("image/")) {
        const f = it.getAsFile();
        if (f) imgs.push(f);
      }
    }
    if (!imgs.length) return;          // 剪贴板里没有图片，交给默认行为(粘文字)
    e.preventDefault();
    if (!PASTE_TARGET) { alert("请先把鼠标移到要粘贴的图片区域上（框会高亮），再按 Ctrl+V"); return; }
    // 给截图起个带时间戳的文件名，避免同名覆盖
    const stamp = new Date().toISOString().replace(/[-:T.]/g, "").slice(0, 14);
    const named = imgs.map((f, i) => {
      const ext = (f.type.split("/")[1] || "png").replace("jpeg", "jpg");
      return new File([f], `粘贴截图_${stamp}_${i + 1}.${ext}`, { type: f.type });
    });
    // 标准库的图片用专门的上传接口
    if (PASTE_TARGET.isStd) {
      stdImgUpload(PASTE_TARGET.stdRow, named);
    } else {
      handleFiles(named, PASTE_TARGET.t, PASTE_TARGET.arr);
    }
  });
}
init();
