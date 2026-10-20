/* 回归测试入口：快照比对（真实文件）+ 单元测试（合成工作簿，不依赖任何真实文件）
 * 用法：
 *   node run_tests.js            跑全部测试，与 snapshot.json 逐项比对
 *   node run_tests.js --update   用当前源文件重新固化 snapshot.json（源文件合法变更后使用）
 * 设计意图（P0-②）：期望值冻结在 snapshot.json 里，改一行 core.js 就能发现输出漂移；
 *   源文件不在机器上时对应条目 SKIP 而不是误报失败。 */
const fs = require("fs");
const path = require("path");
const XLSX = require("./xlsx.full.min.js");
globalThis.XLSX = XLSX;
const core = require("./core.js");

const SNAP = path.join(__dirname, "snapshot.json");
const UPDATE = process.argv.includes("--update");

/* 源文件清单（与 snapshot.json 的 files 一一对应；快照缺失时以本清单为准） */
const DEFAULT_FILES = [
  { name: "附件2：2026年员工岗位技能培训测试公共题库.xlsx", path: "E:/中转3/附件2：2026年员工岗位技能培训测试公共题库.xlsx" },
  { name: "附件3：2026年运营条线岗位技能测试知识题库.xlsx", path: "E:/中转3/附件3：2026年运营条线岗位技能测试知识题库.xlsx" },
  { name: "附件1：2024年产品管理条线岗位技能测试知识题库.xlsx", path: "D:/desktop/银行文件/2024题库/附件1：2024年产品管理条线岗位技能测试知识题库.xlsx" },
  { name: "附件2：2024年岗位技能培训测试公共知识题库.xlsx", path: "D:/desktop/银行文件/2024题库/附件2：2024年岗位技能培训测试公共知识题库.xlsx" },
  { name: "附件：广西农村合作金融机构2023年信贷岗位资格认证考试试题库.xls", path: "D:/desktop/银行文件/2024题库/附件：广西农村合作金融机构2023年信贷岗位资格认证考试试题库.xls" }
];

/* ---------- 小工具 ---------- */
let pass = 0, fail = 0, skip = 0;
const failures = [];
function ok(cond, label, detail) {
  if (cond) { pass++; console.log("  ✓ " + label); }
  else { fail++; failures.push(label + (detail ? " → " + detail : "")); console.log("  ✗ " + label + (detail ? " → " + detail : "")); }
}
function skipped(label) { skip++; console.log("  - " + label + "（SKIP）"); }
function eq(a, b) { return JSON.stringify(a) === JSON.stringify(b); }
function distOf(qs) {
  const d = { "单选题": 0, "多选题": 0, "判断题": 0 };
  qs.forEach(q => { if (d[q.type] != null) d[q.type]++; });
  return d;
}
function wbFrom(aoa, name) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), name || "S1");
  return wb;
}

/* ============================================================ 单元测试（合成数据） */
console.log("\n■ 单元测试（合成工作簿，不依赖真实文件）");

/* U1 横排同义词表头：题目/类型/正确答案 + 选项X 列 */
{
  const wb = wbFrom([
    ["序号", "类型", "题目", "选项A", "选项B", "选项C", "选项D", "正确答案"],
    [1, "单选题", "合成题一：1+1=？", "1", "2", "3", "4", "B"],
    [2, "多选题", "合成题二：多选", "甲", "乙", "丙", "丁", "AC"],
    [3, "判断题", "合成题三：判断", "", "", "", "", "对"]
  ], "题库");
  const p = core.parseFile(wb, "unit.xlsx");
  ok(p.questions.length === 3, "U1 同义词表头识别 3 题", "got " + p.questions.length);
  ok(p.questions[0].answer === "B", "U1 单选答案取字母", "got " + p.questions[0].answer);
  ok(p.questions[1].answer === "AC" && p.questions[1].type === "多选题", "U1 多选答案与判型", "got " + p.questions[1].answer);
  ok(p.questions[2].answer === "正确" && p.questions[2].options.every(o => !o), "U1 判断题「对」归一且选项留空", "got " + p.questions[2].answer);
}

/* U2 判断题答案各种写法归一 */
{
  const words = [["正确", "正确"], ["错误", "错误"], ["对", "正确"], ["错", "错误"], ["√", "正确"], ["×", "错误"], ["A", "正确"], ["B", "错误"], ["是", "正确"], ["否", "错误"]];
  const aoa = [["序号", "题型", "题干", "选项A", "选项B", "答案"]];
  words.forEach((w, i) => aoa.push([i + 1, "判断题", "TF题" + i, "", "", w[0]]));
  const qs = core.parseFile(wbFrom(aoa, "判断"), "unit.xlsx").questions;
  const bad = qs.filter((q, i) => q.answer !== words[i][1]);
  ok(bad.length === 0, "U2 判断题答案 10 种写法全部归一", "bad: " + bad.map(q => q.answer).join(","));
}

/* U3 答案列写选项文本 → 映射字母 */
{
  const wb = wbFrom([
    ["序号", "题型", "题干", "选项A", "选项B", "选项C", "选项D", "答案"],
    [1, "单选题", "额度题", "50", "100", "200", "300", "300"]
  ], "题库");
  const q = core.parseFile(wb, "unit.xlsx").questions[0];
  ok(q.answer === "D", "U3 答案文本「300」映射为 D", "got " + q.answer);
}

/* U4 竖排带 序号|题号 前缀列 + 是否正确答案 */
{
  const wb = wbFrom([
    ["序号", "题号", "题干", "选项", "是否正确答案", "题目解析"],
    [1, "Q1", "竖排题一", "选项甲", "Y", "解析甲"],
    ["", "", "", "选项乙", "", ""],
    ["", "", "", "选项丙", "", ""],
    [2, "Q2", "竖排题二", "选项A", "", ""],
    ["", "", "", "选项B", "Y", ""]
  ], "单选");
  const qs = core.parseFile(wb, "unit.xlsx").questions;
  ok(qs.length === 2, "U4 竖排列头前移仍识别 2 题", "got " + qs.length);
  ok(qs[0].answer === "A" && qs[0].options.length === 3, "U4 Y 标记映射为 A、选项收全", "got " + qs[0].answer + "/" + qs[0].options.length);
  ok(qs[1].answer === "B", "U4 第二题答案 B", "got " + qs[1].answer);
}

/* U5 竖排 sheet 名不含“判断”但选项全是 正确/错误 → 按判断题 */
{
  const wb = wbFrom([
    ["题干", "选项", "是否正确答案"],
    [" TF 竖排题", "正确", "Y"],
    ["", "错误", ""]
  ], "题库");
  const q = core.parseFile(wb, "unit.xlsx").questions[0];
  ok(q.type === "判断题" && q.answer === "正确" && q.options.length === 0, "U5 竖排判断题启发式", q.type + "/" + q.answer);
}

/* U6 去重：完全重复只留一条 */
{
  const mk = () => ({ stem: "重复题", answer: "A", options: ["甲", "乙", "丙", "丁"], jiexi: "", type: "单选题", src: { file: "a.xlsx", sheet: "S", row: 1 } });
  const fin = core.finalize([mk(), mk()], { dedup: true });
  ok(fin.questions.length === 1 && fin.report.dups.length === 1, "U6 完全重复去重", fin.questions.length + "/" + fin.report.dups.length);
}

/* U7 同题干不同答案 → 冲突报告、两条都保留 */
{
  const mk = (ans, row) => ({ stem: "冲突题", answer: ans, options: ["甲", "乙", "丙", "丁"], jiexi: "", type: "单选题", src: { file: "a.xlsx", sheet: "S", row: row } });
  const fin = core.finalize([mk("A", 1), mk("B", 2)], { dedup: true });
  ok(fin.questions.length === 2, "U7 冲突题不拦截、全部保留", "got " + fin.questions.length);
  ok(fin.report.conflicts.length === 1 && fin.report.conflicts[0].critical === true, "U7 报告 1 组答案冲突且标记 critical", JSON.stringify(fin.report.conflicts.map(c => c.diff)));
  ok(fin.report.conflicts[0].occurrences.length === 2, "U7 冲突组含 2 处出现位置", "got " + (fin.report.conflicts[0] || {}).occurrences);
}

/* U8 同题干同答案不同选项 → 选项差异（非 critical） */
{
  const mk = (opts, row) => ({ stem: "选项差异题", answer: "C", options: opts, jiexi: "", type: "单选题", src: { file: "a.xlsx", sheet: "S", row: row } });
  const fin = core.finalize([mk(["甲", "乙", "丙", "丁"], 1), mk(["甲", "乙", "丙", "戊"], 2)], { dedup: true });
  ok(fin.report.conflicts.length === 1 && fin.report.conflicts[0].critical === false, "U8 选项差异报告为非 critical", JSON.stringify(fin.report.conflicts.map(c => c.diff)));
}

/* U9 剔除规则：8 选项 / 答案字母越界 / 无答案 */
{
  const base = { stem: "剔除题", jiexi: "", type: "单选题", src: { file: "a.xlsx", sheet: "S", row: 1 } };
  const fin = core.finalize([
    Object.assign({}, base, { stem: "八选项题", answer: "H", options: ["1", "2", "3", "4", "5", "6", "7", "8"] }),
    Object.assign({}, base, { stem: "越界题", answer: "E", options: ["甲", "乙", "丙", "丁"] }),
    Object.assign({}, base, { stem: "无答案题", answer: "", options: ["甲", "乙", "丙", "丁"] })
  ], { dedup: true });
  ok(fin.questions.length === 0 && fin.report.excluded.length === 3, "U9 三类剔除规则全部生效", fin.questions.length + "/" + fin.report.excluded.length);
  ok(fin.report.excluded[0].reason.indexOf("7 选项上限") >= 0, "U9 剔除原因含选项上限说明", fin.report.excluded[0].reason);
}

/* U10 去重关闭时完全重复不算冲突 */
{
  const mk = () => ({ stem: "重复非冲突", answer: "A", options: ["甲", "乙"], jiexi: "", type: "单选题", src: { file: "a.xlsx", sheet: "S", row: 1 } });
  const fin = core.finalize([mk(), mk()], { dedup: false });
  ok(fin.questions.length === 2 && fin.report.conflicts.length === 0, "U10 去重关闭时重复不算冲突", fin.report.conflicts.length + "");
}

/* U11 normalizeStem 忽略空白差异 */
{
  ok(core.normalizeStem(" 同一 题　干 ") === core.normalizeStem("同一题干"), "U11 题干归一化忽略空白", "");
  const mk = (stem) => ({ stem: stem, answer: "A", options: ["甲", "乙"], jiexi: "", type: "单选题", src: { file: "a.xlsx", sheet: "S", row: 1 } });
  const fin = core.finalize([mk("题干 甲"), mk("题干甲")], { dedup: true });
  ok(fin.report.conflicts.length === 0 && fin.report.dups.length === 1, "U11 仅空白不同的同题按重复处理", fin.report.conflicts.length + "/" + fin.report.dups.length);
}

/* U12 竖排题干纵向合并：合并区内的选项归该题；合并区之外的选项行记 warning 但仍并入 */
{
  const wb = wbFrom([
    ["题干", "选项", "是否正确答案"],
    ["合并题干题", "甲", "Y"],   /* 0-based r1：题干合并 r1..r3 */
    ["", "乙", ""],
    ["", "丙", ""],
    ["", "丁", ""],             /* r4：合并区之外仍出现选项 → warning */
    ["下一题", "A", "Y"],
    ["", "B", ""]
  ], "单选");
  wb.Sheets["单选"]["!merges"] = [{ s: { r: 1, c: 0 }, e: { r: 3, c: 0 } }];
  const p = core.parseFile(wb, "unit.xlsx");
  ok(p.questions.length === 2, "U12 合并题干仍分出 2 题", "got " + p.questions.length);
  ok(p.questions[0].answer === "A" && p.questions[0].options.length === 4, "U12 合并题选项收全（含合并区外一行）", p.questions[0].answer + "/" + p.questions[0].options.length);
  ok(p.questions[1].answer === "A" && p.questions[1].options.length === 2, "U12 下一题不受影响", p.questions[1].answer + "/" + p.questions[1].options.length);
  const mb = p.report.warnings.filter(w => w.kind === "merge-bounds");
  ok(mb.length === 1 && mb[0].row === 5 && mb[0].qrow === 2, "U12 合并区外选项行记 warning（行号+归属题行号）", JSON.stringify(mb));
}

/* U13 孤立选项行（首个题干之前）记 warning 且不丢 */
{
  const wb = wbFrom([
    ["题干", "选项", "是否正确答案"],
    ["", "孤立选项", "Y"],
    ["第一题", "A", "Y"]
  ], "单选");
  const p = core.parseFile(wb, "unit.xlsx");
  ok(p.questions.length === 1, "U13 孤立选项行不成题", "got " + p.questions.length);
  const orph = p.report.warnings.filter(w => w.kind === "orphan-option");
  ok(orph.length === 1 && orph[0].row === 2, "U13 孤立选项 warning 行号正确", JSON.stringify(orph));
}

/* U14 判型不一致：源标注 vs 按答案内容判型 */
{
  const wb = wbFrom([
    ["序号", "类型", "题目", "选项A", "选项B", "选项C", "选项D", "答案"],
    [1, "单选题", "答案却是多选的题", "甲", "乙", "丙", "丁", "AB"],
    [2, "单选题", "类型错标的判断题", "", "", "", "", "正确"],
    [3, "单选题", "正常题", "甲", "乙", "丙", "丁", "C"],
    [4, "多选题", "答案却是单选的题", "甲", "乙", "丙", "丁", "B"]
  ], "题库");
  const qs = core.parseFile(wb, "unit.xlsx").questions;
  const fin = core.finalize(qs, { dedup: true });
  const mm = fin.report.typeMismatch;
  ok(mm.length === 3, "U14 三处判型不一致全部报告", "got " + mm.length + ": " + mm.map(m => m.srcType + "→" + m.judged).join(","));
  ok(mm.some(m => m.srcType === "单选题" && m.judged === "多选题"), "U14 单选→多选 不一致", "");
  ok(mm.some(m => m.srcType === "单选题" && m.judged === "判断题"), "U14 单选→判断 不一致（错标题型）", "");
  ok(mm.some(m => m.srcType === "多选题" && m.judged === "单选题"), "U14 多选→单选 不一致", "");
  ok(fin.questions.length === 4, "U14 判型不一致不拦截，四题全在成品", "got " + fin.questions.length);
}

/* U15 竖排 sheet 名即源标注：多选 sheet 里只标一个 Y → 判型不一致 */
{
  const wb = wbFrom([
    ["题干", "选项", "是否正确答案"],
    ["多选sheet里的单选实质题", "A", "Y"],
    ["", "B", ""],
    ["正常多选题", "A", "Y"],
    ["", "B", "Y"]
  ], "多选");
  const fin = core.finalize(core.parseFile(wb, "unit.xlsx").questions, { dedup: true });
  const mm = fin.report.typeMismatch;
  ok(mm.length === 1 && mm[0].judged === "单选题", "U15 竖排多选 sheet 单 Y 题报告判型不一致", JSON.stringify(mm.map(m => m.judged)));
  ok(fin.questions[0].type === "单选题" && fin.questions[1].type === "多选题", "U15 两题判型本身正确", fin.questions.map(q => q.type).join(","));
}

/* U16 表头搜索窗口 50 行：表头在第 31 行（前 30 行是说明文字）也能识别 */
{
  const filler = [];
  for (let i = 0; i < 30; i++) filler.push(["第 " + (i + 1) + " 行说明文字，非表头"]);
  const wb = wbFrom(filler.concat([
    ["序号", "题型", "题干", "选项A", "选项B", "答案"],
    [1, "单选题", "深处的表头题", "甲", "乙", "A"]
  ]), "题库");
  const p = core.parseFile(wb, "unit.xlsx");
  ok(p.questions.length === 1, "U16 表头在第 31 行仍能识别（窗口 50 行）", "got " + p.questions.length);
}

/* U17 无法识别诊断：无候选行 / 缺关键列 */
{
  const p1 = core.parseFile(wbFrom([["甲", "乙"], [1, 2]], "杂表"), "unit.xlsx");
  const s1 = p1.report.sheets[0];
  ok(s1.layout === "无法识别" && s1.diag && s1.diag.hint.indexOf("未找到") >= 0,
    "U17a 无候选表头行 → 诊断提示未找到", s1.diag ? s1.diag.hint.slice(0, 30) : "no diag");
  const p2 = core.parseFile(wbFrom([
    ["序号", "题干", "选项A", "选项B"],
    [1, "没有答案列的题", "甲", "乙"]
  ], "题库"), "unit.xlsx");
  const s2 = p2.report.sheets[0];
  ok(s2.layout === "无法识别" && s2.diag && s2.diag.missing.some(m => m.indexOf("答案列") >= 0),
    "U17b 缺答案列 → 诊断列出缺失项", s2.diag ? JSON.stringify(s2.diag.missing) : "no diag");
}

/* U18 横排判断题选项列有值（2026 附件3 变体）：按 TF 处理 + 选项不带入 + 记 warning */
{
  const wb = wbFrom([
    ["序号", "题型", "题干", "选项A", "选项B", "答案"],
    [1, "判断题", "变体判断题", "正确", "错误", "A"]
  ], "题库");
  const p = core.parseFile(wb, "unit.xlsx");
  const w = p.report.warnings;
  ok(p.questions.length === 1 && p.questions[0].answer === "正确" && p.questions[0].options.every(o => !o),
    "U18 选项列写 正确/错误 的判断题正确处理", "got " + p.questions[0].answer);
  ok(w.length === 1 && w[0].kind === "tf-options" && w[0].row === 2,
    "U18 记 tf-options 警告且行号正确", JSON.stringify(w));
}

/* ============================================================ 快照比对（真实文件） */
console.log("\n■ 快照比对（真实源文件 ↔ snapshot.json）");

let snap = null;
if (fs.existsSync(SNAP)) snap = JSON.parse(fs.readFileSync(SNAP, "utf8"));
else console.log("  （snapshot.json 不存在，将只做结构输出不比对）");

const files = snap ? snap.files : DEFAULT_FILES.map(f => Object.assign({}, f, { status: "unknown" }));
const available = [];
const results = {};

for (const f of files) {
  const label = f.name.slice(0, 14);
  if (!fs.existsSync(f.path)) {
    if (UPDATE) skipped(label + " 不在机器上，保留原快照条目");
    else skipped(label + " 不在机器上（E 盘/路径变更），无法比对");
    continue;
  }
  const wb = XLSX.read(fs.readFileSync(f.path), { type: "buffer", cellDates: true });
  const p = core.parseFile(wb, f.name);
  const perFin = core.finalize(p.questions, { dedup: true });
  available.push(f.name);
  results[f.name] = {
    sheets: p.report.sheets.map(s => ({ sheet: s.sheet, layout: s.layout, count: s.count })),
    parsed: p.questions.length,
    types: distOf(p.questions),
    excluded: perFin.report.excluded.length,
    dups: perFin.report.dups.length,
    conflicts: perFin.report.conflicts.length,
    conflictsCritical: perFin.report.conflicts.filter(c => c.critical).length,
    typeMismatch: perFin.report.typeMismatch.length,
    warnings: p.report.warnings.length
  };
  if (UPDATE) { console.log("  · " + label + " 已重新解析"); continue; }
  if (!snap) { skipped(label + " 无快照可比对（跑 --update 固化）"); continue; }
  if (f.status === "unavailable") {
    ok(false, label + " 文件出现了但快照标记为不可用", "请审查输出后跑 node run_tests.js --update 重新固化");
    continue;
  }
  console.log("  " + label + "：");
  ok(eq(results[f.name].sheets, f.sheets), "sheet 识别明细一致", JSON.stringify(results[f.name].sheets) + " vs " + JSON.stringify(f.sheets));
  ok(results[f.name].parsed === f.parsed, "解析题数一致（" + f.parsed + "）", "got " + results[f.name].parsed);
  if (f.types) ok(eq(results[f.name].types, f.types), "题型分布一致", JSON.stringify(results[f.name].types) + " vs " + JSON.stringify(f.types));
  else skipped(label + " 题型分布（快照未记录）");
  ok(results[f.name].excluded === f.excluded, "剔除数一致（" + f.excluded + "）", "got " + results[f.name].excluded);
  if (f.dups != null) ok(results[f.name].dups === f.dups, "去重数一致（" + f.dups + "）", "got " + results[f.name].dups);
  if (f.conflicts != null) ok(results[f.name].conflicts === f.conflicts, "冲突组数一致（" + f.conflicts + "）", "got " + results[f.name].conflicts);
  if (f.typeMismatch != null) ok(results[f.name].typeMismatch === f.typeMismatch, "判型不一致数一致（" + f.typeMismatch + "）", "got " + results[f.name].typeMismatch);
  if (f.warnings != null) ok(results[f.name].warnings === f.warnings, "分组警告数一致（" + f.warnings + "）", "got " + results[f.name].warnings);
}

/* 聚合比对：available 集合与快照记录的文件集一致时才比对 */
let agg = null;
let aggWarn = 0;
if (available.length) {
  let all = [];
  for (const f of files) {
    if (!available.includes(f.name)) continue;
    const wb = XLSX.read(fs.readFileSync(f.path), { type: "buffer", cellDates: true });
    const pp = core.parseFile(wb, f.name);
    all = all.concat(pp.questions);
    aggWarn = (aggWarn || 0) + pp.report.warnings.length;
  }
  const fin = core.finalize(all, { dedup: true });
  agg = {
    files: available.slice().sort(),
    total: fin.questions.length,
    types: distOf(fin.questions),
    excluded: fin.report.excluded.length,
    dups: fin.report.dups.length,
    conflicts: fin.report.conflicts.length,
    conflictsCritical: fin.report.conflicts.filter(c => c.critical).length,
    typeMismatch: fin.report.typeMismatch.length,
    warnings: aggWarn || 0
  };
}
if (!UPDATE && snap && snap.finalize_all && agg) {
  const sameSet = eq(agg.files, snap.finalize_all.files.slice().sort());
  if (sameSet) {
    console.log("\n  聚合（" + available.length + " 文件）：");
    ok(agg.total === snap.finalize_all.total, "成品总题数一致（" + snap.finalize_all.total + "）", "got " + agg.total);
    ok(eq(agg.types, snap.finalize_all.types), "成品题型分布一致", JSON.stringify(agg.types) + " vs " + JSON.stringify(snap.finalize_all.types));
    ok(agg.excluded === snap.finalize_all.excluded, "剔除总数一致（" + snap.finalize_all.excluded + "）", "got " + agg.excluded);
    ok(agg.dups === snap.finalize_all.dups, "去重总数一致（" + snap.finalize_all.dups + "）", "got " + agg.dups);
    if (snap.finalize_all.conflicts != null) {
      ok(agg.conflicts === snap.finalize_all.conflicts, "冲突组数一致（" + snap.finalize_all.conflicts + "）", "got " + agg.conflicts);
      ok(agg.conflictsCritical === snap.finalize_all.conflictsCritical, "答案冲突组数一致（" + snap.finalize_all.conflictsCritical + "）", "got " + agg.conflictsCritical);
    } else skipped("聚合冲突数（快照未记录）");
    if (snap.finalize_all.typeMismatch != null) ok(agg.typeMismatch === snap.finalize_all.typeMismatch, "判型不一致数一致（" + snap.finalize_all.typeMismatch + "）", "got " + agg.typeMismatch);
    if (snap.finalize_all.warnings != null) ok(agg.warnings === snap.finalize_all.warnings, "分组警告数一致（" + snap.finalize_all.warnings + "）", "got " + agg.warnings);
  } else {
    skipped("聚合比对：可用文件集与快照记录不同（缺 " + snap.finalize_all.files.filter(n => !available.includes(n)).join("、") + "）");
  }
}

/* ---------- 汇总 / 固化 ---------- */
if (UPDATE) {
  const out = {
    version: 1,
    updated: new Date().toISOString().slice(0, 10),
    note: "期望值快照：改 core.js 后跑 node run_tests.js 比对；源文件合法变更后跑 node run_tests.js --update 重新固化。unavailable 条目为历史验证值，文件恢复后需 --update。",
    files: files.map(f => {
      if (results[f.name]) return Object.assign({}, f, { status: "available" }, results[f.name]);
      return f; /* 保留 unavailable 条目原样 */
    }),
    finalize_all: agg || (snap && snap.finalize_all) || null,
    finalize_all5_historical: snap && snap.finalize_all5_historical ? snap.finalize_all5_historical : {
      files: DEFAULT_FILES.map(f => f.name),
      total: 3571, types: { "单选题": 1316, "多选题": 1040, "判断题": 1215 },
      excluded: 3, dups: 91, conflicts: null,
      note: "2026-10-19 五文件全量验证记录；conflicts 为后续新增检测项当时未统计，E 盘恢复后重新固化"
    }
  };
  fs.writeFileSync(SNAP, JSON.stringify(out, null, 2), "utf8");
  console.log("\n snapshot.json 已固化（available " + available.length + " / " + files.length + "）");
}

console.log("\n========== 结果：✓ " + pass + " 通过，✗ " + fail + " 失败，- " + skip + " 跳过 ==========");
if (failures.length) { console.log("失败项："); failures.forEach(f => console.log("  · " + f)); }
process.exit(fail ? 1 : 0);
