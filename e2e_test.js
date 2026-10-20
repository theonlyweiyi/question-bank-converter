/* 端到端测试：真实浏览器加载 HTML 工具，上传题库样本，核对结果并下载成品
 * 样本文件定位顺序：① 环境变量 E2E_FILES（分号分隔的路径列表）② ./test-fixtures/ 下的 xls/xlsx
 *   都没有时打印说明并跳过（退出码 0），不误报失败
 * 浏览器定位：环境变量 CHROME_PATH 指定 chrome-headless-shell / chrome 可执行文件；
 *   未设置时用 playwright-core 默认查找路径
 * 快照机制（不依赖本机固定路径的期望值冻结）见 run_tests.js / snapshot.json */
const path = require("path");
const fs = require("fs");
const url = require("url");
const { chromium } = require("playwright-core");

const TOOL = url.pathToFileURL(path.join(__dirname, "index.html")).href;

/* 样本文件定位：E2E_FILES > test-fixtures/ > 跳过 */
function resolveFiles() {
  if (process.env.E2E_FILES) {
    return process.env.E2E_FILES.split(";").map(s => s.trim()).filter(Boolean);
  }
  const dir = path.join(__dirname, "test-fixtures");
  if (fs.existsSync(dir)) {
    return fs.readdirSync(dir)
      .filter(f => /\.(xlsx|xls)$/i.test(f))
      .sort()
      .map(f => path.join(dir, f));
  }
  return [];
}
const FILES = resolveFiles();
const DL = path.join(__dirname, "e2e_download.xlsx");

(async () => {
  if (FILES.length === 0) {
    console.log("e2e SKIP：未找到样本文件。把题库 xls/xlsx 放进 ./test-fixtures/（或用 E2E_FILES=路径1;路径2 指定）后重跑。");
    process.exit(0);
  }
  const launchOpts = { headless: true };
  if (process.env.CHROME_PATH) launchOpts.executablePath = process.env.CHROME_PATH;
  const browser = await chromium.launch(launchOpts);
  const page = await browser.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  await page.goto(TOOL, { waitUntil: "load", timeout: 60000 });
  await page.waitForSelector("#dropzone", { timeout: 15000 });

  // 上传三个文件
  await page.setInputFiles("#fileInput", FILES);

  // 等待结果区出现
  await page.waitForSelector("#resultArea:not(.hidden)", { timeout: 120000 });

  const stats = await page.evaluate(() => ({
    total: document.getElementById("stTotal").textContent,
    single: document.getElementById("stSingle").textContent,
    multi: document.getElementById("stMulti").textContent,
    tf: document.getElementById("stTf").textContent,
    excl: document.getElementById("stExcl").textContent,
    dup: document.getElementById("stDup").textContent,
    conflict: document.getElementById("stConflict").textContent,
    mis: document.getElementById("stMis").textContent,
    warn: document.getElementById("stWarn").textContent,
    conflictHot: document.getElementById("stConflictBox").classList.contains("hot"),
    misHot: document.getElementById("stMisBox").classList.contains("hot"),
    conflictGroups: document.querySelectorAll("#conflictBody > div").length,
    criticalBadges: Array.from(document.querySelectorAll("#conflictBody .badge")).filter(b => b.textContent.indexOf("答案") >= 0).length,
    warnAlert: (document.getElementById("errors").textContent || "").indexOf("答案不同") >= 0,
    misRows: document.querySelectorAll("#misBody tr").length,
    warnRows: document.querySelectorAll("#warnBody tr").length,
    sheetRows: document.querySelectorAll("#sheetTbl tbody tr").length,
    sheetInfo: Array.from(document.querySelectorAll("#sheetTbl tbody tr")).map(tr => Array.from(tr.querySelectorAll("td")).map(td => td.textContent.trim()).join(" | ")),
    exclRows: document.querySelectorAll("#exclBody tr").length,
    dupRows: document.querySelectorAll("#dupBody tr").length,
    previewRows: document.querySelectorAll("#previewTbl tbody tr").length,
    dlDisabled: document.getElementById("btnDownload").disabled,
    sheetFileCells: Array.from(document.querySelectorAll("#sheetTbl tbody tr")).map(tr => {
      const cell = tr.querySelector("td.fname .ell");
      return { text: cell ? cell.textContent : "", title: cell ? cell.getAttribute("title") : "" };
    }),
    previewOpts: (function () {
      const arr = [];
      document.querySelectorAll("#previewTbl tbody tr td.opt .ell").forEach(function (el) {
        arr.push({ text: el.textContent, title: el.getAttribute("title") });
      });
      return arr;
    })(),
    previewStems: (function () {
      const arr = [];
      document.querySelectorAll("#previewTbl tbody tr td.stem .ell").forEach(function (el) {
        arr.push({ text: el.textContent, title: el.getAttribute("title") });
      });
      return arr;
    })()
  }));
  console.log("STATS:", JSON.stringify(stats, null, 2));
  const total = Number(stats.total) || 0;

  /* ---- A-⑤ 预览“显示全部”切换（题量 >30 时才出现“显示全部”按钮） ---- */
  let allBtnText0 = "", allView = { rows: 0, mode: "", btn: "", longOpt: null }, back30 = 0;
  if (total > 30) {
    allBtnText0 = await page.evaluate(() => document.getElementById("btnPreviewAll").textContent);
    await page.click("#btnPreviewAll");
    await page.waitForTimeout(600);
    allView = await page.evaluate(() => {
      var best = null;
      document.querySelectorAll("#previewTbl tbody tr td.opt .ell").forEach(function (el) {
        var t = el.getAttribute("title") || "";
        if (!best || t.length > best.title.length) best = { text: el.textContent, title: t };
      });
      return {
        rows: document.querySelectorAll("#previewTbl tbody tr").length,
        mode: document.getElementById("previewMode").textContent,
        btn: document.getElementById("btnPreviewAll").textContent,
        longOpt: best
      };
    });
    console.log("显示全部视图:", JSON.stringify(allView));
    // 切回前 30 题
    await page.click("#btnPreviewAll");
    await page.waitForTimeout(200);
    back30 = await page.evaluate(() => document.querySelectorAll("#previewTbl tbody tr").length);
    console.log("切回前 30 题 rows:", back30);
  }

  /* ---- A-⑥ 冲突区“仅看答案冲突”筛选 ---- */
  let cfView = { groups: 0, badges: 0 }, cfAll = stats.conflictGroups;
  if (stats.conflictGroups > 0) {
    await page.check("#cfCriticalOnly");
    await page.waitForTimeout(200);
    cfView = await page.evaluate(() => ({
      groups: document.querySelectorAll("#conflictBody > div").length,
      badges: Array.from(document.querySelectorAll("#conflictBody .badge")).filter(b => b.textContent.indexOf("答案") >= 0).length
    }));
    console.log("仅看答案冲突:", JSON.stringify(cfView));
    await page.uncheck("#cfCriticalOnly");
    await page.waitForTimeout(150);
    cfAll = await page.evaluate(() => document.querySelectorAll("#conflictBody > div").length);
    console.log("取消筛选 groups:", cfAll);
  }

  /* ---- P1-⑧ 报告联动预览交互 ---- */
  let linkView = null, backView = null, misView = null, linkLabel = "";
  if (stats.conflictGroups > 0) {
    // 1) 点击第一个冲突组的「查看全部 N 处」
    linkLabel = (await page.textContent('#conflictBody a.rowlink[data-gi="0"]')) || "";
    await page.click('#conflictBody a.rowlink[data-gi="0"]');
    await page.waitForTimeout(250);
    linkView = await page.evaluate(() => ({
      mode: document.getElementById("previewMode").textContent,
      rows: document.querySelectorAll("#previewTbl tbody tr").length,
      backVisible: !document.getElementById("btnPreviewBack").classList.contains("hidden"),
      firstCell: (document.querySelector("#previewTbl tbody tr td") || {}).textContent || ""
    }));
    console.log("冲突组定位视图:", JSON.stringify(linkView));
    // 2) 返回默认预览
    await page.click("#btnPreviewBack");
    await page.waitForTimeout(150);
    backView = await page.evaluate(() => ({
      rows: document.querySelectorAll("#previewTbl tbody tr").length,
      mode: document.getElementById("previewMode").textContent,
      backHidden: document.getElementById("btnPreviewBack").classList.contains("hidden")
    }));
    console.log("返回默认预览:", JSON.stringify(backView));
  }
  // 3) 点击判型不一致表格的行号 → 定位单题
  if (stats.misRows > 0) {
    await page.click("#misBody a.rowlink");
    await page.waitForTimeout(250);
    misView = await page.evaluate(() => ({
      rows: document.querySelectorAll("#previewTbl tbody tr").length,
      mode: document.getElementById("previewMode").textContent,
      firstCell: (document.querySelector("#previewTbl tbody tr td") || {}).textContent || ""
    }));
    console.log("判型不一致定位视图:", JSON.stringify(misView));
  }

  // 下载成品
  const [download] = await Promise.all([
    page.waitForEvent("download", { timeout: 30000 }),
    page.click("#btnDownload")
  ]);
  await download.saveAs(DL);
  console.log("downloaded:", download.suggestedFilename(), "->", DL, fs.statSync(DL).size, "bytes");

  // 截图看 UI
  await page.screenshot({ path: path.join(__dirname, "e2e_shot.png"), fullPage: false });

  console.log("page errors:", errors.length ? errors : "none");
  await browser.close();

  /* ================= 判定 ================= */
  let pass = true;
  const strict = process.env.E2E_EXPECT_STRICT === "1";
  const defRows = Math.min(30, total);

  /* ----- 结构断言（与样本无关：任何合法题库跑完都应成立） ----- */
  // S1 统计卡数字与对应明细表行数一致
  const cardRows = { excl: stats.exclRows, dup: stats.dupRows, mis: stats.misRows, warn: stats.warnRows };
  for (const k in cardRows) {
    if (String(cardRows[k]) !== String(stats[k])) { console.log(`FAIL S1 ${k} 卡与表不一致: 卡=${stats[k]} 表行=${cardRows[k]}`); pass = false; }
  }
  // S2 sheet 行数 ≥ 上传文件数（一个文件可有多个 sheet）
  if (stats.sheetRows < FILES.length) { console.log(`FAIL S2 sheet 行数 ${stats.sheetRows} < 文件数 ${FILES.length}`); pass = false; }
  // S3 文件名完整显示：文本与 title 一致且非空
  (stats.sheetFileCells || []).forEach((c, i) => {
    if (!c.title || c.text !== c.title) { console.log(`FAIL S3 sheetFileCells[${i}]: "${c.text}" vs "${c.title}"`); pass = false; }
  });
  // S4 预览默认行数 = min(30, 总数)；下载按钮可用；下载文件非空
  if (stats.previewRows !== defRows) { console.log(`FAIL S4 previewRows: got ${stats.previewRows}, expect ${defRows}`); pass = false; }
  if (stats.dlDisabled) { console.log("FAIL S4 下载按钮应为可用"); pass = false; }
  if (!fs.statSync(DL).size) { console.log("FAIL S4 下载文件为空"); pass = false; }
  // S5 选项结构：字母前缀 + 文本包含完整 title
  if (!(stats.previewOpts || []).length) { console.log("FAIL S5 未采集到选项"); pass = false; }
  (stats.previewOpts || []).forEach((o, i) => {
    if (!/^[A-G]\. /.test(o.text) || o.text.indexOf(o.title) < 0) { console.log(`FAIL S5 previewOpts[${i}]: "${o.text}"`); pass = false; }
  });
  // S6 题干不截断：所有题干文本与 title 一致
  (stats.previewStems || []).forEach((o, i) => {
    if (o.text !== o.title) { console.log(`FAIL S6 previewStems[${i}]: "${String(o.text).slice(0, 40)}…"`); pass = false; }
  });
  // S7 显示全部切换（题量 >30 时）
  if (total > 30) {
    if (allBtnText0.indexOf(String(total)) < 0) { console.log(`FAIL S7 显示全部按钮初始文案: ${allBtnText0}`); pass = false; }
    if (allView.rows !== total || allView.mode.indexOf("全部") < 0 || allView.btn.indexOf("只看前 30") < 0) {
      console.log(`FAIL S7 显示全部视图: ${JSON.stringify(allView)}`); pass = false;
    }
    if (back30 !== defRows) { console.log(`FAIL S7 切回默认预览: got ${back30}, expect ${defRows}`); pass = false; }
  }
  // S8 冲突筛选：勾选后组数 ≤ 全量
  if (cfView.groups > cfAll) { console.log(`FAIL S8 筛选后组数 ${cfView.groups} > 全量 ${cfAll}`); pass = false; }
  // S9 条件高亮与警告条
  if (stats.conflict > 0 && !stats.conflictHot) { console.log("FAIL S9 conflictHot: 冲突数>0 时统计卡应有高亮样式"); pass = false; }
  if (stats.mis > 0 && !stats.misHot) { console.log("FAIL S9 misHot: 判型不符>0 时统计卡应有高亮样式"); pass = false; }
  if (stats.criticalBadges > 0 && !stats.warnAlert) { console.log("FAIL S9 答案冲突警告条缺失"); pass = false; }
  // S10 报告联动：有冲突组时，定位视图行数=组内处数、模式含"冲突组"、行号列为 r+数字；返回后恢复默认
  if (linkView) {
    const m = linkLabel.match(/(\d+)\s*处/);
    const linkN = m ? Number(m[1]) : 0;
    if (linkN > 0 && linkView.rows !== linkN) { console.log(`FAIL S10 冲突组定位视图行数 ${linkView.rows} ≠ 组内处数 ${linkN}`); pass = false; }
    if (linkView.mode.indexOf("冲突组") < 0) { console.log(`FAIL S10 定位视图模式: ${linkView.mode}`); pass = false; }
    if (!/^r\d+$/.test((linkView.firstCell || "").trim())) { console.log(`FAIL S10 行号列应显示 r+行号: ${linkView.firstCell}`); pass = false; }
    if (!backView || backView.rows !== defRows || !backView.backHidden) { console.log(`FAIL S10 返回默认预览: ${JSON.stringify(backView)}`); pass = false; }
  }
  // S11 判型不一致定位：有判型行时，点击行号 → 单题视图
  if (misView) {
    if (misView.rows !== 1 || misView.mode.indexOf("定位") < 0) { console.log(`FAIL S11 判型定位视图: ${JSON.stringify(misView)}`); pass = false; }
    if (!/^r\d+$/.test((misView.firstCell || "").trim())) { console.log(`FAIL S11 判型定位行号列: ${misView.firstCell}`); pass = false; }
  }

  /* ----- 严格数值断言（E2E_EXPECT_STRICT=1；样本与 snapshot.json 的 3 文件聚合值对齐时启用） ----- */
  if (strict) {
    const expect = { total: "2344", single: "874", multi: "627", tf: "843", excl: "1", dup: "36", conflict: "57", mis: "4", warn: "8", sheetRows: "5" };
    for (const k in expect) {
      if (String(stats[k]) !== expect[k]) { console.log(`FAIL ${k}: got ${stats[k]}, expect ${expect[k]}`); pass = false; }
    }
    if (stats.conflictGroups !== 57) { console.log(`FAIL conflictGroups: got ${stats.conflictGroups}, expect 57`); pass = false; }
    if (stats.criticalBadges !== 4) { console.log(`FAIL criticalBadges: got ${stats.criticalBadges}, expect 4`); pass = false; }
    if (stats.exclRows !== 1) { console.log(`FAIL exclRows: got ${stats.exclRows}, expect 1`); pass = false; }
    if (stats.dupRows !== 36) { console.log(`FAIL dupRows: got ${stats.dupRows}, expect 36`); pass = false; }
    if (stats.misRows !== 4) { console.log(`FAIL misRows: got ${stats.misRows}, expect 4`); pass = false; }
    if (stats.warnRows !== 8) { console.log(`FAIL warnRows: got ${stats.warnRows}, expect 8（2023 库 tf-options 警告）`); pass = false; }
    if (stats.previewRows !== 30) { console.log(`FAIL previewRows: got ${stats.previewRows}, expect 30`); pass = false; }
    if (linkView && (linkView.rows !== 2 || !linkView.backVisible || linkView.mode.indexOf("冲突组全部 2 处") < 0)) {
      console.log(`FAIL 冲突组定位视图: ${JSON.stringify(linkView)}`); pass = false;
    }
    if (backView && (backView.rows !== 30 || !backView.backHidden || backView.mode.indexOf("前 30 题") < 0)) {
      console.log(`FAIL 返回默认预览: ${JSON.stringify(backView)}`); pass = false;
    }
    if (misView && (misView.rows !== 1 || misView.mode.indexOf("定位") < 0)) {
      console.log(`FAIL 判型不一致定位视图: ${JSON.stringify(misView)}`); pass = false;
    }
    const longName = "附件：广西农村合作金融机构2023年信贷岗位资格认证考试试题库.xls";
    const cell = (stats.sheetFileCells || []).find(c => c.title === longName);
    if (!cell) { console.log("FAIL sheetFileCells: 未找到完整文件名的 title"); pass = false; }
    else if (cell.text !== longName) { console.log(`FAIL sheetFileCells: 文本被截断 "${cell.text}"`); pass = false; }
    const longStem = (stats.previewStems || []).find(o => o.title && o.title.length > 52);
    if (!longStem) { console.log("FAIL previewStems: 未找到长度超过 52 字符的题干"); pass = false; }
    else if (longStem.text !== longStem.title) { console.log(`FAIL previewStems: 题干文本与 title 不一致`); pass = false; }
    if (!allView.longOpt || allView.longOpt.title.length <= 16) {
      console.log(`FAIL 显示全部视图长选项: ${JSON.stringify(allView.longOpt)}`); pass = false;
    } else if (allView.longOpt.text.indexOf(allView.longOpt.title) < 0) {
      console.log(`FAIL 长选项文本不完整: "${allView.longOpt.text}"`); pass = false;
    }
    if (allBtnText0.indexOf("2344") < 0) { console.log(`FAIL 显示全部按钮初始文案: ${allBtnText0}`); pass = false; }
    if (allView.rows !== 2344 || allView.mode.indexOf("全部 2344") < 0 || allView.btn.indexOf("只看前 30") < 0) {
      console.log(`FAIL 显示全部视图: ${JSON.stringify(allView)}`); pass = false;
    }
    if (back30 !== 30) { console.log(`FAIL 切回前 30 题: got ${back30}`); pass = false; }
    if (cfView.groups !== 4 || cfView.badges !== 4) { console.log(`FAIL 仅看答案冲突: ${JSON.stringify(cfView)}`); pass = false; }
    if (cfAll !== 57) { console.log(`FAIL 取消筛选后应恢复 57 组: got ${cfAll}`); pass = false; }
  } else {
    console.log("（结构断言模式；样本与 snapshot.json 聚合值对齐时设 E2E_EXPECT_STRICT=1 启用数值断言）");
  }
  console.log(pass ? "E2E PASS" : "E2E FAIL");
  process.exit(pass && errors.length === 0 ? 0 : 1);
})().catch(e => { console.error("E2E ERROR:", e); process.exit(2); });
