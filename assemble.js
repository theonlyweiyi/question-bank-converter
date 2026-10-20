/* 装配：把 SheetJS 库与转换核心内联进 HTML 模板，产出最终单文件工具 */
const fs = require("fs");
const path = require("path");

const DIR = __dirname;
const tpl = fs.readFileSync(path.join(DIR, "tool_template.html"), "utf8");
const lib = fs.readFileSync(path.join(DIR, "xlsx.full.min.js"), "utf8");
const core = fs.readFileSync(path.join(DIR, "core.js"), "utf8");

// 防止内联 JS 中出现 </script 提前终止脚本块
function safe(js) {
  return js.replace(/<\/script/gi, "<\\/script");
}

if (!tpl.includes("/*__XLSX_LIB__*/") || !tpl.includes("/*__CORE_JS__*/")) {
  console.error("template placeholders missing");
  process.exit(1);
}
if (/<\s*script[^>]+src\s*=/i.test(tpl) || /<\s*link[^>]+href\s*=/i.test(tpl)) {
  console.error("template contains external resource references!");
  process.exit(1);
}

const out = tpl
  .replace("/*__XLSX_LIB__*/", () => safe(lib))
  .replace("/*__CORE_JS__*/", () => safe(core));

/* 仓库版：产出写到仓库根目录 index.html（本地工作副本输出到上级目录的 题库格式转换工具.html） */
const OUT = path.join(DIR, "index.html");
fs.writeFileSync(OUT, out, "utf8");
console.log("written:", OUT, (fs.statSync(OUT).size / 1024).toFixed(0) + " KB");
