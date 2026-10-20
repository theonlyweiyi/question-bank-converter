# 题库格式转换工具

把各年份、各格式的题库 Excel（xlsx/xls）一键转换成《导题模板》导入格式的单文件 HTML 工具。**零依赖、零安装、数据不出本机**——下载后双击即可在浏览器里使用，题库文件不会上传到任何服务器。

## 功能特性

- **双布局识别**：竖排（一题一行、选项分列 A/B/C/D）与横排（选项纵向堆叠在题干下方）都能自动解析；表头同义词组匹配，表头不在前 20 行也能找到（搜索窗口 50 行）
- **跨文件合并**：多个文件、多个 sheet 一次拖入，自动按题型归并
- **全局去重**：题干归一化后跨文件去重，明细可查
- **同题干冲突检测**：同一题干答案不同（critical）/ 选项有差异，只报告不拦截，支持"仅看答案冲突"筛选
- **判型不一致报告**：源文件标注题型 vs 按答案推断题型不符的单独列出
- **合并单元格保护**：跨行合并的题干/选项按分组保护，不串行
- **判断题归一**：「对/错/√/×/是/否/A/B」等 10 种写法统一为「正确/错误」；选项列被误填内容的给出警告（tf-options）
- **诊断徽章**：无法识别的 sheet 悬停问号徽章，告诉你缺了哪一列（题干/答案/选项）
- **预览联动**：冲突组、判型不符行点击即定位到对应题目；题干/选项/文件名完整显示不截断（CSS 省略号 + 悬停全览）
- **成品即模板**：导出结果直接匹配《导题模板》列结构

## 快速使用

方式一：直接打开本仓库的 **[index.html](index.html)**（或访问 Pages 在线版），把题库 xlsx/xls 拖进去，点"开始转换"，下载成品 xlsx。

方式二：GitHub Pages 在线版（仓库开启 Pages 后可用）：`https://<用户名>.github.io/question-bank-converter/`

> `index.html` 即本地工作环境中的「题库格式转换工具.html」，内容完全一致。

## 从源码构建

```bash
node assemble.js   # 读 src 三件套，产出根目录 index.html（约 1MB）
```

构建产物为单文件：SheetJS 0.20.3 + 转换核心 + UI 全部内联，无任何外部引用（assemble.js 会强制校验模板里没有 `src=`/`href=` 外链）。

## 运行测试

```bash
node run_tests.js            # 单元测试（合成工作簿）+ 快照比对，约 70 项
node run_tests.js --update   # 源文件合法变更后重新固化 snapshot.json
```

- **单元测试**不依赖任何真实文件，任何机器上都能全过；
- **快照比对**的期望值冻结在 `snapshot.json`，源文件不在机器上时对应条目自动 SKIP，不会误报；
- 快照条目里记录的是题量/分布/剔除/去重/冲突等计数，**不含任何题目内容**。

端到端测试（真实浏览器跑全流程）：

```bash
# 1) 把题库样本放进 ./test-fixtures/（.xls/.xlsx），或用环境变量指定：
E2E_FILES="路径1;路径2" node e2e_test.js
# 2) 样本与 snapshot.json 聚合值一致时，可再加严格数值断言：
E2E_EXPECT_STRICT=1 E2E_FILES="..." node e2e_test.js
```

e2e 默认跑**结构断言**（统计卡与明细表一致、文件名/题干/选项不截断、显示全部切换、冲突筛选、报告联动定位、下载非空等，与样本无关）；严格模式额外核对 2344 题等聚合数值。浏览器用 playwright-core 启动，可用 `CHROME_PATH` 指定 chrome-headless-shell 路径。

## 目录结构

```
├── index.html            # 成品：单文件工具（双击即用）
├── assemble.js           # 装配脚本：模板 + SheetJS + 核心 → index.html
├── core.js               # 转换核心（Node/浏览器双端）：解析/合并/去重/冲突/判型
├── tool_template.html    # UI 模板（含 /*__XLSX_LIB__*/ 与 /*__CORE_JS__*/ 占位符）
├── xlsx.full.min.js      # SheetJS 0.20.3
├── run_tests.js          # 回归测试入口（单测 + 快照比对）
├── e2e_test.js           # Playwright 端到端测试
├── snapshot.json         # 期望值快照基线
└── test-fixtures/        # （可选）e2e 样本目录，gitignore，不入库
```

## 数据说明

本仓库**不含任何真实题库数据**：快照与测试只记录计数结构。请勿把含题目内容的 xlsx 提交进仓库——放到 `test-fixtures/` 本地跑测试即可（该目录已被 gitignore）。

## 环境要求

- 使用成品 `index.html`：任意现代浏览器，无其他要求
- 构建/测试：Node.js 16+；e2e 另需 `npm i playwright-core` 与本机 Chromium
