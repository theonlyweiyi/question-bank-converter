/* 题库 → 导题模板 转换核心（浏览器 / Node 双端可用）
 * 依赖全局 XLSX（SheetJS）。
 * 智能识别两种源布局，表头一律按「同义词组」匹配，不依赖固定用词：
 *  竖排（附件2式）：题干|选项|是否正确答案|... 一题占多行，Y 标记正确选项
 *     兼容表头前移（序号|题号|题干|选项|是否正确答案）
 *  横排（附件3/附件1-2023式）：序号|题型|题目|选项A..G|答案|... 一行一题
 *     题干列叫“题干/题目/试题”均可，题型列叫“题型/类型”均可，题型列缺失时按内容启发式判断
 * 输出：A=题干 B=答案 C=解析 D..J=选项A..G（与《导题模板》一致）
 * finalize 输出四类报告：excluded（剔除）/ dups（去重）/ conflicts（同题干冲突）/ typeMismatch（判型与源标注不一致），均只报告不拦截
 * parseFile 另输出 warnings（竖排合并区边界、孤立选项行等分组异常）
 */
(function (root, factory) {
  var core = factory(root.XLSX);
  if (typeof module !== "undefined" && module.exports) module.exports = core;
  else root.QBankCore = core;
})(typeof globalThis !== "undefined" ? globalThis : this, function (XLSX) {
  "use strict";

  var SHEET_MAIN = "工作表 1 - 导入模板";
  var SHEET_ALT = "任选";
  var TF_WORDS = { "正确": 1, "错误": 1, "对": 1, "错": 1, "√": 1, "×": 1, "Y": 1, "N": 1, "是": 1, "否": 1, "TRUE": 1, "FALSE": 1 };

  /* ---------- 表头同义词组：不同年份题库用词不同，按组匹配实现智能识别 ---------- */
  var HDR_SYNONYMS = {
    stem:        ["题干", "题目", "试题", "问题", "题目内容", "考题"],
    type:        ["题型", "类型", "题目类型", "试题类型", "题目性质"],
    answer:      ["答案", "正确答案", "参考答案", "标准答案"],
    option:      ["选项", "备选项"],                              /* 竖排：单列存放全部选项 */
    correctFlag: ["是否正确答案", "是否正确", "是否答案", "答案标记"],
    jiexi:       ["题目解析", "解析", "答案解析", "答案解释", "试题解析"]
  };
  /* 按优先级在同义词组里找第一个存在的表头，返回列号（-1 = 没有） */
  function findCol(h, names) {
    for (var n = 0; n < names.length; n++) {
      var i = h.indexOf(names[n]);
      if (i >= 0) return i;
    }
    return -1;
  }
  function hasOptionLetterCols(h) {
    return h.some(function (v) { return /^选项[A-Z]$/.test(v); });
  }

  function s(v) { return v == null ? "" : String(v).trim(); }
  function upper(v) { return s(v).toUpperCase(); }
  function letters(ans) { return s(ans).replace(/[^A-Za-z]/g, "").toUpperCase(); }
  /* 判断题答案归一：对/√/Y/是/TRUE → 正确；错/×/N/否/FALSE → 错误；原样保留 正确/错误 */
  function normalizeTf(v) {
    var u = upper(v);
    if (u === "对" || u === "√" || u === "Y" || u === "是" || u === "TRUE" || u === "T") return "正确";
    if (u === "错" || u === "×" || u === "N" || u === "否" || u === "FALSE" || u === "F") return "错误";
    return s(v);
  }
  function isTfWord(v) { return !!TF_WORDS[s(v)] || !!TF_WORDS[upper(v)]; }
  /* 源类型标注 → 标准题型（用于判型不一致报告；无法识别返回空串） */
  function typeLabelOf(v) {
    var t = s(v);
    if (!t) return "";
    if (t.indexOf("判断") >= 0) return "判断题";
    if (t.indexOf("多选") >= 0 || t.indexOf("组合") >= 0) return "多选题";
    if (t.indexOf("单选") >= 0) return "单选题";
    return "";
  }
  function colToIdx(col) { var n = 0; for (var i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64); return n - 1; }

  /* 有些源文件 dimension 虚标到 XFD（16384 列），先按真实单元格裁剪 !ref，
     否则 sheet_to_json 会为每行生成上万个空单元格。 */
  function trimRef(ws) {
    if (!ws || !ws["!ref"]) return;
    var maxCol = -1, maxRow = -1;
    for (var k in ws) {
      if (!Object.prototype.hasOwnProperty.call(ws, k) || k.charAt(0) === "!") continue;
      var m = /^([A-Z]+)(\d+)$/.exec(k);
      if (!m) continue;
      var c = colToIdx(m[1]), r = parseInt(m[2], 10) - 1;
      if (c > maxCol) maxCol = c;
      if (r > maxRow) maxRow = r;
    }
    if (maxCol < 0) return;
    var range = XLSX.utils.decode_range(ws["!ref"]);
    range.e.c = Math.max(maxCol, range.s.c);
    range.e.r = Math.max(maxRow, range.s.r);
    ws["!ref"] = XLSX.utils.encode_range(range);
  }

  function readSheetRows(ws) {
    trimRef(ws);
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "", blankrows: true });
    return aoa.map(function (r) { return (r || []).map(s); });
  }

  /* 表头搜索窗口：有些题库文件前面有几十行标题/填报须知，表头可能出现在较后的行 */
  var HEADER_SEARCH_LIMIT = 50;

  function findHeaderRow(rows, pred) {
    var n = Math.min(rows.length, HEADER_SEARCH_LIMIT);
    for (var i = 0; i < n; i++) if (pred(rows[i] || [])) return i;
    return -1;
  }

  /* ---------- 竖排（附件2式） ----------
   * 表头按同义词组定位（题干|选项|是否正确答案），兼容：
   *  - 题干在第 1 列（2026 版）
   *  - 前面带 序号|题号 前缀列（2024 版）
   * merges：worksheet 的 !merges（合并单元格信息）。
   *  - 题干列被纵向合并时，合并区覆盖的行视为该题的范围；选项行若出现在合并区之外
   *   （即下一个题干出现之前），归属可疑——仍然并入当前题（不丢数据），但记入 warnings。
   *  - 首个题干之前出现选项行（孤立选项）同样记入 warnings。 */
  function parseVertical(rows, sheetName, fileName, merges) {
    var hdr = findHeaderRow(rows, function (r) {
      return findCol(r, HDR_SYNONYMS.stem) >= 0 &&
             findCol(r, HDR_SYNONYMS.option) >= 0 &&
             findCol(r, HDR_SYNONYMS.correctFlag) >= 0;
    });
    if (hdr < 0) return null;
    var h = rows[hdr];
    var colStem = findCol(h, HDR_SYNONYMS.stem);
    var colOpt = findCol(h, HDR_SYNONYMS.option);
    var colY = findCol(h, HDR_SYNONYMS.correctFlag);
    var colJx = findCol(h, HDR_SYNONYMS.jiexi);
    if (colStem < 0 || colOpt < 0 || colY < 0) return null;
    /* 题干列的纵向合并：起始行(0-based) → 结束行(0-based) */
    var mergeEndAt = {};
    (merges || []).forEach(function (m) {
      if (!m || !m.s || m.e.r <= m.s.r) return;             /* 只关心纵向合并 */
      if (m.s.c <= colStem && colStem <= m.e.c) {
        mergeEndAt[m.s.r] = Math.max(mergeEndAt[m.s.r] || 0, m.e.r);
      }
    });
    var isTf = sheetName.indexOf("判断") >= 0;
    var warnings = [];
    var groups = [], cur = null, curBound = -1;
    for (var r = hdr + 1; r < rows.length; r++) {
      var row = rows[r];
      var stem = row[colStem];
      if (stem) {
        cur = { stem: stem, options: [], ansIdx: [], jiexi: colJx >= 0 ? row[colJx] : "", row: r + 1 };
        groups.push(cur);
        curBound = mergeEndAt[r] != null ? mergeEndAt[r] : -1;   /* -1 = 题干未合并，不设边界 */
      }
      if (!cur) {
        if (row[colOpt]) warnings.push({ row: r + 1, kind: "orphan-option", text: "首个题干之前出现选项行，已忽略", qrow: 0 });
        continue;
      }
      var opt = row[colOpt];
      if (opt) {
        if (curBound >= 0 && r > curBound) {
          warnings.push({ row: r + 1, kind: "merge-bounds", qrow: cur.row,
            text: "选项行出现在题干合并区（第 " + cur.row + " 行起）之外，已并入该题，请人工核对是否串题" });
        }
        cur.options.push(opt);
        if (upper(row[colY]) === "Y") cur.ansIdx.push(cur.options.length);
      }
    }
    if (!isTf && groups.length && groups.every(function (g) {
      return g.options.length > 0 && g.options.every(function (o) { return TF_WORDS[o]; });
    })) isTf = true;

    var out = groups.map(function (g) {
      var answer, options;
      if (isTf) {
        if (!g.ansIdx.length) answer = "";
        else {
          var txt = g.options[g.ansIdx[0] - 1];
          answer = isTfWord(txt) ? normalizeTf(txt) : (g.ansIdx[0] === 1 ? "正确" : "错误");
        }
        options = [];
      } else {
        answer = g.ansIdx.map(function (i) { return String.fromCharCode(64 + i); }).join("");
        options = g.options;
      }
      return {
        stem: g.stem, answer: answer, options: options,
        jiexi: g.jiexi === "无" ? "" : g.jiexi,
        type: isTf ? "判断题" : (answer.length > 1 ? "多选题" : "单选题"),
        srcType: sheetName,   /* 竖排没有题型列，用 sheet 名作源标注 */
        src: { file: fileName, sheet: sheetName, row: g.row }
      };
    });
    return { questions: out, layout: "竖排", warnings: warnings };
  }

  /* ---------- 横排（附件3 / 附件1 / 2023 信贷库式） ----------
   * 表头按同义词组定位：题干(题干|题目|试题) + 答案(答案|正确答案) + 至少一个 选项X 列。
   * 题型列（题型|类型）可选：缺失时用内容启发式判断判断题。 */
  function parseHorizontal(rows, sheetName, fileName, merges) {
    var hdr = -1;
    for (var i = 0; i < Math.min(rows.length, HEADER_SEARCH_LIMIT); i++) {
      var r = rows[i] || [];
      if (findCol(r, HDR_SYNONYMS.stem) >= 0 && findCol(r, HDR_SYNONYMS.answer) >= 0 && hasOptionLetterCols(r)) { hdr = i; break; }
    }
    if (hdr < 0) return null;
    var h = rows[hdr];
    var colType = findCol(h, HDR_SYNONYMS.type);
    var colStem = findCol(h, HDR_SYNONYMS.stem);
    var colAns = findCol(h, HDR_SYNONYMS.answer);
    var colJx = findCol(h, HDR_SYNONYMS.jiexi);
    var optCols = [];
    h.forEach(function (v, i) { if (/^选项[A-Z]$/.test(v)) optCols.push(i); });
    if (colStem < 0 || colAns < 0) return null;
    optCols.sort(function (a, b) { return a - b; });

    var out = [], warnings = [];
    for (var r = hdr + 1; r < rows.length; r++) {
      var row = rows[r];
      var stem = row[colStem];
      if (!stem) continue;
      var type = colType >= 0 ? row[colType] : "";
      var rawAns = s(row[colAns]);
      var ans = letters(rawAns);
      var jx = colJx >= 0 ? row[colJx] : "";
      var options = optCols.map(function (c) { return row[c]; });
      /* 判断题判定（智能启发式，按优先级）：
         1) 题型列含“判断”；
         2) 选项列为空且答案是 正确/错误 等 TF 词——无论题型列怎么写
            （覆盖：无题型列的文件、以及类型被错标成“单选题”的题，如 2023 库 r1191） */
      var isTf = type.indexOf("判断") >= 0 ||
        (!options.some(function (o) { return o; }) && isTfWord(rawAns));
      var answer;
      if (isTf) {
        /* 答案列可能是中文“正确/错误”、字母 A/B 或其它 TF 词 */
        if (isTfWord(rawAns)) answer = normalizeTf(rawAns);
        else if (ans.charAt(0) === "A") answer = "正确";
        else if (ans.charAt(0) === "B") answer = "错误";
        else answer = rawAns;
      } else {
        answer = ans;
        /* 兜底：个别题答案列直接写选项文本而非字母（如 2023 库答案写“300”而非“D”）
           → 与选项文本比对，映射为对应字母 */
        if (!answer && rawAns) {
          for (var oi = 0; oi < options.length; oi++) {
            if (options[oi] && options[oi] === rawAns) { answer = String.fromCharCode(65 + oi); break; }
          }
        }
      }
      /* 判断题但选项列有值（如 2026 附件3 的“选项A=正确”变体）：
         按模板规范选项列留空，这里记 warning 提示用户有 N 行选项未带入 */
      if (isTf && options.some(function (o) { return o; })) {
        warnings.push({ row: r + 1, kind: "tf-options",
          text: "题型为判断题但选项列有值（“" + options.filter(function (o) { return o; }).join("、") + "”），按模板规范未带入选项列", qrow: r + 1 });
      }
      out.push({
        stem: stem, answer: answer, options: isTf ? [] : options,
        jiexi: jx === "无" ? "" : jx,
        type: isTf ? "判断题" : (answer.length > 1 ? "多选题" : "单选题"),
        srcType: type,   /* 保留源文件的题型标注，用于判型不一致报告 */
        src: { file: fileName, sheet: sheetName, row: r + 1 }
      });
    }
    return { questions: out, warnings: warnings, layout: "横排" };
  }

  /* ---------- 无法识别诊断：扫描表头窗口，报告缺少什么 ---------- */
  function diagSheet(rows) {
    var all = [].concat(HDR_SYNONYMS.stem, HDR_SYNONYMS.option, HDR_SYNONYMS.correctFlag,
      HDR_SYNONYMS.answer, HDR_SYNONYMS.type);
    var n = Math.min(rows.length, HEADER_SEARCH_LIMIT), best = null, bestRow = -1, bestHits = 0;
    for (var i = 0; i < n; i++) {
      var r = rows[i] || [], hits = 0;
      all.forEach(function (k) { if (r.indexOf(k) >= 0) hits++; });
      if (hits > bestHits) { bestHits = hits; best = r; bestRow = i; }
    }
    if (!best) {
      return { missing: [], hint: "前 " + n + " 行内未找到含「题干/题目/选项/答案」等关键词的表头行。本工具支持两种格式：竖排（题干/选项/是否正确答案 表头）与横排（题干/选项A…/答案 表头）" };
    }
    var miss = [];
    if (findCol(best, HDR_SYNONYMS.stem) < 0) miss.push("题干列（题干/题目/试题）");
    if (findCol(best, HDR_SYNONYMS.answer) < 0) miss.push("答案列（答案/正确答案）");
    if (findCol(best, HDR_SYNONYMS.option) < 0 &&
        !best.some(function (v) { return /^选项[A-Z]$/.test(v); })) miss.push("选项列（选项 或 选项A…）");
    if (findCol(best, HDR_SYNONYMS.correctFlag) < 0) miss.push("是否正确答案列（竖排格式必需）");
    return {
      missing: miss,
      hint: "第 " + (bestRow + 1) + " 行是最接近的表头行" +
        (miss.length ? "，缺少：" + miss.join("、") : "，但列组合不满足任一格式") +
        "。竖排需 题干+选项+是否正确答案；横排需 题干+答案+选项A…"
    };
  }

  /* ---------- 单文件解析（不做跨文件去重） ---------- */
  function parseFile(wb, fileName) {
    var report = { sheets: [], warnings: [] };
    var questions = [];
    (wb.SheetNames || []).forEach(function (name) {
      var ws = wb.Sheets[name];
      if (!ws) return;
      var rows = readSheetRows(ws);
      var hasData = rows.some(function (r) { return r.some(function (c) { return c; }); });
      if (!hasData) { report.sheets.push({ file: fileName, sheet: name, layout: "空表", count: 0 }); return; }
      var merges = ws["!merges"] || [];
      var parsed = parseVertical(rows, name, fileName, merges) || parseHorizontal(rows, name, fileName, merges);
      if (!parsed) {
        report.sheets.push({ file: fileName, sheet: name, layout: "无法识别", count: 0, diag: diagSheet(rows) });
        return;
      }
      report.sheets.push({ file: fileName, sheet: name, layout: parsed.layout, count: parsed.questions.length });
      (parsed.warnings || []).forEach(function (w) {
        report.warnings.push({ file: fileName, sheet: name, row: w.row, kind: w.kind, text: w.text, qrow: w.qrow || 0 });
      });
      questions = questions.concat(parsed.questions);
    });
    return { questions: questions, report: report };
  }

  /* ---------- 校验 + 全局去重 + 同题干冲突检测 + 判型不一致报告（跨所有文件） ---------- */
  /* 题干归一化：去掉所有空白后比较，兼容同一题在不同文件里空格不一致 */
  function normalizeStem(stem) { return s(stem).replace(/\s+/g, ""); }

  function finalize(questions, opts) {
    opts = opts || {};
    var report = { excluded: [], dups: [], conflicts: [], typeMismatch: [] };
    var valid = [];
    questions.forEach(function (q) {
      var nopt = q.options.filter(function (o) { return o; }).length;
      if (!q.stem) return;
      if (!q.answer) { report.excluded.push({ file: q.src.file, sheet: q.src.sheet, row: q.src.row, stem: q.stem, reason: "无正确答案" }); return; }
      if (nopt > 7) { report.excluded.push({ file: q.src.file, sheet: q.src.sheet, row: q.src.row, stem: q.stem, reason: nopt + " 个选项，超出模板 7 选项上限" }); return; }
      if (q.type !== "判断题") {
        for (var i = 0; i < q.answer.length; i++) {
          if (q.answer.charCodeAt(i) - 64 > nopt) {
            report.excluded.push({ file: q.src.file, sheet: q.src.sheet, row: q.src.row, stem: q.stem, reason: "答案字母超出选项范围" });
            return;
          }
        }
      }
      valid.push(q);
    });

    var finalList = [];
    if (opts.dedup === false) {
      finalList = valid;
    } else {
      var seen = {};
      valid.forEach(function (q) {
        var key = normalizeStem(q.stem) + "\u0001" + q.options.join("\u0002") + "\u0001" + q.answer;
        if (seen[key]) { report.dups.push({ file: q.src.file, sheet: q.src.sheet, row: q.src.row, stem: q.stem, reason: "题干+选项+答案重复（忽略空白差异）" }); return; }
        seen[key] = 1;
        finalList.push(q);
      });
    }
    /* 同题干冲突检测（报告不拦截）：题干相同但答案/选项不一致的题，
       全部保留在成品中并列入报告交人工裁决——
       防止“同一题两个答案”进正式题库导致考试误判事故。 */
    var byStem = {};
    finalList.forEach(function (q) {
      var k = normalizeStem(q.stem);
      if (!byStem[k]) byStem[k] = [];
      byStem[k].push(q);
    });
    Object.keys(byStem).forEach(function (k) {
      var g = byStem[k];
      if (g.length < 2) return;
      var ansSet = {}, optSet = {};
      g.forEach(function (q) {
        ansSet[q.answer] = 1;
        optSet[q.options.join("\u0002")] = 1;
      });
      var nAns = Object.keys(ansSet).length, nOpt = Object.keys(optSet).length;
      if (nAns < 2 && nOpt < 2) return; /* 完全重复（去重关闭时会出现）不算冲突 */
      report.conflicts.push({
        stem: g[0].stem,
        diff: nAns >= 2 ? (nOpt >= 2 ? "答案与选项均不同" : "答案不同") : "选项不同",
        critical: nAns >= 2,
        answers: Object.keys(ansSet),
        occurrences: g.map(function (q) {
          return {
            file: q.src.file, sheet: q.src.sheet, row: q.src.row, type: q.type, answer: q.answer,
            options: q.options.filter(function (o) { return o; })
          };
        })
      });
    });
    /* 答案冲突的排前面，同组内出现次数多的排前面 */
    report.conflicts.sort(function (a, b) {
      if (a.critical !== b.critical) return a.critical ? -1 : 1;
      return b.occurrences.length - a.occurrences.length;
    });
    /* 判型不一致报告（报告不拦截）：按答案内容判出的题型与源文件标注不一致时列出，
       覆盖“源标注写错的题”（如 2023 库类型错标成单选题的判断题） */
    finalList.forEach(function (q) {
      var label = typeLabelOf(q.srcType);
      if (label && label !== q.type) {
        report.typeMismatch.push({
          file: q.src.file, sheet: q.src.sheet, row: q.src.row, stem: q.stem,
          srcType: label, judged: q.type,
          reason: "源标注「" + label + "」，按答案内容判为「" + q.type + "」"
        });
      }
    });
    return { questions: finalList, report: report };
  }

  /* 单文件便捷入口（内部等价于 parseFile + finalize） */
  function convertWorkbook(wb, fileName, opts) {
    var p = parseFile(wb, fileName);
    return finalize(p.questions, opts);
  }

  /* ---------- 生成模板工作簿 ---------- */
  function buildAoa(questions) {
    return questions.map(function (q) {
      var row = [q.stem, q.answer, q.jiexi, "", "", "", "", "", "", ""];
      for (var i = 0; i < Math.min(q.options.length, 7); i++) row[3 + i] = q.options[i] || "";
      return row;
    });
  }

  function buildWorkbook(questions) {
    var aoa = buildAoa(questions);
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 29.2 }, { wch: 16.3 }, { wch: 22.2 }, { wch: 16.3 }, { wch: 16.3 }, { wch: 16.3 }, { wch: 16.3 }, { wch: 16.3 }, { wch: 16.3 }, { wch: 16.3 }];
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, SHEET_MAIN);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[""]]), SHEET_ALT);
    if (wb.Props) wb.Props.Title = "题库导入表";
    return wb;
  }

  return {
    SHEET_MAIN: SHEET_MAIN,
    SHEET_ALT: SHEET_ALT,
    parseFile: parseFile,
    finalize: finalize,
    convertWorkbook: convertWorkbook,
    normalizeStem: normalizeStem,
    typeLabelOf: typeLabelOf,
    buildAoa: buildAoa,
    buildWorkbook: buildWorkbook
  };
});
