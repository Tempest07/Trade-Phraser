(() => {
  "use strict";

  const EXCEL_COLUMNS = [
    "谈判日",
    "交易日",
    "债券代码",
    "债券简称",
    "债券类型",
    "净价",
    "收益率(%)",
    "估值收益率",
    "我行方向",
    "面值（万元）",
    "真实交易对手",
    "交易对手",
    "组合",
    "中介",
    "清算速度(0/1)",
    "成本",
    "价差",
    "清算速度",
    "结算方式",
  ];

  const FORMULA_COLUMNS = new Set(["债券简称", "债券类型", "估值收益率"]);

  const COLUMN_FORMATS = {
    谈判日: "yyyy-mm-dd",
    交易日: "yyyy-mm-dd",
    债券代码: "@",
    债券简称: "@",
    债券类型: "@",
    净价: "0.000",
    "收益率(%)": "0.0000",
    估值收益率: "@",
    我行方向: "@",
    "面值（万元）": "0",
    真实交易对手: "@",
    交易对手: "@",
    组合: "@",
    中介: "@",
    "清算速度(0/1)": "0",
    成本: "0.00",
    价差: "0.00",
    清算速度: "@",
    结算方式: "@",
  };

  const COLUMN_WIDTHS = {
    谈判日: 12,
    交易日: 12,
    债券代码: 16,
    债券简称: 22,
    债券类型: 18,
    净价: 10,
    "收益率(%)": 12,
    估值收益率: 14,
    我行方向: 10,
    "面值（万元）": 13,
    真实交易对手: 22,
    交易对手: 18,
    组合: 16,
    中介: 12,
    "清算速度(0/1)": 15,
    成本: 10,
    价差: 10,
    清算速度: 12,
    结算方式: 12,
  };

  const DEFAULT_TABLE_COLUMN_WIDTHS = {
    谈判日: 108,
    交易日: 108,
    债券代码: 138,
    债券简称: 184,
    债券类型: 150,
    净价: 96,
    "收益率(%)": 112,
    估值收益率: 128,
    我行方向: 100,
    "面值（万元）": 126,
    真实交易对手: 190,
    交易对手: 166,
    组合: 140,
    中介: 100,
    "清算速度(0/1)": 142,
    成本: 96,
    价差: 96,
    清算速度: 116,
    结算方式: 116,
  };

  const SAMPLE_TEXT = [
    "【中诚】 174D 012580499 25鄂交投SCP001 2.10 3000 03.05+0 兴业银行 出给 天弘基金",
    "【国利】 1) 2.37Y(休2) 245008.SH 26创控K1 1.62 3000 06.03交易所 兴业银行 出给 工银瑞信基金",
    "【宁波】25穗投06 243375.SH 净价100.567 2000w 6.3+0 华创证券 to 兴业银行",
    "【利顺】4) 卖出 2.98Y 282690.SH(05.29 远) 26建园01 私募债 1.89 5000 05.29交易所 兴业银行 出给 宁银理财",
  ].join("\n");

  const state = {
    activeMode: "paste",
    trades: [],
    diagnostics: [],
    columnWidths: { ...DEFAULT_TABLE_COLUMN_WIDTHS },
    bankNameBeforeEdit: "",
  };

  const $ = (selector) => document.querySelector(selector);

  function pad2(value) {
    return String(value).padStart(2, "0");
  }

  function formatDate(date) {
    return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
  }

  function parseInputDate(value) {
    const match = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(String(value || ""));
    if (!match) return new Date();
    return makeDate(Number(match[1]), Number(match[2]), Number(match[3])) || new Date();
  }

  function makeDate(year, month, day) {
    const parsed = new Date(year, month - 1, day);
    if (
      parsed.getFullYear() !== year ||
      parsed.getMonth() !== month - 1 ||
      parsed.getDate() !== day
    ) {
      return null;
    }
    return parsed;
  }

  function normalizeText(text) {
    if (text === null || text === undefined) return "";

    const replacements = {
      "\u00a0": " ",
      "\u3000": " ",
      "，": " ",
      ",": " ",
      "；": " ",
      ";": " ",
      "：": ":",
      "（": "(",
      "）": ")",
      "＋": "+",
      "％": "%",
    };

    let value = String(text);
    for (const [oldText, newText] of Object.entries(replacements)) {
      value = value.split(oldText).join(newText);
    }

    return value.replace(/\s+/g, " ").trim();
  }

  function escapeRegExp(value) {
    return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  function removeOnce(text, fragment) {
    const normalizedFragment = normalizeText(fragment);
    if (!normalizedFragment) return text;
    return normalizeText(text.replace(new RegExp(escapeRegExp(normalizedFragment)), " "));
  }

  function parseDateText(dateText, year) {
    let value = normalizeText(dateText);
    value = value.replace(/^\d+[)、]\s*/, "");
    value = value.replace("日", "");
    if (!value) return null;

    let match = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(value);
    if (match) {
      return makeDate(Number(match[1]), Number(match[2]), Number(match[3]));
    }

    match = /^(\d{1,2})月(\d{1,2})$/.exec(value);
    if (match) {
      return makeDate(year, Number(match[1]), Number(match[2]));
    }

    match = /^(\d{1,2})[./](\d{1,2})$/.exec(value);
    if (match) {
      return makeDate(year, Number(match[1]), Number(match[2]));
    }

    if (/^\d+$/.test(value)) {
      if (value.length === 4) {
        return makeDate(year, Number(value.slice(0, 2)), Number(value.slice(2)));
      }
      if (value.length === 3) {
        return makeDate(year, Number(value.slice(0, 1)), Number(value.slice(1)));
      }
    }

    return null;
  }

  function parseFaceValue(token) {
    const value = normalizeText(token).toLowerCase();
    const match = /^(\d+(?:\.\d+)?)(千万|kw|k|w|e|万|亿)?$/.exec(value);
    if (!match) return null;

    const number = Number(match[1]);
    const unit = match[2] || "";

    if (unit === "e" || unit === "亿") return Math.round(number * 10000);
    if (unit === "kw" || unit === "k" || unit === "千万") return Math.round(number * 1000);
    if (unit === "w" || unit === "万") return Math.round(number);
    if (number >= 50 && Number.isInteger(number)) return number;

    return null;
  }

  function parseYieldToken(token) {
    let clean = normalizeText(token).replace("%", "").replace("行权", "");
    clean = clean.replace(/^(收益率|收益|ytm)[:：]?/i, "");

    if (!/^\d+(?:\.\d+)?$/.test(clean)) return null;

    const value = Number(clean);
    if (value > 0 && value < 20) return clean;

    return null;
  }

  function parseNetPriceFromText(text) {
    const patterns = [/净价[:：]?\s*(\d{2,3}(?:\.\d+)?)/i, /(\d{2,3}(?:\.\d+)?)\s*净价/i];

    for (const pattern of patterns) {
      const match = pattern.exec(text);
      if (!match) continue;
      const price = match[1];
      const value = Number(price);
      if (value >= 50 && value <= 150) {
        return { price, fragment: match[0] };
      }
    }

    return { price: "", fragment: "" };
  }

  function extractIntermediary(text) {
    const match = /【([^】]+)】/.exec(text);
    if (!match) return { intermediary: "", text };

    return {
      intermediary: match[1].trim(),
      text: removeOnce(text, match[0]),
    };
  }

  function extractBondCodeAndParentheticalDate(text, year) {
    const codePattern = /(?:^|\s)(\d{6,9}(?:\.(?:IB|SH|SZ))?)(?:\(([^)]*)\))?/i;
    const match = codePattern.exec(text);
    if (!match) {
      return { code: "", parentheticalDate: null, text };
    }

    const rawCode = match[1];
    let code = rawCode.toUpperCase();
    if (!code.includes(".")) code = `${code}.IB`;

    let parentheticalDate = null;
    const inside = normalizeText(match[2] || "");
    if (inside) {
      parentheticalDate = parseDateText(inside.split(" ")[0], year);
    }

    return {
      code,
      parentheticalDate,
      text: removeOnce(text, match[0]),
    };
  }

  function extractTradeDateAndSpeed(text, year) {
    const dateExpr =
      "(?:\\d{4}-\\d{1,2}-\\d{1,2}|\\d{1,2}[./]\\d{1,2}|\\d{1,2}月\\d{1,2}日?|\\d{3,4})";
    const serialExpr = "(?:\\d+[)、]\\s*)?";

    const speedPattern = new RegExp(`${serialExpr}(${dateExpr})\\s*\\+\\s*([01])`);
    let match = speedPattern.exec(text);
    if (match) {
      const parsed = parseDateText(match[1], year);
      if (parsed) {
        return {
          tradeDate: parsed,
          speed: match[2],
          text: removeOnce(text, match[0]),
        };
      }
    }

    const exchangePattern = new RegExp(`${serialExpr}(${dateExpr})\\s*(?:交易所|现券交易)`);
    match = exchangePattern.exec(text);
    if (match) {
      const parsed = parseDateText(match[1], year);
      if (parsed) {
        return {
          tradeDate: parsed,
          speed: "1",
          text: removeOnce(text, match[0]),
        };
      }
    }

    return { tradeDate: null, speed: "", text };
  }

  function cleanPartyName(name) {
    let value = normalizeText(name);
    value = value.replace(/^(出给|to)\s*/i, "");
    value = value.replace(/\s*(出给|to)$/i, "");
    return value.trim();
  }

  function firstPartyAfterOperator(text, bankName) {
    const value = normalizeText(text);
    if (!value) return "";
    if (value.startsWith(bankName)) return bankName;
    return cleanPartyName(value.split(" ")[0]);
  }

  function lastPartyBeforeOperator(text, bankName) {
    const value = normalizeText(text);
    if (!value) return "";
    if (value.endsWith(bankName)) return bankName;
    const parts = value.split(" ");
    return cleanPartyName(parts[parts.length - 1]);
  }

  function extractDirection(text, bankName) {
    const normalizedBankName = normalizeText(bankName);
    const operators = [
      { label: "出给", pattern: /出给/ },
      { label: "to", pattern: /to/i },
    ];

    for (const operator of operators) {
      const match = operator.pattern.exec(text);
      if (!match) continue;

      const leftText = text.slice(0, match.index).trim();
      const rightText = text.slice(match.index + match[0].length).trim();
      const leftParty = lastPartyBeforeOperator(leftText, normalizedBankName);
      const rightParty = firstPartyAfterOperator(rightText, normalizedBankName);

      let direction = "";
      let counterparty = "";

      if (leftParty === normalizedBankName) {
        direction = "卖出";
        counterparty = rightParty;
      } else if (rightParty === normalizedBankName) {
        direction = "买入";
        counterparty = leftParty;
      }

      if (direction) {
        let cleanedText = text;
        for (const fragment of [match[0], leftParty, rightParty]) {
          cleanedText = removeOnce(cleanedText, fragment);
        }
        return { direction, counterparty, text: cleanedText };
      }

      return { direction: "", counterparty: "", text: removeOnce(text, operator.label) };
    }

    return { direction: "", counterparty: "", text };
  }

  function removeRemainingTerm(text) {
    return normalizeText(text.replace(/\b\d+(?:\.\d+)?\s*[dDyY](?:\([^)]*\))?/g, " "));
  }

  function cleanResidualText(text) {
    let value = removeRemainingTerm(text);
    value = value.replace(/\b\d+[)、]\s*/g, " ");
    value = value.replace(/(^|\s)(买入|卖出)(?=\s|$)/g, " ");
    return normalizeText(value);
  }

  function extractYieldAndFace(text) {
    const tokens = normalizeText(text).split(" ").filter(Boolean);
    let yieldValue = "";
    let faceValue = "";
    const usedIndexes = new Set();

    for (const [index, token] of tokens.entries()) {
      const value = parseYieldToken(token);
      if (!value) continue;

      const hasYieldHint =
        token.includes(".") ||
        token.includes("%") ||
        token.includes("行权") ||
        token.includes("收益") ||
        token.toLowerCase().includes("ytm");

      if (hasYieldHint) {
        yieldValue = value;
        usedIndexes.add(index);
        break;
      }
    }

    if (!yieldValue) {
      for (const [index, token] of tokens.entries()) {
        const value = parseYieldToken(token);
        if (value && parseFaceValue(token) === null) {
          yieldValue = value;
          usedIndexes.add(index);
          break;
        }
      }
    }

    for (const [index, token] of tokens.entries()) {
      if (usedIndexes.has(index)) continue;
      const value = parseFaceValue(token);
      if (value !== null) {
        faceValue = String(value);
        usedIndexes.add(index);
        break;
      }
    }

    const remaining = tokens.filter((_, index) => !usedIndexes.has(index)).join(" ");
    return { yieldValue, faceValue, remaining };
  }

  function blankTrade(negotiationDate) {
    const trade = {};
    for (const column of EXCEL_COLUMNS) trade[column] = "";
    trade["谈判日"] = formatDate(negotiationDate);
    return trade;
  }

  function parseTradeLine(rawLine, negotiationDate, bankName) {
    let text = normalizeText(rawLine);
    const year = negotiationDate.getFullYear();
    const trade = blankTrade(negotiationDate);
    const warnings = [];

    const intermediaryResult = extractIntermediary(text);
    trade["中介"] = intermediaryResult.intermediary;
    text = intermediaryResult.text;

    const directionResult = extractDirection(text, bankName);
    trade["我行方向"] = directionResult.direction;
    trade["真实交易对手"] = directionResult.counterparty;
    text = directionResult.text;

    const dateResult = extractTradeDateAndSpeed(text, year);
    let tradeDate = dateResult.tradeDate;
    trade["清算速度(0/1)"] = dateResult.speed;
    text = dateResult.text;

    const codeResult = extractBondCodeAndParentheticalDate(text, year);
    trade["债券代码"] = codeResult.code;
    text = codeResult.text;

    if (!tradeDate && codeResult.parentheticalDate) {
      tradeDate = codeResult.parentheticalDate;
    }

    if (tradeDate) trade["交易日"] = formatDate(tradeDate);

    const netPriceResult = parseNetPriceFromText(text);
    if (netPriceResult.price) {
      trade["净价"] = netPriceResult.price;
      text = removeOnce(text, netPriceResult.fragment);
    }

    text = cleanResidualText(text);
    const amountResult = extractYieldAndFace(text);
    trade["收益率(%)"] = amountResult.yieldValue;
    trade["面值（万元）"] = amountResult.faceValue;

    if (!trade["中介"]) warnings.push("未识别中介");
    if (!trade["债券代码"]) warnings.push("未识别债券代码");
    if (!trade["交易日"]) warnings.push("未识别交易日");
    if (!trade["清算速度(0/1)"]) warnings.push("未识别清算速度");
    if (!trade["收益率(%)"] && !trade["净价"]) warnings.push("未识别收益率或净价");
    if (!trade["面值（万元）"]) warnings.push("未识别面值");
    if (!trade["我行方向"] || !trade["真实交易对手"]) warnings.push("未识别方向或真实交易对手");

    return { trade, warnings };
  }

  function parseTradeText(rawText, negotiationDate, bankName) {
    const trades = [];
    const diagnostics = [];
    const lines = String(rawText || "").split(/\r?\n/);

    lines.forEach((rawLine, index) => {
      const line = normalizeText(rawLine);
      if (!line) return;

      const result = parseTradeLine(line, negotiationDate, bankName);
      trades.push(result.trade);

      if (result.warnings.length) {
        diagnostics.push({
          lineNumber: index + 1,
          original: line,
          message: result.warnings.join("；"),
        });
      }
    });

    return { trades, diagnostics };
  }

  function coerceWorksheetValue(value, columnName) {
    if (value === null || value === undefined || value === "") return "";

    if (columnName === "谈判日" || columnName === "交易日") {
      const parsed = parseInputDate(value);
      return parsed;
    }

    if (["净价", "收益率(%)", "成本", "价差"].includes(columnName)) {
      const numeric = Number(value);
      return Number.isFinite(numeric) ? numeric : value;
    }

    if (columnName === "面值（万元）" || columnName === "清算速度(0/1)") {
      const numeric = Number(value);
      return Number.isFinite(numeric) ? Math.trunc(numeric) : value;
    }

    return value;
  }

  function formulaForColumn(columnName, rowNumber) {
    if (columnName === "债券简称") return `@B_INFO_NAME(C${rowNumber})`;
    if (columnName === "债券类型") return `B_INFO_WINDL2TYPE(C${rowNumber})`;
    if (columnName === "估值收益率") {
      return `@IF(@B_ANAL_YIELD_CNBD(C${rowNumber},A${rowNumber}-1,1)=0,"-",B_ANAL_YIELD_CNBD(C${rowNumber},A${rowNumber}-1,1))`;
    }
    return "";
  }

  function createWorkbook(trades) {
    if (typeof XLSX === "undefined") {
      throw new Error("Excel 导出库没有加载完成，请刷新页面后重试。");
    }

    const aoa = [
      EXCEL_COLUMNS,
      ...trades.map((trade) =>
        EXCEL_COLUMNS.map((columnName) => coerceWorksheetValue(trade[columnName], columnName)),
      ),
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });

    for (let rowIndex = 1; rowIndex <= trades.length; rowIndex += 1) {
      const excelRow = rowIndex + 1;
      EXCEL_COLUMNS.forEach((columnName, columnIndex) => {
        const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
        const cell = worksheet[address] || {};

        if (FORMULA_COLUMNS.has(columnName)) {
          cell.f = formulaForColumn(columnName, excelRow);
          cell.t = "s";
          cell.v = "";
          worksheet[address] = cell;
        }

        if (COLUMN_FORMATS[columnName]) {
          worksheet[address] = worksheet[address] || {};
          worksheet[address].z = COLUMN_FORMATS[columnName];
        }
      });
    }

    worksheet["!cols"] = EXCEL_COLUMNS.map((columnName) => ({
      wch: COLUMN_WIDTHS[columnName] || 14,
    }));

    if (worksheet["!ref"]) {
      worksheet["!autofilter"] = { ref: worksheet["!ref"] };
    }

    const workbook = XLSX.utils.book_new();
    workbook.Workbook = { CalcPr: { fullCalcOnLoad: "1" } };
    XLSX.utils.book_append_sheet(workbook, worksheet, "Trade Records");
    return workbook;
  }

  function setMode(mode) {
    state.activeMode = mode;
    $("#pasteModeButton").classList.toggle("active", mode === "paste");
    $("#wordModeButton").classList.toggle("active", mode === "word");
    $("#pasteModeButton").setAttribute("aria-selected", String(mode === "paste"));
    $("#wordModeButton").setAttribute("aria-selected", String(mode === "word"));
    $("#pastePane").classList.toggle("active", mode === "paste");
    $("#wordPane").classList.toggle("active", mode === "word");
  }

  function activeRawText() {
    return state.activeMode === "paste" ? $("#rawText").value : $("#wordText").value;
  }

  function beginBankNameEdit() {
    const editor = $("#bankNameEditor");
    const input = $("#bankName");
    if (!input.readOnly) return;

    state.bankNameBeforeEdit = input.value;
    input.readOnly = false;
    editor.classList.add("editing");
    input.focus();
    input.select();
  }

  function finishBankNameEdit({ cancel = false } = {}) {
    const editor = $("#bankNameEditor");
    const input = $("#bankName");
    if (input.readOnly) return;

    if (cancel) {
      input.value = state.bankNameBeforeEdit;
    } else if (!normalizeText(input.value)) {
      input.value = state.bankNameBeforeEdit || "兴业银行";
    } else {
      input.value = normalizeText(input.value);
    }

    input.readOnly = true;
    editor.classList.remove("editing");
  }

  function applyColumnWidths() {
    const table = $("#resultTable");
    const columns = table.querySelectorAll("col[data-column]");
    let totalWidth = 0;

    columns.forEach((column) => {
      const columnName = column.dataset.column;
      const width = state.columnWidths[columnName] || 110;
      column.style.width = `${width}px`;
      totalWidth += width;
    });

    table.style.width = `${totalWidth}px`;
  }

  function startColumnResize(event) {
    event.preventDefault();
    event.stopPropagation();

    const handle = event.currentTarget;
    const columnName = handle.dataset.column;
    if (!columnName) return;

    const startX = event.clientX;
    const startWidth = state.columnWidths[columnName] || 110;
    document.body.classList.add("resizing-column");

    const onPointerMove = (moveEvent) => {
      const nextWidth = Math.max(72, Math.min(520, startWidth + moveEvent.clientX - startX));
      state.columnWidths[columnName] = nextWidth;
      applyColumnWidths();
    };

    const onPointerUp = () => {
      document.body.classList.remove("resizing-column");
      window.removeEventListener("pointermove", onPointerMove);
      window.removeEventListener("pointerup", onPointerUp);
      window.removeEventListener("pointercancel", onPointerUp);
    };

    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
    window.addEventListener("pointercancel", onPointerUp);
  }

  function renderTable() {
    const table = $("#resultTable");
    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody");
    let colgroup = table.querySelector("colgroup");
    if (!colgroup) {
      colgroup = document.createElement("colgroup");
      table.insertBefore(colgroup, thead);
    }

    colgroup.replaceChildren();
    thead.replaceChildren();
    tbody.replaceChildren();

    const headerRow = document.createElement("tr");
    EXCEL_COLUMNS.forEach((columnName) => {
      const col = document.createElement("col");
      col.dataset.column = columnName;
      colgroup.appendChild(col);

      const th = document.createElement("th");
      th.className = "resizable-header";
      th.title = `${columnName}：拖动右侧边缘调整宽度`;

      const label = document.createElement("span");
      label.className = "header-label";
      label.textContent = columnName;

      const resizer = document.createElement("button");
      resizer.className = "column-resizer";
      resizer.type = "button";
      resizer.dataset.column = columnName;
      resizer.setAttribute("aria-label", `调整${columnName}列宽`);
      resizer.title = `拖动调整${columnName}列宽`;
      resizer.addEventListener("pointerdown", startColumnResize);

      th.append(label, resizer);
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    state.trades.forEach((trade, rowIndex) => {
      const tr = document.createElement("tr");
      EXCEL_COLUMNS.forEach((columnName) => {
        const td = document.createElement("td");
        const input = document.createElement("input");
        input.value = trade[columnName] || "";
        input.dataset.row = String(rowIndex);
        input.dataset.column = columnName;
        input.title = input.value;

        if (FORMULA_COLUMNS.has(columnName)) {
          input.readOnly = true;
          input.placeholder = "公式";
          input.title = "导出 Excel 后由公式生成";
        }

        td.appendChild(input);
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    applyColumnWidths();
    $("#emptyState").classList.toggle("hidden", state.trades.length > 0);
  }

  function renderDiagnostics() {
    const panel = $("#diagnosticPanel");
    const list = $("#diagnosticList");
    list.replaceChildren();

    if (!state.diagnostics.length) {
      panel.hidden = true;
      return;
    }

    for (const item of state.diagnostics) {
      const row = document.createElement("div");
      row.className = "diagnostic-item";
      row.innerHTML = `<strong>第 ${item.lineNumber} 行：</strong>${escapeHtml(item.message)}<br>${escapeHtml(item.original)}`;
      list.appendChild(row);
    }

    panel.hidden = false;
  }

  function updateSummary() {
    const count = state.trades.length;
    const warningCount = state.diagnostics.length;
    if (!count) {
      $("#summaryText").textContent = "等待输入";
    } else if (warningCount) {
      $("#summaryText").textContent = `已转换 ${count} 行，${warningCount} 行需要复核`;
    } else {
      $("#summaryText").textContent = `已转换 ${count} 行`;
    }

    $("#downloadButton").disabled = count === 0;
  }

  function renderAll() {
    renderTable();
    renderDiagnostics();
    updateSummary();
    if (window.lucide) window.lucide.createIcons();
  }

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  async function readWordFile(file) {
    if (typeof mammoth === "undefined") {
      throw new Error("Word 读取库没有加载完成，请刷新页面后重试。");
    }

    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer });
    return result.value || "";
  }

  function parseCurrentInput() {
    const rawText = activeRawText();
    if (!normalizeText(rawText)) {
      state.trades = [];
      state.diagnostics = [
        {
          lineNumber: 0,
          original: "",
          message: "请先粘贴交易记录或上传 Word 文档",
        },
      ];
      renderAll();
      return;
    }

    const negotiationDate = parseInputDate($("#negotiationDate").value);
    const bankName = $("#bankName").value || "兴业银行";
    const result = parseTradeText(rawText, negotiationDate, bankName);
    state.trades = result.trades;
    state.diagnostics = result.diagnostics;
    renderAll();
  }

  function downloadExcel() {
    try {
      const workbook = createWorkbook(state.trades);
      const fileName = `Trade_Records_${formatDate(new Date()).replace(/-/g, "")}.xlsx`;
      XLSX.writeFile(workbook, fileName, { bookType: "xlsx", cellDates: true });
    } catch (error) {
      state.diagnostics = [
        {
          lineNumber: 0,
          original: "",
          message: error.message || "导出失败",
        },
      ];
      renderDiagnostics();
      updateSummary();
    }
  }

  function clearAll() {
    $("#rawText").value = "";
    $("#wordText").value = "";
    $("#wordFile").value = "";
    $("#fileName").textContent = "尚未选择文件";
    state.trades = [];
    state.diagnostics = [];
    renderAll();
  }

  function initDom() {
    $("#negotiationDate").value = formatDate(new Date());

    $("#editBankNameButton").addEventListener("pointerdown", (event) => {
      event.preventDefault();
    });
    $("#editBankNameButton").addEventListener("click", beginBankNameEdit);
    $("#bankName").addEventListener("dblclick", beginBankNameEdit);
    $("#bankName").addEventListener("blur", () => finishBankNameEdit());
    $("#bankName").addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        finishBankNameEdit();
        $("#bankName").blur();
      } else if (event.key === "Escape") {
        event.preventDefault();
        finishBankNameEdit({ cancel: true });
        $("#bankName").blur();
      }
    });

    $("#pasteModeButton").addEventListener("click", () => setMode("paste"));
    $("#wordModeButton").addEventListener("click", () => setMode("word"));
    $("#loadExampleButton").addEventListener("click", () => {
      setMode("paste");
      $("#rawText").value = SAMPLE_TEXT;
      parseCurrentInput();
    });
    $("#parseButton").addEventListener("click", parseCurrentInput);
    $("#clearButton").addEventListener("click", clearAll);
    $("#downloadButton").addEventListener("click", downloadExcel);

    $("#wordFile").addEventListener("change", async (event) => {
      const file = event.target.files && event.target.files[0];
      if (!file) return;
      $("#fileName").textContent = file.name;

      try {
        $("#wordText").value = await readWordFile(file);
      } catch (error) {
        state.diagnostics = [
          {
            lineNumber: 0,
            original: file.name,
            message: error.message || "Word 文件读取失败",
          },
        ];
        renderDiagnostics();
      }
    });

    $("#resultTable").addEventListener("input", (event) => {
      const input = event.target;
      if (!(input instanceof HTMLInputElement)) return;

      const row = Number(input.dataset.row);
      const column = input.dataset.column;
      if (!Number.isInteger(row) || !column || !state.trades[row]) return;
      state.trades[row][column] = input.value;
      input.title = input.value;
    });

    renderAll();
  }

  const core = {
    EXCEL_COLUMNS,
    parseTradeText,
    parseTradeLine,
    normalizeText,
    createWorkbook,
  };

  if (typeof window !== "undefined") {
    window.TradeParserCore = core;
    document.addEventListener("DOMContentLoaded", initDom);
  }

  if (typeof module !== "undefined") {
    module.exports = core;
  }
})();
