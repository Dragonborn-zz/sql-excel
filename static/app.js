const dropZone = document.querySelector("#drop-zone");
const tableList = document.querySelector("#table-list");
const statusEl = document.querySelector("#status");
const sqlEl = document.querySelector("#sql");
const runBtn = document.querySelector("#run-btn");
const exportBtn = document.querySelector("#export-btn");
const exportExcelBtn = document.querySelector("#export-excel-btn");
const resetBtn = document.querySelector("#reset-btn");
const queryMeta = document.querySelector("#query-meta");
const messageEl = document.querySelector("#message");
const resultWrap = document.querySelector("#result-wrap");
const overlay = document.querySelector("#overlay");
const tabBar = document.querySelector("#tab-bar");
const panels = document.querySelector("#panels");
const mainPanel = document.querySelector("#main-panel");
const mainTab = document.querySelector("#main-tab");

let lastResult = null;
const previews = new Map();
let savedCursor = { start: 0, end: 0 };
const savedSql = localStorage.getItem("sql-excel-last");
if (savedSql) sqlEl.value = savedSql;
savedCursor = { start: sqlEl.value.length, end: sqlEl.value.length };

function quoteIdent(name) {
  return `"${String(name).replaceAll('"', '""')}"`;
}

function setStatus(text, kind) {
  statusEl.textContent = text || "";
  statusEl.className = `status${kind ? ` ${kind}` : ""}`;
}

function errorText(error) {
  if (!error) return "操作失败";
  if (typeof error === "string") return error;
  return error.message || error.error || "操作失败";
}

function api() {
  if (!window.pywebview?.api) throw new Error("窗口还没准备好");
  return window.pywebview.api;
}

function renderTables(tables) {
  tableList.replaceChildren();
  if (!tables.length) {
    const empty = document.createElement("p");
    empty.className = "hint";
    empty.textContent = "还没有表。";
    tableList.append(empty);
    return;
  }

  const groups = new Map();
  for (const table of tables) {
    const key = table.file || "手动创建";
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(table);
  }

  for (const [fileName, items] of groups) {
    const group = document.createElement("section");
    group.className = "file-group";
    const title = document.createElement("h2");
    title.textContent = fileName;
    group.append(title);

    for (const table of items) {
      const card = document.createElement("article");
      card.className = "table-card";

      const top = document.createElement("div");
      top.className = "table-top";
      const names = document.createElement("span");
      const sheet = document.createElement("span");
      sheet.className = "sheet";
      sheet.textContent = table.sheet;
      const sqlName = document.createElement("button");
      sqlName.type = "button";
      sqlName.className = "sqlname";
      sqlName.textContent = table.name;
      sqlName.title = "单击插入表名";
      names.append(sheet, sqlName);
      const rows = document.createElement("span");
      rows.className = "rows";
      rows.textContent = `${table.rows} 行`;
      top.append(names, rows);
      top.title = "单击打开查询页签";
      sqlName.addEventListener("click", (event) => {
        event.stopPropagation();
        insertAtCursor(quoteIdent(table.name));
      });
      top.addEventListener("click", (event) => {
        if (event.target.closest(".sqlname")) return;
        openTableTab(table);
      });

      const cols = document.createElement("div");
      cols.className = "cols";
      for (const column of table.columns) {
        const button = document.createElement("button");
        button.type = "button";
        button.textContent = column;
        button.title = "插入列名";
        button.addEventListener("click", () => insertAtCursor(quoteIdent(column)));
        cols.append(button);
      }

      card.append(top, cols);
      group.append(card);
    }
    tableList.append(group);
  }
}

function rememberCursor() {
  savedCursor = {
    start: sqlEl.selectionStart ?? savedCursor.start,
    end: sqlEl.selectionEnd ?? savedCursor.end,
  };
}

function activeEditor() {
  const id = tabBar.querySelector(".tab.active")?.dataset.tab;
  if (id && id !== "main") {
    const item = previews.get(id);
    if (item?.sqlInput) return item.sqlInput;
  }
  return sqlEl;
}

function insertAtCursor(text) {
  const input = activeEditor();
  const start = input.selectionStart ?? input.value.length;
  const end = input.selectionEnd ?? start;
  const before = input.value.slice(0, start);
  const after = input.value.slice(end);
  const pad = before && !/\s$/.test(before) ? " " : "";
  input.value = `${before}${pad}${text}${after}`;
  const cursor = (before + pad + text).length;
  if (input === sqlEl) savedCursor = { start: cursor, end: cursor };
  input.focus();
  input.setSelectionRange(cursor, cursor);
}

function activateTab(id) {
  tabBar.querySelectorAll(".tab").forEach((tab) => {
    tab.classList.toggle("active", tab.dataset.tab === id);
  });
  mainPanel.hidden = id !== "main";
  for (const [key, item] of previews) item.panel.hidden = key !== id;
}

function closePreview(id) {
  const item = previews.get(id);
  if (!item) return;
  const wasActive = item.tab.classList.contains("active");
  item.tab.remove();
  item.panel.remove();
  previews.delete(id);
  if (wasActive) activateTab("main");
}

function closeAllPreviews() {
  for (const id of [...previews.keys()]) closePreview(id);
  activateTab("main");
}

function paintResult(wrap, metaEl, data) {
  wrap.replaceChildren();
  if (data.message) {
    metaEl.textContent = `${data.elapsed_ms} ms`;
    const note = document.createElement("div");
    note.className = "empty";
    note.textContent = data.message;
    wrap.append(note);
    return;
  }

  const truncated = data.truncated ? `，仅显示前 ${data.row_count} 行` : "";
  metaEl.textContent = `${data.elapsed_ms} ms · ${data.row_count} 行${truncated}`;
  if (!data.columns.length) {
    const note = document.createElement("div");
    note.className = "empty";
    note.textContent = "查询没有返回列。";
    wrap.append(note);
    return;
  }

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  const indexHead = document.createElement("th");
  indexHead.className = "index-col";
  indexHead.textContent = "#";
  headRow.append(indexHead);
  for (const column of data.columns) {
    const th = document.createElement("th");
    th.textContent = column;
    th.title = column;
    headRow.append(th);
  }
  thead.append(headRow);

  const tbody = document.createElement("tbody");
  data.rows.forEach((row, index) => {
    const tr = document.createElement("tr");
    const indexCell = document.createElement("td");
    indexCell.className = "index-col";
    indexCell.textContent = String(index + 1);
    tr.append(indexCell);
    row.forEach((value) => {
      const td = document.createElement("td");
      if (value === null) {
        td.className = "null";
        td.textContent = "NULL";
      } else {
        td.textContent = String(value);
        td.title = String(value);
      }
      tr.append(td);
    });
    tbody.append(tr);
  });

  if (!data.rows.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = data.columns.length + 1;
    td.className = "null";
    td.textContent = "没有匹配的行";
    tr.append(td);
    tbody.append(tr);
  }

  table.append(thead, tbody);
  wrap.append(table);
}

function openTableTab(table) {
  const id = `preview:${table.name}`;
  if (previews.has(id)) {
    activateTab(id);
    return;
  }

  const tab = document.createElement("button");
  tab.type = "button";
  tab.className = "tab";
  tab.dataset.tab = id;
  tab.title = table.name;
  const label = document.createElement("span");
  label.className = "tab-label";
  label.textContent = table.name;
  const close = document.createElement("span");
  close.className = "tab-close";
  close.textContent = "×";
  close.title = "关闭";
  close.addEventListener("click", (event) => {
    event.stopPropagation();
    closePreview(id);
  });
  tab.append(label, close);
  tab.addEventListener("click", () => activateTab(id));
  tabBar.append(tab);

  const panel = document.createElement("section");
  panel.className = "preview-panel";
  panel.hidden = true;

  const editor = document.createElement("section");
  editor.className = "editor";
  const bar = document.createElement("div");
  bar.className = "editor-bar";
  const sqlLabel = document.createElement("label");
  sqlLabel.textContent = "SQL";
  const actions = document.createElement("div");
  actions.className = "actions";
  const meta = document.createElement("span");
  meta.className = "meta";
  const csvBtn = document.createElement("button");
  csvBtn.type = "button";
  csvBtn.className = "ghost";
  csvBtn.textContent = "导出 CSV";
  csvBtn.disabled = true;
  const excelBtn = document.createElement("button");
  excelBtn.type = "button";
  excelBtn.className = "ghost";
  excelBtn.textContent = "导出 Excel";
  excelBtn.disabled = true;
  const run = document.createElement("button");
  run.type = "button";
  run.className = "primary";
  run.textContent = "运行";
  actions.append(meta, csvBtn, excelBtn, run);
  bar.append(sqlLabel, actions);
  const sqlInput = document.createElement("textarea");
  sqlInput.spellcheck = false;
  sqlInput.value = `SELECT * FROM ${quoteIdent(table.name)} LIMIT 200`;
  const hint = document.createElement("p");
  hint.className = "hint slim";
  hint.textContent = "Ctrl + Enter 运行。这里的 SQL 不会改到「查询」页。";
  editor.append(bar, sqlInput, hint);

  const results = document.createElement("section");
  results.className = "results";
  const message = document.createElement("div");
  message.className = "message";
  message.hidden = true;
  const wrap = document.createElement("div");
  wrap.className = "result-wrap";
  results.append(message, wrap);
  panel.append(editor, results);
  panels.append(panel);

  const item = { tab, panel, sqlInput, meta, message, wrap, csvBtn, excelBtn, run, lastResult: null, table };
  previews.set(id, item);
  run.addEventListener("click", () => runPreview(id));
  sqlInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
      event.preventDefault();
      runPreview(id);
    }
  });
  csvBtn.addEventListener("click", () => exportCsvFrom(item.lastResult, (text) => showPreviewError(item, text)));
  excelBtn.addEventListener("click", () => exportExcelFrom(item.lastResult, (text) => showPreviewError(item, text)));
  activateTab(id);
  sqlInput.focus();
  runPreview(id);
}

function showPreviewError(item, text) {
  item.message.hidden = false;
  item.message.textContent = text;
  item.meta.textContent = "";
}

async function runPreview(id) {
  const item = previews.get(id);
  if (!item) return;
  const sql = item.sqlInput.value.trim();
  if (!sql) {
    showPreviewError(item, "请输入 SQL");
    return;
  }
  item.run.disabled = true;
  item.message.hidden = true;
  try {
    const data = await api().query(sql);
    if (!previews.has(id)) return;
    if (data.error) {
      showPreviewError(item, data.error);
      return;
    }
    item.message.hidden = true;
    item.lastResult = data.columns.length ? data : null;
    item.csvBtn.disabled = !item.lastResult;
    item.excelBtn.disabled = !item.lastResult;
    paintResult(item.wrap, item.meta, data);
    if (data.message) await refreshSchema();
  } catch (error) {
    if (!previews.has(id)) return;
    showPreviewError(item, errorText(error));
  } finally {
    if (previews.has(id)) item.run.disabled = false;
  }
}

async function refreshSchema() {
  try {
    const data = await api().schema();
    renderTables(data.tables);
    return data.tables;
  } catch (error) {
    setStatus(errorText(error), "error");
    return [];
  }
}

async function applyImport(data) {
  if (!data || data.error) {
    setStatus(data?.error || "导入失败", "error");
    runBtn.disabled = false;
    return;
  }
  renderTables(data.tables || []);
  const skipped = data.skipped?.length
    ? `，跳过 ${data.skipped.length} 个（${data.skipped.map((item) => item.sheet || item.file).filter(Boolean).slice(0, 3).join("、")}）`
    : "";
  setStatus(`已导入 ${data.imported.length} 张表${skipped}`, "ok");
  if (data.imported.length) {
    const first = data.imported[0];
    sqlEl.value = `SELECT * FROM ${quoteIdent(first.table)} LIMIT 100`;
    await runQuery();
    return;
  }
  runBtn.disabled = false;
}

async function chooseFiles() {
  setStatus("正在选择文件…");
  try {
    const data = await api().choose_files();
    if (!data) {
      setStatus("");
      return;
    }
    await applyImport(data);
  } catch (error) {
    setStatus(errorText(error), "error");
    runBtn.disabled = false;
  }
}

window.onImportStart = () => {
  setStatus("正在导入…");
  runBtn.disabled = true;
};

window.onImportFinished = () => {
  runBtn.disabled = false;
  overlay.hidden = true;
};

window.onImported = (data) => {
  overlay.hidden = true;
  applyImport(data);
};

function renderResult(data) {
  lastResult = data.columns.length ? data : null;
  exportBtn.disabled = !lastResult;
  exportExcelBtn.disabled = !lastResult;
  messageEl.hidden = true;
  paintResult(resultWrap, queryMeta, data);
}

function showError(text) {
  messageEl.hidden = false;
  messageEl.textContent = text;
  queryMeta.textContent = "";
}

async function runQuery() {
  const sql = sqlEl.value.trim();
  if (!sql) {
    showError("请输入 SQL");
    return;
  }
  localStorage.setItem("sql-excel-last", sqlEl.value);
  runBtn.disabled = true;
  messageEl.hidden = true;
  try {
    const data = await api().query(sql);
    if (data.error) {
      showError(data.error);
      return;
    }
    renderResult(data);
    if (data.message) await refreshSchema();
  } catch (error) {
    showError(errorText(error));
  } finally {
    runBtn.disabled = false;
  }
}

function csvCell(value) {
  const text = value === null || value === undefined ? "" : String(value);
  if (/[",\n\r]/.test(text)) return `"${text.replaceAll('"', '""')}"`;
  return text;
}

async function exportExcelFrom(result, show) {
  if (!result) return;
  try {
    const saved = await api().save_excel(result.columns, result.rows);
    if (saved && saved.error) show(saved.error);
  } catch (error) {
    show(errorText(error));
  }
}

async function exportCsvFrom(result, show) {
  if (!result) return;
  const lines = [
    result.columns.map(csvCell).join(","),
    ...result.rows.map((row) => row.map(csvCell).join(",")),
  ];
  try {
    const saved = await api().save_csv(lines.join("\r\n"));
    if (saved && saved.error) show(saved.error);
  } catch (error) {
    show(errorText(error));
  }
}

function exportExcel() {
  return exportExcelFrom(lastResult, showError);
}

function exportCsv() {
  return exportCsvFrom(lastResult, showError);
}

dropZone.addEventListener("click", chooseFiles);
dropZone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    chooseFiles();
  }
});

let dragTimer = 0;
window.addEventListener("dragover", (event) => {
  if (![...event.dataTransfer.types].includes("Files")) return;
  event.preventDefault();
  overlay.hidden = false;
  clearTimeout(dragTimer);
  dragTimer = setTimeout(() => {
    overlay.hidden = true;
  }, 180);
});
window.addEventListener("drop", (event) => {
  if (![...event.dataTransfer.types].includes("Files")) return;
  event.preventDefault();
  clearTimeout(dragTimer);
});

sqlEl.addEventListener("keyup", rememberCursor);
sqlEl.addEventListener("click", rememberCursor);
sqlEl.addEventListener("blur", rememberCursor);
mainTab.addEventListener("click", () => activateTab("main"));

runBtn.addEventListener("click", runQuery);
exportBtn.addEventListener("click", exportCsv);
exportExcelBtn.addEventListener("click", exportExcel);
sqlEl.addEventListener("keydown", (event) => {
  if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
    event.preventDefault();
    runQuery();
  }
});

resetBtn.addEventListener("click", async () => {
  if (!window.confirm("清空已导入的全部表？")) return;
  try {
    await api().reset_data();
  } catch (error) {
    setStatus(errorText(error), "error");
    return;
  }
  closeAllPreviews();
  await refreshSchema();
  setStatus("已清空", "ok");
  resultWrap.replaceChildren();
  const empty = document.createElement("div");
  empty.className = "empty";
  empty.id = "empty";
  empty.textContent = "导入 Excel 后，在这里查看查询结果。";
  resultWrap.append(empty);
  exportBtn.disabled = true;
  exportExcelBtn.disabled = true;
  lastResult = null;
  queryMeta.textContent = "";
});

function boot() {
  refreshSchema();
}

if (window.pywebview?.api) boot();
else window.addEventListener("pywebviewready", boot);
