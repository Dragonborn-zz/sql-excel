"""双击打开的桌面窗口。"""

from __future__ import annotations

import json
import re
from pathlib import Path

import webview
from webview.dom import DOMEventHandler

import app

EXCEL_TYPES = ("Excel (*.xlsx;*.xlsm;*.xls)",)
CSV_TYPES = ("CSV (*.csv)",)
XLSX_TYPES = ("Excel (*.xlsx)",)


class Api:
    def schema(self):
        return {"tables": app.list_tables()}

    def query(self, sql: str):
        try:
            return app.run_sql(sql)
        except app.AppError as exc:
            return {"error": str(exc)}

    def reset_data(self):
        return app.reset()

    def choose_files(self):
        selected = webview.windows[0].create_file_dialog(
            webview.FileDialog.OPEN,
            allow_multiple=True,
            file_types=EXCEL_TYPES,
        )
        if not selected:
            return None
        try:
            return app.import_paths(list(selected))
        except app.AppError as exc:
            return {"error": str(exc)}

    def save_csv(self, content: str):
        selected = webview.windows[0].create_file_dialog(
            webview.FileDialog.SAVE,
            save_filename="query.csv",
            file_types=CSV_TYPES,
        )
        if not selected:
            return None
        path = Path(selected if isinstance(selected, str) else selected[0])
        if path.suffix.lower() != ".csv":
            path = path.with_suffix(".csv")
        path.write_text(content, encoding="utf-8-sig")
        return str(path)

    def save_excel(self, columns, rows):
        if not columns:
            return {"error": "没有可以导出的结果"}
        selected = webview.windows[0].create_file_dialog(
            webview.FileDialog.SAVE,
            save_filename="query.xlsx",
            file_types=XLSX_TYPES,
        )
        if not selected:
            return None
        path = Path(selected if isinstance(selected, str) else selected[0])
        if path.suffix.lower() != ".xlsx":
            path = path.with_suffix(".xlsx")
        try:
            write_excel(path, columns, rows or [])
        except Exception as exc:
            return {"error": f"无法保存：{exc}"}
        return str(path)


def write_excel(path: Path, columns, rows) -> None:
    from openpyxl import Workbook

    book = Workbook()
    sheet = book.active
    sheet.title = "查询结果"
    sheet.append([_excel_cell(column) for column in columns])
    for row in rows:
        values = list(row)
        if len(values) < len(columns):
            values.extend([None] * (len(columns) - len(values)))
        sheet.append([_excel_cell(value) for value in values[: len(columns)]])
    book.save(path)


def _excel_cell(value):
    if value is None or isinstance(value, (int, float, bool)):
        return value
    text = str(value)
    return re.sub(r"[\000-\010\013\014\016-\037]", "", text)


def dropped_paths(event: dict) -> list[str]:
    files = (event.get("dataTransfer") or {}).get("files") or []
    paths: list[str] = []
    for item in files:
        if not isinstance(item, dict):
            continue
        path = item.get("pywebviewFullPath") or ""
        if path:
            paths.append(path)
    return paths


def on_drop(window, event: dict):
    paths = dropped_paths(event)
    if not paths:
        window.evaluate_js("window.onImportFinished && window.onImportFinished()")
        return
    window.evaluate_js("window.onImportStart && window.onImportStart()")
    try:
        payload = app.import_paths(paths)
    except app.AppError as exc:
        payload = {"error": str(exc)}
    window.evaluate_js(f"window.onImported({json.dumps(payload, ensure_ascii=False)})")


def bind(window):
    window.dom.document.events.drop += DOMEventHandler(lambda event: on_drop(window, event), True, True)


def main():
    window = webview.create_window(
        "SQL 查 Excel",
        url=str(app.STATIC_DIR / "index.html"),
        js_api=Api(),
        width=1200,
        height=800,
        min_size=(960, 640),
        text_select=True,
        background_color="#E7E1D6",
    )
    app.DATA_DIR.mkdir(parents=True, exist_ok=True)
    webview.start(bind, (window,), private_mode=False, storage_path=str(app.DATA_DIR / "webview"))


if __name__ == "__main__":
    main()
