"""把 Excel 工作表导入 SQLite，供桌面窗口查询。"""

from __future__ import annotations

import math
import re
import sqlite3
import sys
import threading
from datetime import date, datetime, time
from pathlib import Path

import pandas as pd


def _runtime_paths() -> tuple[Path, Path, Path]:
    """源码运行时资源和数据都在项目目录。打包后页面在解压目录，数据库放在 exe 旁边。"""
    if getattr(sys, "frozen", False):
        root = Path(sys.executable).resolve().parent
        static = Path(getattr(sys, "_MEIPASS", root)) / "static"
        return root, static, root / "data"
    root = Path(__file__).resolve().parent
    return root, root / "static", root / "data"


ROOT, STATIC_DIR, DATA_DIR = _runtime_paths()
DB_PATH = DATA_DIR / "workspace.sqlite"

MAX_RESULT_ROWS = 2000
ALLOWED_EXTS = {".xlsx", ".xlsm", ".xls"}
ALLOWED_SQL = {
    "select",
    "with",
    "explain",
    "pragma",
    "insert",
    "update",
    "delete",
    "replace",
    "create",
    "drop",
    "alter",
    "values",
}
BLOCKED_SQL = re.compile(r"(?is)\b(attach|detach|load_extension|vacuum\s+into)\b")

_lock = threading.RLock()
_conn: sqlite3.Connection | None = None


class AppError(Exception):
    """可以展示给使用者的错误。"""


def quote_ident(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def is_empty(value) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def header_text(value) -> str:
    if is_empty(value):
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer() and abs(value) < 2**53:
        return str(int(value))
    if isinstance(value, int) and not isinstance(value, bool):
        return str(value)
    text = re.sub(r"\s+", " ", str(value).strip())
    if re.fullmatch(r"(?i)unnamed:\s*\d+", text):
        return ""
    return text


def dedupe_names(names: list[str]) -> list[str]:
    """同名列里第一个保留原名，后面的重复项用下一个空着的 _2、_3……"""
    result: list[str | None] = [None] * len(names)
    used: set[str] = set()
    pending: list[int] = []
    for index, name in enumerate(names):
        if name not in used:
            result[index] = name
            used.add(name)
        else:
            pending.append(index)
    for index in pending:
        base = names[index]
        number = 2
        candidate = f"{base}_{number}"
        while candidate in used:
            number += 1
            candidate = f"{base}_{number}"
        used.add(candidate)
        result[index] = candidate
    return result


def normalize_headers(raw_headers: list) -> list[str]:
    cleaned: list[str] = []
    empty_index = 0
    for raw in raw_headers:
        text = header_text(raw)
        if not text:
            empty_index += 1
            cleaned.append(f"列{empty_index}")
        else:
            cleaned.append(text)
    return dedupe_names(cleaned)


def cell_value(value):
    if hasattr(value, "item") and not isinstance(value, (str, bytes, datetime, date, time)):
        try:
            value = value.item()
        except (ValueError, AttributeError):
            pass
    if is_empty(value):
        return None
    if isinstance(value, bool):
        return 1 if value else 0
    if isinstance(value, datetime):
        if value.hour or value.minute or value.second or value.microsecond:
            if value.microsecond:
                text = value.strftime("%Y-%m-%d %H:%M:%S.%f").rstrip("0")
                return text[:-1] if text.endswith(".") else text
            return value.strftime("%Y-%m-%d %H:%M:%S")
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, time):
        return value.strftime("%H:%M:%S")
    if isinstance(value, int) and not isinstance(value, bool):
        return int(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            return None
        if value.is_integer() and abs(value) < 2**53:
            return int(value)
        return float(value)
    if isinstance(value, str):
        return value
    return str(value)


def unique_table_name(file_stem: str, sheet_name: str, taken: set[str]) -> str:
    stem = re.sub(r"\s+", " ", file_stem).strip() or "文件"
    sheet = re.sub(r"\s+", " ", sheet_name).strip() or "Sheet"
    base = f"{stem}__{sheet}".replace('"', "'")
    name = base
    number = 2
    while name.lower() in taken:
        name = f"{base}_{number}"
        number += 1
    taken.add(name.lower())
    return name


def engines_for(path: Path) -> list[str]:
    if path.suffix.lower() == ".xls":
        return ["xlrd", "openpyxl"]
    return ["openpyxl", "xlrd"]


def read_sheets(path: Path) -> list[dict]:
    errors: list[str] = []
    workbook = None
    for engine in engines_for(path):
        try:
            workbook = pd.ExcelFile(path, engine=engine)
            break
        except Exception as exc:  # noqa: BLE001 - 需要把引擎失败汇总成一条可读错误
            errors.append(f"{engine}: {exc}")
    if workbook is None:
        detail = "；".join(errors) if errors else "未知错误"
        raise ValueError(f"无法读取（文件可能已加密或已损坏）：{detail}")

    prepared: list[dict] = []
    try:
        for sheet_name in workbook.sheet_names:
            frame = pd.read_excel(workbook, sheet_name=sheet_name, header=None, dtype=object)
            prepared.append(prepare_sheet(sheet_name, frame))
    finally:
        workbook.close()
    return prepared


def prepare_sheet(sheet_name: str, frame: pd.DataFrame) -> dict:
    if frame.empty:
        return {"sheet": sheet_name, "skip": "空表", "columns": [], "rows": []}

    matrix = frame.values.tolist()
    width = max((len(row) for row in matrix), default=0)
    rows = [list(row) + [None] * (width - len(row)) for row in matrix]
    rows = [row for row in rows if not all(is_empty(cell) for cell in row)]
    if not rows:
        return {"sheet": sheet_name, "skip": "空表", "columns": [], "rows": []}

    header_row = rows[0]
    data_rows = rows[1:]
    keep: list[int] = []
    for index in range(width):
        has_header = not is_empty(header_row[index]) and header_text(header_row[index]) != ""
        has_data = any(not is_empty(row[index]) for row in data_rows)
        if has_header or has_data:
            keep.append(index)
    if not keep:
        return {"sheet": sheet_name, "skip": "空表", "columns": [], "rows": []}

    columns = normalize_headers([header_row[index] for index in keep])
    values = [
        tuple(cell_value(row[index]) for index in keep)
        for row in data_rows
        if any(not is_empty(row[index]) for index in keep)
    ]
    return {"sheet": sheet_name, "skip": None, "columns": columns, "rows": values}


def get_conn() -> sqlite3.Connection:
    global _conn
    if _conn is None:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        _conn = sqlite3.connect(DB_PATH, check_same_thread=False, timeout=30)
        _conn.row_factory = sqlite3.Row
        _conn.isolation_level = None
        _conn.execute("PRAGMA journal_mode=WAL")
        _conn.execute(
            """
            CREATE TABLE IF NOT EXISTS _meta_imports (
                table_name TEXT PRIMARY KEY,
                file_name TEXT NOT NULL,
                sheet_name TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )
        _conn.commit()
    return _conn


def existing_table_keys(conn: sqlite3.Connection) -> set[str]:
    rows = conn.execute(
        """
        SELECT name FROM sqlite_master
        WHERE type = 'table' AND name NOT LIKE 'sqlite_%'
        """
    ).fetchall()
    return {row["name"].lower() for row in rows}


def import_files(saved: list[tuple[str, Path]]) -> dict:
    parsed: list[tuple[str, list[dict]]] = []
    failures: list[dict] = []
    for file_name, path in saved:
        try:
            parsed.append((file_name, read_sheets(path)))
        except Exception as exc:  # noqa: BLE001 - 单个文件失败不影响其他文件
            failures.append({"file": file_name, "sheet": "", "reason": str(exc)})

    imported: list[dict] = []
    skipped: list[dict] = list(failures)
    with _lock:
        conn = get_conn()
        taken = existing_table_keys(conn)
        try:
            conn.execute("BEGIN")
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for file_name, sheets in parsed:
                stem = Path(file_name).stem
                for sheet in sheets:
                    if sheet["skip"]:
                        skipped.append(
                            {"file": file_name, "sheet": sheet["sheet"], "reason": sheet["skip"]}
                        )
                        continue
                    table_name = unique_table_name(stem, sheet["sheet"], taken)
                    column_sql = ", ".join(quote_ident(column) for column in sheet["columns"])
                    conn.execute(f"CREATE TABLE {quote_ident(table_name)} ({column_sql})")
                    if sheet["rows"]:
                        placeholders = ", ".join(["?"] * len(sheet["columns"]))
                        conn.executemany(
                            f"INSERT INTO {quote_ident(table_name)} ({column_sql}) VALUES ({placeholders})",
                            sheet["rows"],
                        )
                    conn.execute(
                        """
                        INSERT OR REPLACE INTO _meta_imports (table_name, file_name, sheet_name, created_at)
                        VALUES (?, ?, ?, ?)
                        """,
                        (table_name, file_name, sheet["sheet"], now),
                    )
                    imported.append(
                        {
                            "table": table_name,
                            "file": file_name,
                            "sheet": sheet["sheet"],
                            "rows": len(sheet["rows"]),
                            "columns": sheet["columns"],
                        }
                    )
            conn.commit()
        except Exception:
            if conn.in_transaction:
                conn.rollback()
            raise

    return {"imported": imported, "skipped": skipped, "tables": list_tables()}


def import_paths(paths: list[str]) -> dict:
    usable: list[tuple[str, Path]] = []
    failures: list[dict] = []
    for raw in paths:
        path = Path(raw)
        if path.suffix.lower() not in ALLOWED_EXTS:
            failures.append({"file": path.name, "sheet": "", "reason": "不支持的格式"})
            continue
        if not path.is_file() or path.stat().st_size == 0:
            failures.append({"file": path.name, "sheet": "", "reason": "文件为空或不存在"})
            continue
        usable.append((path.name, path))
    if not usable:
        reason = failures[0]["reason"] if failures else "请选择 Excel 文件"
        name = failures[0]["file"] if failures else ""
        raise AppError(f"{reason}：{name}".rstrip("："))
    result = import_files(usable)
    result["skipped"] = failures + result["skipped"]
    return result


def list_tables() -> list[dict]:
    with _lock:
        conn = get_conn()
        conn.execute(
            """
            DELETE FROM _meta_imports
            WHERE table_name NOT IN (SELECT name FROM sqlite_master WHERE type = 'table')
            """
        )
        ordered = conn.execute(
            """
            SELECT m.table_name AS name, m.file_name, m.sheet_name
            FROM _meta_imports AS m
            JOIN sqlite_master AS s ON s.name = m.table_name AND s.type = 'table'
            ORDER BY m.rowid
            """
        ).fetchall()
        known = {row["name"] for row in ordered}
        extras = conn.execute(
            """
            SELECT name FROM sqlite_master
            WHERE type = 'table'
              AND name NOT LIKE 'sqlite_%'
              AND name NOT LIKE '\\_meta%' ESCAPE '\\'
            ORDER BY name
            """
        ).fetchall()
        meta = {row["name"]: row for row in ordered}
        names = [{"name": row["name"]} for row in ordered]
        names.extend({"name": row["name"]} for row in extras if row["name"] not in known)
        tables: list[dict] = []
        for item in names:
            name = item["name"]
            info = meta.get(name)
            columns = [
                row["name"]
                for row in conn.execute(f"PRAGMA table_info({quote_ident(name)})").fetchall()
            ]
            count = conn.execute(f"SELECT COUNT(*) AS n FROM {quote_ident(name)}").fetchone()["n"]
            tables.append(
                {
                    "name": name,
                    "file": info["file_name"] if info else "",
                    "sheet": info["sheet_name"] if info else name,
                    "rows": count,
                    "columns": columns,
                }
            )
        return tables


def strip_sql_comments(sql: str) -> str:
    without_block = re.sub(r"/\*.*?\*/", " ", sql, flags=re.S)
    return re.sub(r"--[^\n]*", " ", without_block)


def json_safe(value):
    if isinstance(value, float) and not math.isfinite(value):
        return None
    if isinstance(value, (bytes, memoryview)):
        return bytes(value).hex()
    return value


def run_sql(sql: str) -> dict:
    cleaned = strip_sql_comments(sql).strip().rstrip(";").strip()
    if not cleaned:
        raise AppError("请输入 SQL")
    if BLOCKED_SQL.search(cleaned):
        raise AppError("不支持 ATTACH、DETACH 或把数据库导出到其他文件")
    match = re.match(r"([A-Za-z_]+)", cleaned)
    keyword = match.group(1).lower() if match else ""
    if keyword not in ALLOWED_SQL:
        raise AppError("这条语句不能在当前数据里执行")

    started = datetime.now()
    with _lock:
        conn = get_conn()
        try:
            cursor = conn.execute(cleaned)
            if cursor.description is None:
                conn.commit()
                elapsed = (datetime.now() - started).total_seconds() * 1000
                affected = cursor.rowcount if cursor.rowcount is not None and cursor.rowcount >= 0 else 0
                return {
                    "columns": [],
                    "rows": [],
                    "row_count": affected,
                    "truncated": False,
                    "elapsed_ms": round(elapsed, 1),
                    "message": f"执行完成，影响 {affected} 行",
                }
            fetched = cursor.fetchmany(MAX_RESULT_ROWS + 1)
            columns = [item[0] for item in cursor.description]
        except sqlite3.Error as exc:
            if conn.in_transaction:
                conn.rollback()
            raise AppError(str(exc)) from exc

    truncated = len(fetched) > MAX_RESULT_ROWS
    visible = fetched[:MAX_RESULT_ROWS]
    elapsed = (datetime.now() - started).total_seconds() * 1000
    return {
        "columns": columns,
        "rows": [[json_safe(cell) for cell in row] for row in visible],
        "row_count": len(visible),
        "truncated": truncated,
        "elapsed_ms": round(elapsed, 1),
        "message": "",
    }


def reset():
    global _conn
    with _lock:
        if _conn is not None:
            _conn.close()
            _conn = None
        if DATA_DIR.exists():
            for path in DATA_DIR.glob("workspace.sqlite*"):
                path.unlink(missing_ok=True)
        get_conn()
    return {"ok": True, "tables": []}
