"""Turn a spreadsheet of subjects into one named video per row."""

import io
import json
import os
import re
import shutil
import threading
import zipfile
from dataclasses import dataclass
from pathlib import Path
from uuid import uuid4
from xml.etree import ElementTree

from loguru import logger

from app.config import config
from app.models import const
from app.models.schema import VideoParams
from app.services import task as tm
from app.utils import utils

_SPREADSHEET_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
_STATUS_NAME = "excel_batch_status.json"
_VIDEO_ID_HEADERS = {"video id", "video_id", "videoid"}
_SAFE_VIDEO_ID = re.compile(r"[A-Za-z0-9][A-Za-z0-9._-]{0,80}")

_batch_lock = threading.Lock()
_batch_thread: threading.Thread | None = None
_stop_requested = threading.Event()


@dataclass(frozen=True)
class ExcelVideoRow:
    video_id: str
    subject: str
    search_terms: list[str]


def parse_excel_rows(file_bytes: bytes) -> list[ExcelVideoRow]:
    """Read Video ID and Subject columns from the first worksheet."""
    table = _read_first_sheet(file_bytes)
    if not table:
        raise ValueError("The spreadsheet is empty.")

    header_index = _header_row_index(table)
    headers = [_normalize_header(cell) for cell in table[header_index]]
    video_id_column = _find_video_id_column(headers)
    subject_column = _find_subject_column(headers)
    terms_column = _find_terms_column(headers)
    if video_id_column is None or subject_column is None:
        raise ValueError(
            "The spreadsheet needs a Video ID column and a Subject column."
        )

    rows: list[ExcelVideoRow] = []
    seen_ids: set[str] = set()
    for raw in table[header_index + 1 :]:
        video_id = _cell(raw, video_id_column)
        subject = _cell(raw, subject_column)
        if not video_id and not subject:
            continue
        if not video_id or not subject:
            raise ValueError(
                "Every used row needs both a Video ID and a Subject. "
                f"Found Video ID {video_id!r} and Subject {subject!r}."
            )
        safe_name = video_filename(video_id)
        if safe_name.casefold() in seen_ids:
            raise ValueError(f"Video ID {video_id!r} is repeated.")
        seen_ids.add(safe_name.casefold())
        terms = _split_terms(_cell(raw, terms_column) if terms_column is not None else "")
        rows.append(ExcelVideoRow(video_id=video_id, subject=subject, search_terms=terms))

    if not rows:
        raise ValueError("The spreadsheet has headers but no video rows.")
    return rows


def video_filename(video_id: str) -> str:
    """Return a single file name such as MD-001.mp4."""
    name = str(video_id or "").strip()
    if not _SAFE_VIDEO_ID.fullmatch(name):
        raise ValueError(
            f"Video ID {video_id!r} can only use letters, numbers, dots, dashes, and underscores."
        )
    return f"{name}.mp4"


_SKIPPED_FOLDER_NAMES = {"AppData", "__pycache__", "node_modules"}


def allowed_output_roots() -> list[tuple[str, Path]]:
    """Folders the selector is allowed to open."""
    roots = [("Project storage", Path(utils.storage_dir(create=True)))]
    host_root = os.environ.get("EXCEL_HOST_FOLDER_ROOT", "").strip()
    if host_root:
        host_path = Path(host_root)
        if host_path.is_dir():
            label = os.environ.get("EXCEL_HOST_FOLDER_LABEL", "").strip() or str(host_path)
            roots.append((label, host_path))
    return roots


def folder_is_allowed(path: Path, roots: list[Path] | None = None) -> bool:
    resolved = path.resolve()
    allowed = roots if roots is not None else [root for _, root in allowed_output_roots()]
    for root in allowed:
        root_resolved = root.resolve()
        if resolved == root_resolved or root_resolved in resolved.parents:
            return True
    return False


def display_folder(path: Path) -> str:
    """Show a container path as the folder name on this PC when a mapping exists."""
    resolved = path.resolve()
    mappings: list[tuple[Path, str]] = []
    host_root = os.environ.get("EXCEL_HOST_FOLDER_ROOT", "").strip()
    host_label = os.environ.get("EXCEL_HOST_FOLDER_LABEL", "").strip()
    if host_root and host_label:
        mappings.append((Path(host_root), host_label))
    storage_label = os.environ.get("EXCEL_STORAGE_LABEL", "").strip()
    if storage_label:
        mappings.append((Path(utils.storage_dir()), storage_label))
    for root, label in mappings:
        try:
            relative = resolved.relative_to(root.resolve())
        except ValueError:
            continue
        return _join_display_path(label, relative)
    return str(resolved)


def list_child_folders(path: Path) -> list[Path]:
    children = []
    for child in path.iterdir():
        if child.name.startswith(".") or child.name in _SKIPPED_FOLDER_NAMES:
            continue
        if child.is_symlink() or not child.is_dir():
            continue
        if not folder_is_allowed(child):
            continue
        children.append(child)
    return sorted(children, key=lambda item: item.name.lower())


def _join_display_path(label: str, relative: Path) -> str:
    label = label.rstrip("\\/")
    if not relative.parts:
        return label
    windows = bool(re.match(r"^[A-Za-z]:", label)) or "\\" in label
    separator = "\\" if windows else "/"
    return label + separator + separator.join(relative.parts)


def safe_child_folder_name(name: str) -> str:
    cleaned = str(name or "").strip()
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9 ._-]{0,80}", cleaned):
        raise ValueError(
            "Use letters, numbers, spaces, dots, dashes, or underscores for the folder name."
        )
    return cleaned


def resolve_output_dir(folder: str) -> Path:
    """Create the selected folder and confirm this process can write in it."""
    text = str(folder or "").strip()
    if not text:
        raise ValueError("Choose an output folder.")
    path = Path(text)
    if not path.is_absolute():
        path = Path(utils.root_dir()) / path
    path = path.resolve()
    path.mkdir(parents=True, exist_ok=True)
    probe = path / ".excel-batch-write-test"
    probe.write_text("ok", encoding="utf-8")
    probe.unlink()
    return path


def output_path(folder: Path, video_id: str) -> Path:
    destination = (folder / video_filename(video_id)).resolve()
    if destination.parent != folder.resolve():
        raise ValueError(f"Video ID {video_id!r} is not a safe file name.")
    return destination


def status_path() -> Path:
    return Path(utils.storage_dir(create=True)) / _STATUS_NAME


def read_status() -> dict:
    path = status_path()
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if data.get("running") and not batch_is_running():
        data["running"] = False
        data["message"] = "The batch stopped before it finished."
    return data


def batch_is_running() -> bool:
    thread = _batch_thread
    return thread is not None and thread.is_alive()


def request_stop() -> None:
    _stop_requested.set()


def start_batch(
    rows: list[ExcelVideoRow],
    output_dir: Path,
    params: VideoParams,
    skip_existing: bool = True,
) -> None:
    """Generate the rows one at a time and save each file as its Video ID."""
    global _batch_thread
    with _batch_lock:
        if batch_is_running():
            raise RuntimeError("A spreadsheet batch is already running.")
        _stop_requested.clear()
        snapshot = params.model_copy(deep=True)
        snapshot.video_script = ""
        snapshot.video_count = 1
        snapshot.custom_audio_file = None
        _write_status(
            {
                "running": True,
                "total": len(rows),
                "finished": 0,
                "current": "",
                "output_dir": str(output_dir),
                "message": "Waiting to start.",
                "items": [
                    {"video_id": row.video_id, "status": "pending", "file": "", "error": ""}
                    for row in rows
                ],
            }
        )
        _batch_thread = threading.Thread(
            target=_run_batch,
            args=(list(rows), output_dir, snapshot, skip_existing),
            name="excel-video-batch",
            daemon=True,
        )
        _batch_thread.start()


def _run_batch(
    rows: list[ExcelVideoRow],
    output_dir: Path,
    params: VideoParams,
    skip_existing: bool,
) -> None:
    status = read_status()
    items = status.get("items", [])
    finished = 0
    try:
        for index, row in enumerate(rows):
            if _stop_requested.is_set():
                _mark_remaining(items, index, "stopped")
                _write_status(
                    {
                        **status,
                        "running": False,
                        "finished": finished,
                        "current": "",
                        "message": "Stopped. Finished videos were kept.",
                        "items": items,
                    }
                )
                return

            destination = output_path(output_dir, row.video_id)
            if skip_existing and destination.is_file():
                items[index] = {
                    "video_id": row.video_id,
                    "status": "skipped",
                    "file": str(destination),
                    "error": "",
                }
                finished += 1
                status = {
                    **status,
                    "running": True,
                    "finished": finished,
                    "current": "",
                    "message": f"Skipped {row.video_id} because the file already exists.",
                    "items": items,
                }
                _write_status(status)
                continue

            status = {
                **status,
                "running": True,
                "finished": finished,
                "current": row.video_id,
                "message": f"Generating {row.video_id}.",
                "items": items,
            }
            _write_status(status)
            items[index] = {
                "video_id": row.video_id,
                "status": "running",
                "file": "",
                "error": "",
            }
            try:
                saved = _generate_and_save(row, destination, params)
            except Exception as exc:
                logger.exception(f"excel batch row failed: video_id={row.video_id}")
                items[index] = {
                    "video_id": row.video_id,
                    "status": "failed",
                    "file": "",
                    "error": str(exc),
                }
                finished += 1
                status = {
                    **status,
                    "finished": finished,
                    "current": "",
                    "message": f"{row.video_id} failed. Continuing with the next row.",
                    "items": items,
                }
                _write_status(status)
                continue

            items[index] = {
                "video_id": row.video_id,
                "status": "saved",
                "file": str(saved),
                "error": "",
            }
            finished += 1
            status = {
                **status,
                "finished": finished,
                "current": "",
                "message": f"Saved {saved.name}.",
                "items": items,
            }
            _write_status(status)

        _write_status(
            {
                **status,
                "running": False,
                "finished": finished,
                "current": "",
                "message": "Finished.",
                "items": items,
            }
        )
    except Exception as exc:
        logger.exception(f"excel batch stopped unexpectedly: {exc}")
        _write_status(
            {
                **status,
                "running": False,
                "finished": finished,
                "current": "",
                "message": f"Batch failed: {exc}",
                "items": items,
            }
        )


def _generate_and_save(row: ExcelVideoRow, destination: Path, params: VideoParams) -> Path:
    task_params = params.model_copy(deep=True)
    task_params.video_subject = row.subject
    task_params.video_script = ""
    task_params.video_count = 1
    task_params.video_terms = row.search_terms or None
    task_id = str(uuid4())
    with config.runtime_config_lock():
        result = tm.start(task_id=task_id, params=task_params)
    if result.get("state") == const.TASK_STATE_FAILED or not result.get("videos"):
        raise RuntimeError(result.get("error") or "The video was not created.")
    source = Path(result["videos"][0])
    if not source.is_file():
        raise RuntimeError(f"Generated video is missing: {source}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_suffix(".mp4.partial")
    shutil.copy2(source, temporary)
    os.replace(temporary, destination)
    return destination


def _mark_remaining(items: list[dict], start: int, status_name: str) -> None:
    for index in range(start, len(items)):
        if items[index].get("status") in {"pending", "running"}:
            items[index] = {**items[index], "status": status_name}


def _write_status(payload: dict) -> None:
    path = status_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(".json.tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, path)


def _read_first_sheet(file_bytes: bytes) -> list[list[str]]:
    try:
        archive = zipfile.ZipFile(io.BytesIO(file_bytes))
    except zipfile.BadZipFile as exc:
        raise ValueError("Upload an Excel .xlsx file.") from exc
    with archive:
        names = set(archive.namelist())
        shared_strings = _shared_strings(archive, names)
        sheet_name = _first_sheet_name(archive, names)
        root = ElementTree.fromstring(archive.read(sheet_name))

    rows: list[list[str]] = []
    for row in root.findall("m:sheetData/m:row", _SPREADSHEET_NS):
        values: dict[int, str] = {}
        for cell in row.findall("m:c", _SPREADSHEET_NS):
            ref = cell.attrib.get("r", "")
            index = _column_index(ref) if ref else len(values)
            values[index] = _cell_text(cell, shared_strings)
        if not values:
            rows.append([])
            continue
        width = max(values) + 1
        rows.append([values.get(index, "") for index in range(width)])
    return rows


def _shared_strings(archive: zipfile.ZipFile, names: set[str]) -> list[str]:
    if "xl/sharedStrings.xml" not in names:
        return []
    root = ElementTree.fromstring(archive.read("xl/sharedStrings.xml"))
    strings = []
    for item in root.findall("m:si", _SPREADSHEET_NS):
        strings.append(
            "".join(node.text or "" for node in item.findall(".//m:t", _SPREADSHEET_NS))
        )
    return strings


def _first_sheet_name(archive: zipfile.ZipFile, names: set[str]) -> str:
    if "xl/workbook.xml" in names and "xl/_rels/workbook.xml.rels" in names:
        workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        first = workbook.find("m:sheets/m:sheet", _SPREADSHEET_NS)
        relation_id = ""
        if first is not None:
            relation_id = first.attrib.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
                "",
            )
        rels = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        for rel in rels:
            if rel.attrib.get("Id") == relation_id:
                target = rel.attrib.get("Target", "").lstrip("/")
                if target.startswith("xl/"):
                    return target
                return f"xl/{target}"
    if "xl/worksheets/sheet1.xml" in names:
        return "xl/worksheets/sheet1.xml"
    raise ValueError("The Excel file has no worksheet.")


def _cell_text(cell: ElementTree.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    value = cell.find("m:v", _SPREADSHEET_NS)
    if cell_type == "s" and value is not None and value.text:
        index = int(value.text)
        if 0 <= index < len(shared_strings):
            return shared_strings[index].strip()
        return ""
    inline = cell.find("m:is", _SPREADSHEET_NS)
    if cell_type == "inlineStr" and inline is not None:
        return "".join(
            node.text or "" for node in inline.findall(".//m:t", _SPREADSHEET_NS)
        ).strip()
    if value is not None and value.text:
        return value.text.strip()
    return ""


def _column_index(cell_ref: str) -> int:
    letters = "".join(character for character in cell_ref if character.isalpha())
    index = 0
    for character in letters:
        index = index * 26 + ord(character.upper()) - 64
    return index - 1


def _header_row_index(table: list[list[str]]) -> int:
    for index, row in enumerate(table[:10]):
        normalized = [_normalize_header(cell) for cell in row]
        if _find_video_id_column(normalized) is not None and _find_subject_column(
            normalized
        ) is not None:
            return index
    return 0


def _find_video_id_column(headers: list[str]) -> int | None:
    for index, header in enumerate(headers):
        if header in _VIDEO_ID_HEADERS:
            return index
    return None


def _find_subject_column(headers: list[str]) -> int | None:
    for index, header in enumerate(headers):
        if header == "subject" or header.startswith("subject ") or header.startswith("subject/"):
            return index
    return None


def _find_terms_column(headers: list[str]) -> int | None:
    for index, header in enumerate(headers):
        if "keyword" in header and ("pexels" in header or "search" in header):
            return index
    return None


def _normalize_header(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("_", " ")).strip().lower()


def _cell(row: list[str], index: int) -> str:
    if index >= len(row):
        return ""
    return str(row[index] or "").strip()


def _split_terms(value: str) -> list[str]:
    if not value:
        return []
    separator = ";" if ";" in value else ","
    return [part.strip() for part in value.split(separator) if part.strip()]
