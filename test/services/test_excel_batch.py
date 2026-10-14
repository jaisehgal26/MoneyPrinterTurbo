import io
import zipfile
from xml.sax.saxutils import escape

from app.services import excel_batch


def _xlsx(headers, data_rows):
    shared = []

    def shared_index(value):
        text = str(value)
        if text not in shared:
            shared.append(text)
        return shared.index(text)

    def cell(column, row_number, value):
        return (
            f'<c r="{column}{row_number}" t="s">'
            f"<v>{shared_index(value)}</v></c>"
        )

    body_rows = []
    for row_number, values in enumerate([headers, *data_rows], start=1):
        cells = "".join(
            cell(chr(ord("A") + index), row_number, value)
            for index, value in enumerate(values)
        )
        body_rows.append(f'<row r="{row_number}">{cells}</row>')
    sheet = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f"<sheetData>{''.join(body_rows)}</sheetData></worksheet>"
    )
    strings = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + "".join(f"<si><t>{escape(item)}</t></si>" for item in shared)
        + "</sst>"
    )
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Video Plan" sheetId="1" r:id="rId1"/></sheets></workbook>',
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            'Target="worksheets/sheet1.xml"/></Relationships>',
        )
        archive.writestr("xl/worksheets/sheet1.xml", sheet)
        archive.writestr("xl/sharedStrings.xml", strings)
    return buffer.getvalue()


def test_parse_subject_and_video_id_columns():
    workbook = _xlsx(
        ["Video ID", "YouTube Title", "Subject / Video Brief", "Pexels Search Keywords", "YouTube Video ID"],
        [
            ["MD-001", "Title", "Why memories return", "rainy window; thinking", ""],
            ["", "", "", "", ""],
            ["MD-002", "Title 2", "Spotlight effect", "crowd, street", "abc123"],
        ],
    )

    rows = excel_batch.parse_excel_rows(workbook)

    assert [(row.video_id, row.subject, row.search_terms) for row in rows] == [
        ("MD-001", "Why memories return", ["rainy window", "thinking"]),
        ("MD-002", "Spotlight effect", ["crowd", "street"]),
    ]


def test_video_id_rejects_path_characters():
    workbook = _xlsx(
        ["Video ID", "Subject"],
        [["../secret", "A subject"]],
    )

    try:
        excel_batch.parse_excel_rows(workbook)
    except ValueError as exc:
        assert "Video ID" in str(exc)
    else:
        raise AssertionError("path-like video id was accepted")


def test_folder_selector_stays_inside_allowed_roots(tmp_path, monkeypatch):
    storage = tmp_path / "storage"
    home = tmp_path / "home"
    storage.mkdir()
    (home / "Videos").mkdir(parents=True)
    monkeypatch.setenv("EXCEL_HOST_FOLDER_ROOT", str(home))
    monkeypatch.setenv("EXCEL_HOST_FOLDER_LABEL", r"C:\Users\sehga")
    monkeypatch.setenv("EXCEL_STORAGE_LABEL", r"C:\Users\sehga\Videos\storage")
    monkeypatch.setattr(excel_batch.utils, "storage_dir", lambda sub_dir="", create=False: str(storage))

    roots = [root for _, root in excel_batch.allowed_output_roots()]

    assert excel_batch.folder_is_allowed(home / "Videos", roots)
    assert not excel_batch.folder_is_allowed(tmp_path, roots)
    assert excel_batch.display_folder(home / "Videos") == r"C:\Users\sehga\Videos"
    assert excel_batch.safe_child_folder_name("Mind Decoded") == "Mind Decoded"
    destination = excel_batch.output_path(tmp_path, "MD-001")

    assert destination == tmp_path / "MD-001.mp4"
    assert destination.parent == tmp_path
