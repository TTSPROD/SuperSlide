"""Semantic table structure, distinct from PowerPoint's first-row styling.

External column headings are associated by their displayed geometry. Embedded
header depths are reviewed against the exact supplied table XML; unfamiliar
tables are left unspecified rather than guessed from ``firstRow``.
"""
from __future__ import annotations

import hashlib

from lxml import etree as E


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P, "a": A}

# Canonical a:tbl fingerprints of the supplied native tables. Counts describe
# semantic header rows, including multi-level headers, not formatting bands.
_REVIEWED_HEADER_ROWS = {
    "b582fe99003451490d0e2b463ef8f546616a4562d196c91a2d365f80c83a1f67": 0,  # contents 2, 3
    "ecac21fe8094ac566594469f34c0d8ad8dad35c3d9e00fb22cec2fe1fa06c115": 0,  # contents 4
    "eeb457cbf0d87d4686112c0dbcf4b88690c3c908208d10f122e56cd69416a21a": 0,  # cards 10, left
    "151ffc090dab994a4a6ebbad7c00219e4d4a8c63010fcb41a6a7041c97146fca": 0,  # cards 10, right
    "46aae665adf4fbe6b2fabfcc7764c80574ff62d333899c40a1affb968a5b90e4": 1,  # tables 1
    "3758f2213f11b6515e4cfeeea18f7e2a7c7db26a2889889a473336eea6fdcb3c": 1,  # tables 2, right
    "6853f044be89855588dad4aac6d6a9b97290034eaa9964b74b23365e7fc4bf16": 1,  # tables 2, left
    "f9eb45956c37b2b0f813132bb6dff99ff73b0ae2542250023bd2c7fe6560765c": 2,  # tables 3
    "dbb7da5514e452f9e11d32bd37b555a8f0238d05c31955ee8ebfb0bf2d2773b6": 1,  # tables 4
    "513ad4e73336509ba4c99a4c62e57d26c14096ac6e89d3e45ec069f95833dd23": 1,  # tables 5
    "deb6ba0a10ceef16bc0920f0fc636744feca8536e28f24be222f9019da4ad8c2": 2,  # tables 6
    "bb1efeb1d104df94a2ba2df0d80e182dd96cc554e5b9c9f2d05a04388cb3511d": 0,  # tables 8, external
    "87fde60a4d08269bc4d922b06588b0a5e136c2214c5df409f2a0f31b52cd11bc": 1,  # tables 9
    "f968f196f0087660d32c709255e97a40b52245e68c83bdedd9b3aef5e471b1b4": 1,  # tables 10
    "cecba282757df00badee5de1b24cf1bab5b921716aa23d7a9f9a3b7b44fbd810": 1,  # tables 11
    "c679a3847b6b9609f213ca9b96f897716ba2490ffab5f9579e2f1376da088102": 1,  # tables 12
    "8f9be91b00c20786afdc83a0180dabde96561fb4f673cef4f575a1f1ec099046": 2,  # tables 13
    "8b59a0ef1886abb06ac46fadacf30ef9fe74b19b0be3c0c09d9791d5bc8081d7": 2,  # tables 14
    "bb2aaa12b9ed47cb23930c59969ba4b06e78699ee762f236c6674380de196f1d": 1,  # tables 15
    "71e473cf7cfc89391604fbee4cf58f1194b1e074f22d6e565b3d3484398cbf9b": 1,  # tables 16
    "57c5838d800d7c56f30468b0d6800916b40295d6c07361d022a0f7505f9fb0e5": 2,  # tables 17
    "725407a320367954d744ae384019970e4942038976b1e4d1c51f60b3585c589f": 1,  # tables 18
    "d3dfdd5a2e87e4ac9d581d2f440923784cc5b2df08e71d56e7f1c6eb9b1dee66": 1,  # tables 19
    "b52e1d668a97576d37b48cb30e87695bb5f6aaaf60c37c4bbf189d5454be12d6": 1,  # tables 20
    "96affe9ebe00d032adb5438fd0af475a960411472931d94be0970317c65bc126": 0,  # circle_schemes 6
}

# These rows have their own semantic purpose within the source body. They are
# not extra factual data slots. The fingerprints cover the complete source
# table, so unfamiliar revisions never inherit these reviewed row indexes.
_REVIEWED_BODY_ROWS = {
    "deb6ba0a10ceef16bc0920f0fc636744feca8536e28f24be222f9019da4ad8c2": {"total": [17]},
    "87fde60a4d08269bc4d922b06588b0a5e136c2214c5df409f2a0f31b52cd11bc": {"section_header": [1, 6]},
    "f968f196f0087660d32c709255e97a40b52245e68c83bdedd9b3aef5e471b1b4": {"section_header": [1, 5]},
    "57c5838d800d7c56f30468b0d6800916b40295d6c07361d022a0f7505f9fb0e5": {"total": [8]},
}


def _shape_id(shape):
    found = shape.xpath("./*/p:cNvPr/@id", namespaces=NS)
    return int(found[0]) if found else None


def _transform(shape):
    for path in ("p:spPr/a:xfrm", "p:grpSpPr/a:xfrm", "p:xfrm"):
        found = shape.find(path, NS)
        if found is not None:
            return found
    return None


def _straight(transform):
    return (int(transform.get("rot", 0)) % 21600000 == 0
            and transform.get("flipH") not in {"1", "true"}
            and transform.get("flipV") not in {"1", "true"})


def _display_box(shape):
    """Resolve axis-aligned source geometry through nested group transforms."""
    transform = _transform(shape)
    if transform is None or not _straight(transform):
        return None
    off, ext = transform.find("a:off", NS), transform.find("a:ext", NS)
    if off is None or ext is None:
        return None
    x, y, width, height = (float(off.get("x")), float(off.get("y")),
                           float(ext.get("cx")), float(ext.get("cy")))
    for group in shape.iterancestors(f"{{{P}}}grpSp"):
        transform = _transform(group)
        if transform is None or not _straight(transform):
            return None
        off, ext = transform.find("a:off", NS), transform.find("a:ext", NS)
        child_off, child_ext = transform.find("a:chOff", NS), transform.find("a:chExt", NS)
        if any(value is None for value in (off, ext, child_off, child_ext)):
            return None
        if min(float(child_ext.get("cx")), float(child_ext.get("cy"))) <= 0:
            return None
        sx = float(ext.get("cx")) / float(child_ext.get("cx"))
        sy = float(ext.get("cy")) / float(child_ext.get("cy"))
        x = float(off.get("x")) + (x - float(child_off.get("x"))) * sx
        y = float(off.get("y")) + (y - float(child_off.get("y"))) * sy
        width, height = width * sx, height * sy
    return (x, y, width, height) if min(width, height) > 0 else None


def _external_headers(frame, table, fields, shapes):
    box = _display_box(frame)
    widths = [int(column.get("w")) for column in table.findall("a:tblGrid/a:gridCol", NS)]
    if box is None or not widths or min(widths) <= 0:
        return []
    x, y, width, _ = box
    columns = []
    for column_width in widths:
        displayed_width = width * column_width / sum(widths)
        columns.append((x, displayed_width))
        x += displayed_width
    candidates = [[] for _ in columns]
    for field_id, field in fields.items():
        shape = shapes.get(field.get("shape_id"))
        if field.get("kind") != "text" or shape is None or field.get("automatic"):
            continue
        # Slide titles can span the entire table; they are not column headings.
        if shape.xpath("./p:nvSpPr/p:nvPr/p:ph[@type='title' or @type='ctrTitle']", namespaces=NS):
            continue
        header_box = _display_box(shape)
        if header_box is None:
            continue
        hx, hy, hw, hh = header_box
        gap = y - hy - hh
        if not -12700 <= gap <= 457200 or hh > 457200:
            continue
        for index, (cx, cw) in enumerate(columns):
            if (abs(hx - cx) <= max(63500, cw * .15)
                    and hw <= cw * 1.2 and hx + hw <= cx + cw * 1.15):
                candidates[index].append((gap, field_id, hy, hh))
    if any(len(items) != 1 for items in candidates):
        return []  # Ambiguous or partial associations do not define a header row.
    headings = [items[0] for items in candidates]
    if max(item[2] for item in headings) - min(item[2] for item in headings) > 127000:
        return []
    return [item[1] for item in headings]


def _table_row_structure(table, header_rows, fingerprint):
    """Separate header, section, data and total slots in the unchanged grid."""
    rows = table.findall("a:tr", NS)
    roles = ["column_header" if index < header_rows else "data" for index in range(len(rows))]
    for role, indexes in _REVIEWED_BODY_ROWS.get(fingerprint, {}).items():
        for index in indexes:
            roles[index] = role
    data = [index for index, role in enumerate(roles) if role == "data"]
    sections = [index for index, role in enumerate(roles) if role == "section_header"]
    totals = [index for index, role in enumerate(roles) if role == "total"]
    merges, continuations = [], []
    for row_index, row in enumerate(rows):
        for column_index, cell in enumerate(row.findall("a:tc", NS)):
            if any(cell.get(name) in {"1", "true"} for name in ("hMerge", "vMerge")):
                continuations.append([row_index, column_index])
                continue
            row_span, column_span = int(cell.get("rowSpan", 1)), int(cell.get("gridSpan", 1))
            if row_span > 1 or column_span > 1:
                merges.append({"row_index": row_index, "column_index": column_index,
                               "row_span": row_span, "column_span": column_span})
    return {"body_row_count": len(rows) - header_rows, "row_roles": roles,
            "data_row_count": len(data), "data_row_indexes": data,
            "section_header_row_count": len(sections), "section_header_row_indexes": sections,
            "total_row_count": len(totals), "total_row_indexes": totals,
            "merged_cells": merges, "merged_continuation_cells": continuations}


def annotate_table_structures(root, fields, scope):
    """Attach table/header associations to extracted fields, leaving XML intact."""
    shapes = {_shape_id(shape): shape for shape in root.xpath(
        ".//p:sp | .//p:graphicFrame", namespaces=NS)}
    for field_id, field in fields.items():
        if field.get("kind") != "table" or field.get("scope") != scope:
            continue
        frame = shapes.get(field["shape_id"])
        table = frame.find(".//a:tbl", NS) if frame is not None else None
        if table is None:
            continue
        fingerprint = hashlib.sha256(E.tostring(table, method="c14n")).hexdigest()
        external = _external_headers(frame, table, fields, shapes)
        if external:
            field["table_structure"] = {"header_placement": "external", "header_row_count": 0,
                                        "column_header_field_ids": external,
                                        **_table_row_structure(table, 0, fingerprint)}
            for index, header_id in enumerate(external, 1):
                fields[header_id]["table_column_header"] = {"table_field_id": field_id, "column_index": index}
            continue
        header_rows = _REVIEWED_HEADER_ROWS.get(fingerprint)
        if header_rows is None:
            field["table_structure"] = {"header_placement": "unspecified"}
        else:
            field["table_structure"] = {"header_placement": "inside" if header_rows else "none",
                                        "header_row_count": header_rows,
                                        **_table_row_structure(table, header_rows, fingerprint)}
    annotate_drawn_table_structures(root, fields, scope)
    return fields


# These compositions use independent text and chart objects rather than a:tbl.
# Roles were reviewed against their displayed source slides. Whole-slide
# fingerprints prevent the associations silently surviving a design change.
_DRAWN_TABLES = {
    "d24e4c75effad22f10c8357d450de178793e27a214c9a0cd7bfb1f3a9b78360f": {
        "id": "tables_007",
        "headers": [21, 24, 27, 30],
        # Column zero is the unlabelled row-description stub. Missing entries
        # are intentionally empty source areas, not additional fillable slots.
        "rows": [
            {0: 7, 1: 3, 2: 8, 3: 11, 4: 15},
            {0: 10, 1: 13, 2: 62, 3: 64},
            {0: 115, 1: 122, 2: 123, 3: 124, 4: 114},
            {0: 18, 1: 66, 2: 68, 3: 70},
            {0: 16, 1: 72, 2: 74, 3: 76},
            {0: 17, 1: 78, 2: 80, 3: 82},
        ],
        "row_spans": {15: 2},
        "row_sections": [1, 1, 1, 2, 2, 2],
        "sections": {119: 0, 36: 1, 127: 2},
    },
    "9693a7cadb94549a0ba11e33873a55c31a2083c8fd651e5470274170c12375c2": {
        "id": "diagram_009",
        "headers": [28, 40, 42],
        "rows": [{1: sid} for sid in (23, 24, 25, 26, 27, 29)],
        "column_charts": {5: 2, 39: 3},
    },
    "44dbabbb3e32b9f8ef5366f90fbe224c38b1b6d673087d603907ac614626d3b0": {
        "id": "diagram_032",
        "headers": [107, 18, 17, 19],
        "rows": [dict(zip((1, 2, 3), ids)) for ids in zip(
            (21, 23, 24, 25, 26, 27, 28, 29), range(36, 44), range(44, 52))],
        "row_charts": [54, 57, 91, 69, 67, 75, 78, 93],
    },
}
DRAWN_TABLE_ROLE_KEYS = (
    "drawn_table_column_header", "drawn_table_cell",
    "drawn_table_section_header", "drawn_table_column_chart",
)


def annotate_drawn_table_structures(root, fields, scope):
    """Associate the reviewed drawn table headings with their existing cells.

    Each source field keeps its native operation (text or chart). No synthetic
    PowerPoint table or combined content field is introduced, and source XML
    is never changed. Re-annotation removes stale roles if a source changes.
    """
    for field in fields.values():
        for key in DRAWN_TABLE_ROLE_KEYS:
            field.pop(key, None)
    if scope != "slide":
        return fields
    fingerprint = hashlib.sha256(E.tostring(root, method="c14n")).hexdigest()
    reviewed = _DRAWN_TABLES.get(fingerprint)
    if reviewed is None:
        return fields
    table_id = reviewed["id"]

    def tag(sid, role, values, kind="text"):
        field_id = f"{scope}_{kind}_{sid}"
        if field_id in fields:
            fields[field_id][role] = {"table_id": table_id, **values}

    for column, sid in enumerate(reviewed["headers"], 1):
        tag(sid, "drawn_table_column_header", {"column_index": column})
    for row, columns in enumerate(reviewed["rows"], 1):
        for column, sid in columns.items():
            values = {"row_index": row, "column_index": column}
            if sid in reviewed.get("row_spans", {}):
                values["row_span"] = reviewed["row_spans"][sid]
            if "row_sections" in reviewed:
                values["section_index"] = reviewed["row_sections"][row - 1]
            tag(sid, "drawn_table_cell", values)
    for sid, section in reviewed.get("sections", {}).items():
        values = {"section_index": section}
        if section == 0:
            values["role"] = "section_column_header"
        tag(sid, "drawn_table_section_header", values)
    for sid, column in reviewed.get("column_charts", {}).items():
        tag(sid, "drawn_table_column_chart", {"column_index": column,
                                             "row_span": len(reviewed["rows"])}, kind="chart")
    for row, sid in enumerate(reviewed.get("row_charts", []), 1):
        tag(sid, "drawn_table_cell", {"row_index": row,
                                      "column_index": len(reviewed["headers"])}, kind="chart")
    return fields
