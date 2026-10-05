import io
import re
import zipfile
from collections import defaultdict
from copy import copy

import fitz
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Universal Precision Validation", page_icon="🧪", layout="wide")
st.title("🧪 Universal Precision Validation")
st.caption("5 replicates × 5 Intra Runs + dynamic Inter Runs")

NUM = r'[<>]?\s*-?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?'
RECORD_RE = re.compile(
    rf'\bPR\s+(INTRA|INTER)\s+([A-Za-z0-9_.-]+)\s+({NUM})\s*([^\s]+)',
    re.I
)

def normalize_analyte(name):
    name = re.sub(r'[^A-Za-z0-9]+', '_', str(name).strip())
    return re.sub(r'_+', '_', name).strip('_') or "Analyte"

def level_from_sid(mode, sid):
    s = str(sid).upper().strip()
    if mode == "INTRA":
        m = re.match(r'L([123])(?:[.]?01)?$', s) or re.search(r'L([123])', s)
        return int(m.group(1)) if m else None

    if re.fullmatch(r'[34]0[1-3]', s):
        return int(s[-1])

    m = re.fullmatch(r'R20([1-3])', s)
    if m:
        return int(m.group(1))

    m = re.search(r'0([1-3])$', s)
    return int(m.group(1)) if m else None

def extract_pdf_records(uploaded_file, expected_mode):
    doc = fitz.open(stream=uploaded_file.getvalue(), filetype="pdf")
    records = []

    for page_no, page in enumerate(doc, 1):
        text = page.get_text("text") or ""
        if not text:
            continue

        assay = None
        m_assay = re.search(r'Tests:\s*([^;\n]+)', text, re.I)
        if m_assay:
            assay = m_assay.group(1).strip()

        if not assay:
            lines = [x.strip() for x in text.splitlines() if x.strip()]
            for i, line in enumerate(lines):
                if line.lower().startswith("test status") and i + 1 < len(lines):
                    if lines[i + 1].lower() != "assay report":
                        assay = lines[i + 1]
                        break

        for m in RECORD_RE.finditer(text):
            mode = m.group(1).upper()
            if mode != expected_mode:
                continue

            sid = m.group(2)
            raw = m.group(3).replace(" ", "")

            try:
                result = float(raw.replace("<", "").replace(">", ""))
            except ValueError:
                continue

            level = level_from_sid(mode, sid)
            if level is None:
                continue

            records.append({
                "analyte": assay or "Unknown",
                "sid": sid,
                "level": level,
                "result": result,
                "unit": m.group(4),
                "page": page_no
            })

    doc.close()
    return records

def group_records(records):
    grouped = defaultdict(lambda: defaultdict(list))
    units = {}

    for r in records:
        grouped[r["analyte"]][r["level"]].append(r["result"])
        units.setdefault(r["analyte"], r["unit"])

    return grouped, units

def get_cell(ws, row, col):
    cell = ws.cell(row, col)

    if not isinstance(cell, MergedCell):
        return cell

    for rng in ws.merged_cells.ranges:
        if cell.coordinate in rng:
            min_col, min_row, max_col, max_row = rng.bounds
            if row == min_row and col == min_col:
                return ws.cell(min_row, min_col)
            raise ValueError(
                f"Merged cell {cell.coordinate} is inside merged range {rng}."
            )

    return cell

def put(ws, row, col, value):
    get_cell(ws, row, col).value = value

def copy_row_style(ws, source_row, target_row):
    for col in range(1, ws.max_column + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)

        if src.has_style:
            dst._style = copy(src._style)

        dst.number_format = src.number_format
        dst.alignment = copy(src.alignment)
        dst.protection = copy(src.protection)

    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height

def find_text(ws, text):
    target = str(text).strip().lower()

    for row in ws.iter_rows():
        for cell in row:
            if str(cell.value or "").strip().lower() == target:
                return cell.row, cell.column

    return None, None

def detect_template(ws):
    """
    Detects the uploaded precision template.

    Supported primary structure:
      INTRA: 5 runs × 5 replicates
      INTER: one 5-replicate column per run
    The actual uploaded workbook remains the base workbook.
    """

    layout = {
        "intra_start_row": 11,
        "intra_rep_end_row": 15,
        "intra_run_cols": [2, 3, 4, 5, 6],

        "inter_start_row": 11,
        "inter_rep_end_row": 15,
        "inter_run_cols": [7, 8, 9, 10, 11],

        "stats_start_row": 17,
        "stats_end_row": 20,

        "analyte_cell": "A3",
        "unit_cell": "E7",
    }

    # Find the row containing "Replicate".
    rep_row, rep_col = find_text(ws, "Replicate")

    if rep_row is not None:
        layout["rep_header_row"] = rep_row
        layout["rep_col"] = rep_col

        # Five rows immediately below Replicate are the first 5 replicates.
        layout["intra_start_row"] = rep_row + 1
        layout["intra_rep_end_row"] = rep_row + 5
        layout["inter_start_row"] = rep_row + 1
        layout["inter_rep_end_row"] = rep_row + 5

    # Detect "Inter" and "Intra" header columns.
    intra_row, intra_col = find_text(ws, "Intra")
    inter_row, inter_col = find_text(ws, "Inter")

    if intra_col is not None:
        layout["intra_header_col"] = intra_col

    if inter_col is not None:
        layout["inter_header_col"] = inter_col

    # Detect summary rows.
    for r in range(1, min(ws.max_row, 100) + 1):
        values = [
            str(ws.cell(r, c).value or "").strip().lower()
            for c in range(1, min(ws.max_column, 30) + 1)
        ]

        joined = " ".join(values)

        if re.search(r"\bmean\b", joined):
            layout["mean_row"] = r

        if re.search(r"\b(?:sd|s\.d\.?)\b", joined):
            layout["sd_row"] = r

        if re.search(r"\bcv\s*%", joined):
            layout["cv_row"] = r

        if re.search(r"\b(?:n|count)\b", joined):
            layout["n_row"] = r

    return layout

def populate_statistics(ws, cols, start_row, end_row, n_row, mean_row, sd_row, cv_row):
    for col in cols:
        letter = get_column_letter(col)

        put(ws, n_row, col, f'=COUNT({letter}{start_row}:{letter}{end_row})')
        put(
            ws,
            mean_row,
            col,
            f'=IFERROR(AVERAGE({letter}{start_row}:{letter}{end_row}),"")'
        )
        put(
            ws,
            sd_row,
            col,
            f'=IFERROR(STDEV({letter}{start_row}:{letter}{end_row}),"")'
        )
        put(
            ws,
            cv_row,
            col,
            f'=IFERROR(({letter}{sd_row}/{letter}{mean_row})*100,"")'
        )

def populate_workbook(
    template_bytes,
    analyte,
    unit,
    n_levels,
    intra_by_level,
    inter_runs_by_level,
    detected_layout
):
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active

    layout = dict(detected_layout)

    # We need one 5-row block for every Intra run.
    # If the template has 5 Intra run columns, the first five are used.
    # Additional Intra runs can be added by inserting new columns before
    # the Inter section.
    intra_run_count = 5

    # INTER run count is determined from the uploaded INTER PDFs.
    inter_run_count = max(
        [len(inter_runs_by_level.get(level, [])) for level in range(1, n_levels + 1)]
        or [0]
    )

    # -------------------------
    # INTRA
    # -------------------------
    # Each INTRA PDF is treated as a source of 5 replicate blocks.
    # We flatten the extracted values into 5-value runs.
    for level in range(1, n_levels + 1):
        values = list(intra_by_level.get(level, []))

        # The supplied template has 5 Intra run columns × 5 replicates = 25.
        # If more values are available, create additional run columns.
        required_intra_runs = max(5, (len(values) + 4) // 5)

        extra_intra_cols = max(0, required_intra_runs - 5)

        if extra_intra_cols:
            # Insert before the current Inter area.
            insert_at = min(layout["inter_run_cols"])
            ws.insert_cols(insert_at, amount=extra_intra_cols)

            # Copy the last existing Intra column's style into the new columns.
            source_col = insert_at - 1
            for offset in range(extra_intra_cols):
                new_col = insert_at + offset
                for row in range(1, ws.max_row + 1):
                    src = ws.cell(row, source_col)
                    dst = ws.cell(row, new_col)
                    if src.has_style:
                        dst._style = copy(src._style)
                    dst.number_format = src.number_format
                    dst.alignment = copy(src.alignment)

            layout["inter_run_cols"] = [
                c + extra_intra_cols if c >= insert_at else c
                for c in layout["inter_run_cols"]
            ]

            layout["intra_run_cols"] = (
                layout["intra_run_cols"] +
                [insert_at + i for i in range(extra_intra_cols)]
            )

    # -------------------------
    # INTRA population
    # -------------------------
    for level in range(1, n_levels + 1):
        values = list(intra_by_level.get(level, []))

        # Current template uses one level per separate workbook,
        # so level is represented by the extracted analyte workbook.
        # Populate all available Intra runs in 5-value blocks.
        for run_index, col in enumerate(layout["intra_run_cols"]):
            block = values[run_index * 5:(run_index + 1) * 5]

            for rep_index, value in enumerate(block):
                put(
                    ws,
                    layout["intra_start_row"] + rep_index,
                    col,
                    value
                )

    # -------------------------
    # INTER population
    # -------------------------
    # Run 1 = first 5 INTRA results.
    # Run 2 = INTER PDF 1.
    # Run 3 = INTER PDF 2.
    # etc.
    inter_values = {}

    for level in range(1, n_levels + 1):
        values = list(intra_by_level.get(level, []))[:5]
        inter_values[level] = [values]

        for run_values in inter_runs_by_level.get(level, []):
            inter_values[level].append(list(run_values)[:5])

    max_inter_runs = max(
        [len(v) for v in inter_values.values()] or [1]
    )

    # Expand INTER columns if needed.
    existing_inter = len(layout["inter_run_cols"])

    if max_inter_runs > existing_inter:
        extra = max_inter_runs - existing_inter
        insert_at = ws.max_column + 1

        source_col = layout["inter_run_cols"][-1]

        ws.insert_cols(insert_at, amount=extra)

        for i in range(extra):
            new_col = insert_at + i
            for row in range(1, ws.max_row + 1):
                src = ws.cell(row, source_col)
                dst = ws.cell(row, new_col)
                if src.has_style:
                    dst._style = copy(src._style)
                dst.number_format = src.number_format
                dst.alignment = copy(src.alignment)

            layout["inter_run_cols"].append(new_col)

    # Populate the Inter area.
    # Because each generated workbook is analyte-specific, level is the
    # QC level selected by the source data. For a 3-level template, the
    # first three inter columns are used for L1/L2/L3 when applicable.
    for level in range(1, n_levels + 1):
        values_by_run = inter_values[level]

        for run_index, run_values in enumerate(values_by_run):
            if run_index >= len(layout["inter_run_cols"]):
                break

            col = layout["inter_run_cols"][run_index]

            for rep_index, value in enumerate(run_values):
                put(
                    ws,
                    layout["inter_start_row"] + rep_index,
                    col,
                    value
                )

    # -------------------------
    # Statistics
    # -------------------------
    stats_rows = {
        "n": layout.get("n_row", 17),
        "mean": layout.get("mean_row", 18),
        "sd": layout.get("sd_row", 19),
        "cv": layout.get("cv_row", 20)
    }

    populate_statistics(
        ws,
        layout["intra_run_cols"],
        layout["intra_start_row"],
        layout["intra_rep_end_row"],
        stats_rows["n"],
        stats_rows["mean"],
        stats_rows["sd"],
        stats_rows["cv"]
    )

    populate_statistics(
        ws,
        layout["inter_run_cols"],
        layout["inter_start_row"],
        layout["inter_rep_end_row"],
        stats_rows["n"],
        stats_rows["mean"],
        stats_rows["sd"],
        stats_rows["cv"]
    )

    # Template-specific title/unit where those cells exist.
    try:
        if ws["A3"].value is not None:
            ws["A3"] = f"Precision Studies of - {analyte}"
    except Exception:
        pass

    try:
        if unit:
            ws[layout.get("unit_cell", "E7")] = unit
    except Exception:
        pass

    # Hide level 3 columns only when the source is explicitly two-level.
    if n_levels == 2:
        pass

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    return wb

st.sidebar.header("Universal Precision Settings")

template = st.file_uploader(
    "1. Upload Precision Excel template",
    type=["xlsx", "xlsm"]
)

intra_files = st.file_uploader(
    "2. Upload INTRA PDF(s)",
    type=["pdf"],
    accept_multiple_files=True
)

inter_files = st.file_uploader(
    "3. Upload INTER PDF(s) in RUN order",
    type=["pdf"],
    accept_multiple_files=True
)

level_mode = st.sidebar.selectbox(
    "QC level configuration",
    ["Auto Detect", "2 Levels", "3 Levels"]
)

st.info(
    "Population logic: INTRA = 5 replicates per Run. "
    "The template supports 5 Intra Runs by default. "
    "INTER Run 1 = first 5 INTRA values; INTER Run 2 = INTER PDF 1; "
    "INTER Run 3 = INTER PDF 2; and so on."
)

if template:
    try:
        preview = load_workbook(io.BytesIO(template.getvalue()))
        preview_ws = preview.active
        detected = detect_template(preview_ws)

        st.sidebar.success("Template detected")
        st.sidebar.write(
            f"INTRA rows: {detected['intra_start_row']}–"
            f"{detected['intra_rep_end_row']}"
        )
        st.sidebar.write(
            "INTRA run columns: " +
            ", ".join(get_column_letter(c) for c in detected["intra_run_cols"])
        )
        st.sidebar.write(
            "INTER run columns: " +
            ", ".join(get_column_letter(c) for c in detected["inter_run_cols"])
        )
    except Exception as exc:
        detected = None
        st.sidebar.error(f"Template detection error: {exc}")
else:
    detected = None

if st.button(
    "🔍 PROCESS UNIVERSAL PRECISION",
    type="primary",
    use_container_width=True
):
    if not template or not intra_files:
        st.error("Please upload the Excel template and at least one INTRA PDF.")
        st.stop()

    if detected is None:
        st.error("Unable to detect the Excel template.")
        st.stop()

    with st.spinner("Extracting precision data..."):
        intra_records = []

        for f in intra_files:
            intra_records.extend(
                extract_pdf_records(f, "INTRA")
            )

        intra_group, units = group_records(intra_records)

        inter_file_groups = []

        for f in inter_files or []:
            records = extract_pdf_records(f, "INTER")
            grouped, grouped_units = group_records(records)

            inter_file_groups.append(
                (f.name, grouped, grouped_units)
            )

    if not intra_records:
        st.error(
            "No INTRA result records were detected. "
            "The PDF does not match the currently supported PR record format."
        )
        st.stop()

    detected_levels = sorted(
        {
            level
            for analyte_data in intra_group.values()
            for level in analyte_data
            if level in (1, 2, 3)
        }
    )

    n_levels = (
        2 if level_mode == "2 Levels"
        else 3 if level_mode == "3 Levels"
        else (3 if 3 in detected_levels else 2)
    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("Analytes", len(intra_group))
    c2.metric("QC levels", n_levels)
    c3.metric("INTRA records", len(intra_records))
    c4.metric("INTER PDFs", len(inter_file_groups))

    outputs = {}
    warnings = []

    for analyte in sorted(intra_group):

        intra_by_level = {
            level: list(
                intra_group[analyte].get(level, [])
            )
            for level in range(1, n_levels + 1)
        }

        for level in range(1, n_levels + 1):
            if len(intra_by_level[level]) < 25:
                warnings.append(
                    f"{analyte} L{level}: "
                    f"{len(intra_by_level[level])} INTRA values detected; "
                    f"template expects 25 values for 5 runs × 5 replicates."
                )

        inter_runs_by_level = defaultdict(list)

        for filename, grouped, _ in inter_file_groups:
            for level in range(1, n_levels + 1):
                vals = list(
                    grouped.get(analyte, {}).get(level, [])
                )[:5]

                if len(vals) == 5:
                    inter_runs_by_level[level].append(vals)
                elif vals:
                    warnings.append(
                        f"{analyte} L{level}: {filename} has "
                        f"{len(vals)}/5 INTER values; run left blank."
                    )
                else:
                    warnings.append(
                        f"{analyte} L{level}: {filename} has no INTER values; "
                        f"run left blank."
                    )

        wb = populate_workbook(
            template.getvalue(),
            analyte,
            units.get(analyte, ""),
            n_levels,
            intra_by_level,
            inter_runs_by_level,
            detected
        )

        outputs[
            f"{normalize_analyte(analyte)}_Universal_Precision.xlsx"
        ] = wb_bytes(wb)

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(
        zip_buffer,
        "w",
        zipfile.ZIP_DEFLATED
    ) as z:
        for filename, data in outputs.items():
            z.writestr(filename, data)

    zip_buffer.seek(0)

    st.success(
        f"Created {len(outputs)} analyte workbook(s)."
    )

    if warnings:
        st.warning("Review:")
        for warning in warnings:
            st.write("• " + warning)

    st.download_button(
        "⬇️ DOWNLOAD ALL ANALYTES — ZIP",
        data=zip_buffer.getvalue(),
        file_name="Universal_Precision_Validation_Output.zip",
        mime="application/zip",
        use_container_width=True
    )
