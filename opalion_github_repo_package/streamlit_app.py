
import csv
import io
import os
import re
from datetime import datetime, timedelta
from zipfile import ZipFile
import xml.etree.ElementTree as ET

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Opalion Delivery Note Converter", layout="wide")

APP_TITLE = "Opalion Delivery Note Converter"

NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

OUTPUT_FIELDS = [
    "Owner",
    "Additional Ref",
    "Delivery Company Name",
    "Delivery Address 1",
    "Postcode",
    "Delivery Name",
    "Delivery Date",
    "Order Type",
    "Product Code",
    "Quantity",
]


def t(v):
    if v is None:
        return ""
    return str(v).replace("_x000D__x000A_", "\n").replace("_x000A_", "\n").replace("_x000D_", "\r").strip()


def norm_ref(ref):
    return str(ref).upper() if ref else ref


def serial_to_date(v):
    s = t(v)
    if not s:
        return ""
    try:
        n = float(s)
        dt = datetime(1899, 12, 30) + timedelta(days=n)
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return s


def apostrophe_if_leading_zero(code):
    code = t(code)
    if code and code.startswith("0"):
        return "'" + code
    return code


def parse_postcode(block):
    block = t(block)
    m = re.search(r"\b([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})\b", block, re.I)
    return m.group(1).upper().replace("  ", " ") if m else ""


def parse_address_block(block):
    block = t(block)
    lines = [ln.strip() for ln in block.splitlines() if ln.strip()]
    cleaned = []
    for line in lines:
        line = re.sub(r"^(attn?|att):\s*", "", line, flags=re.I).strip()
        if line:
            cleaned.append(line)
    company = cleaned[0] if len(cleaned) > 0 else ""
    address1 = cleaned[1] if len(cleaned) > 1 else ""
    postcode = parse_postcode(block)
    return company, address1, postcode


def load_cells_from_upload(uploaded_file):
    # Streamlit uploads provide a file-like object; we read bytes from it.
    data = uploaded_file.getvalue()
    with ZipFile(io.BytesIO(data)) as zf:
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

        sheet_el = wb.find("a:sheets/a:sheet", NS)
        if sheet_el is None:
            raise ValueError("Workbook has no sheet.")

        rid = sheet_el.attrib.get(f"{{{NS['r']}}}id")
        target = None
        for rel in rels.findall("rel:Relationship", NS):
            if rel.attrib.get("Id") == rid:
                target = rel.attrib.get("Target")
                break
        if not target:
            raise ValueError("Could not resolve first sheet path.")

        sheet_path = target if target.startswith("xl/") else f"xl/{target}"
        sheet = ET.fromstring(zf.read(sheet_path))

        shared_strings = []
        if "xl/sharedStrings.xml" in zf.namelist():
            sst = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in sst.findall("a:si", NS):
                shared_strings.append("".join(tt.text or "" for tt in si.iterfind(".//a:t", NS)))

        cells = {}
        for row in sheet.findall(".//a:sheetData/a:row", NS):
            for c in row.findall("a:c", NS):
                ref = norm_ref(c.attrib.get("r"))
                if not ref:
                    continue
                t_attr = c.attrib.get("t")
                v = c.find("a:v", NS)
                is_ = c.find("a:is", NS)

                if t_attr == "s" and v is not None:
                    idx = int(v.text or "0")
                    cells[ref] = shared_strings[idx] if 0 <= idx < len(shared_strings) else ""
                elif t_attr == "inlineStr" and is_ is not None:
                    cells[ref] = "".join(tt.text or "" for tt in is_.iterfind(".//a:t", NS))
                elif v is not None:
                    cells[ref] = v.text or ""
                elif is_ is not None:
                    cells[ref] = "".join(tt.text or "" for tt in is_.iterfind(".//a:t", NS))
                else:
                    cells[ref] = ""

        return cells


def first_nonempty(cells, refs):
    for ref in refs:
        val = t(cells.get(norm_ref(ref), ""))
        if val:
            return val
    return ""


def parse_numeric_qty(value):
    s = t(value)
    if not s:
        return None
    try:
        n = float(s)
        return int(n) if n.is_integer() else n
    except Exception:
        return None


def parse_delivery_note(uploaded_file):
    cells = load_cells_from_upload(uploaded_file)

    delivery_note_no = first_nonempty(cells, ["P8", "R7", "R8"])
    delivery_date = serial_to_date(first_nonempty(cells, ["P10", "R9", "K6"]))
    delivery_name = first_nonempty(cells, ["P13", "R12", "R11", "R10"])

    address_block = first_nonempty(cells, ["C9", "D8", "C8", "D9"])
    company, address1, postcode = parse_address_block(address_block)

    rows = []
    for r in range(14, 130):
        stock_code = first_nonempty(cells, [f"B{r}", f"E{r}", f"Q{r}", f"P{r}"])
        qty_raw = first_nonempty(cells, [f"R{r}", f"S{r}", f"W{r}", f"U{r}", f"AA{r}", f"AB{r}", f"C{r}"])
        qty = parse_numeric_qty(qty_raw)

        if not stock_code or qty is None:
            continue

        if re.search(
            r"stock|description|quantity|boxes|uom|unit of measure|delivery note no|delivery date|sales order no|customer order no|account no|carrier|signature|special instructions|booking in tel|pallets|vat registration|company registration",
            stock_code,
            re.I,
        ):
            continue

        rows.append({
            "Owner": "Opalion",
            "Additional Ref": delivery_note_no,
            "Delivery Company Name": company,
            "Delivery Address 1": address1,
            "Postcode": postcode,
            "Delivery Name": delivery_name,
            "Delivery Date": delivery_date,
            "Order Type": "Sales",
            "Product Code": apostrophe_if_leading_zero(stock_code),
            "Quantity": qty,
        })

    return rows


st.title(APP_TITLE)
st.caption("Upload one or more delivery notes, preview the combined rows, and download one CSV file.")

uploaded_files = st.file_uploader(
    "Choose delivery note files",
    type=["xlsx"],
    accept_multiple_files=True,
)

if "rows" not in st.session_state:
    st.session_state.rows = []

if uploaded_files:
    all_rows = []
    errors = []
    for f in uploaded_files:
        try:
            all_rows.extend(parse_delivery_note(f))
        except Exception as e:
            errors.append(f"{f.name}: {e}")

    st.session_state.rows = all_rows

    st.subheader("Preview")
    if all_rows:
        preview_df = pd.DataFrame(all_rows)[OUTPUT_FIELDS]
        st.dataframe(preview_df, use_container_width=True, height=500)
    else:
        st.info("No line items were found in the uploaded files.")

    if errors:
        st.warning("Some files could not be read:\n" + "\n".join(errors))

    output = io.StringIO()
    if all_rows:
        pd.DataFrame(all_rows)[OUTPUT_FIELDS].to_csv(output, index=False)
        st.download_button(
            "Download combined CSV",
            data=output.getvalue().encode("utf-8-sig"),
            file_name="Opalion_Order_Import_Combined.csv",
            mime="text/csv",
        )
else:
    st.info("Add one or more delivery notes to see the preview and download button.")
