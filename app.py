import streamlit as st
import anthropic
import openpyxl
from openpyxl.styles import Font
import io
import re
import os
import base64

st.set_page_config(page_title="GWS Quote", page_icon="🏠", layout="wide")

st.markdown("""
<style>
html, body, [data-testid="stAppViewContainer"], [data-testid="stMain"] { background: #ffffff; color: #1a1a1a; }
[data-testid="stAppViewContainer"] { background: #ffffff; }
[data-testid="stHeader"] { background: #ffffff; }
.header-row { display: flex; align-items: center; gap: 1rem; margin-bottom: 1.5rem; padding-bottom: 1rem; border-bottom: 1px solid #e0e0e0; }
.header-title { font-size: 1.8rem; font-weight: 600; color: #1a1a1a; margin: 0; }
.section-label { font-size: 0.85rem; font-weight: 600; color: #444; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 0.4rem; }
.info-box { background: #f0f4ff; border: 1px solid #c8d4f0; border-left: 3px solid #1a3a8a; border-radius: 4px; padding: 0.8rem 1rem; color: #1a3a8a; margin-bottom: 1rem; font-size: 0.875rem; line-height: 1.5; }
.preview-panel { background: #f8f9fa; border: 1px solid #e0e0e0; border-radius: 6px; padding: 1.5rem; min-height: 400px; font-size: 0.875rem; line-height: 1.7; color: #333; white-space: pre-wrap; font-family: monospace; }
.preview-placeholder { background: #f8f9fa; border: 1px solid #e0e0e0; border-radius: 6px; padding: 1.5rem; min-height: 400px; display: flex; align-items: center; justify-content: center; color: #aaa; font-style: italic; font-size: 0.9rem; text-align: center; }
.error-box { background: #fff5f5; border: 1px solid #fcc; border-left: 3px solid #e05555; border-radius: 4px; padding: 0.8rem 1rem; color: #c0392b; margin: 0.5rem 0; font-size: 0.875rem; }
.stTextArea textarea { background: #ffffff !important; border: 1px solid #d0d0d0 !important; color: #1a1a1a !important; border-radius: 4px !important; font-size: 0.9rem !important; }
.stTextArea textarea:focus { border-color: #1a3a8a !important; box-shadow: 0 0 0 1px #1a3a8a20 !important; }
.stButton > button { background: #12254a; color: #ffffff; border: none; font-weight: 600; font-size: 0.9rem; padding: 0.55rem 1.5rem; border-radius: 4px; width: 100%; transition: background 0.2s; }
.stButton > button:hover { background: #1a3a8a; }
[data-testid="stDownloadButton"] > button { background: #1a5c2a; color: #ffffff; border: none; font-weight: 600; border-radius: 4px; width: 100%; }
[data-testid="stDownloadButton"] > button:hover { background: #236b33; }
</style>
""", unsafe_allow_html=True)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def get_logo_b64():
    for fname in ["GWS Roofing Logo.png", "GWS_Roofing_Logo.png", "GWS_Roofing_Logo.jpg", "GWS Roofing Logo.jpg"]:
        p = os.path.join(BASE_DIR, fname)
        if os.path.exists(p):
            ext = os.path.splitext(fname)[1].lower()
            mime = "image/png" if ext == ".png" else "image/jpeg"
            with open(p, "rb") as f:
                return base64.b64encode(f.read()).decode(), mime
    return None, None

logo_b64, logo_mime = get_logo_b64()

if logo_b64:
    st.markdown(
        f'<div class="header-row"><img src="data:{logo_mime};base64,{logo_b64}" style="height:60px;width:auto;"><span class="header-title">Quote</span></div>',
        unsafe_allow_html=True
    )
else:
    st.markdown('<div class="header-row"><span class="header-title">GWS Quote</span></div>', unsafe_allow_html=True)

SYSTEM_PROMPT = """You are a quote parser for GWS Roofing. Extract structured data from dictated roofing job information and return JSON only - no explanation, no preamble, no markdown fences.

RETURN FORMAT (strict JSON):
{
  "estimator": "string",
  "date": "DD/MM/YYYY",
  "customer_name": "string",
  "address": "string",
  "postcode": "string",
  "guarantee": "string or null",
  "notes": "string or null",
  "items": [
    {"type": "subheading or item", "number": "null or integer", "description": "string", "cost": "null or number"}
  ]
}

RULES:
- All text to proper sentence case.
- Date to DD/MM/YYYY format.
- Cost: only valid if pound symbol or word pounds present. Numbers in descriptions are NOT costs.
- If item has description but no valid cost, return {"error": "missing_cost", "item": description}.
- Subheadings: type=subheading, number=null, cost=null.
- Items: type=item, number=integer, cost=number.
- guarantee and notes: only if explicitly dictated, else null.
- Do not carry costs between items.
- Return ONLY valid JSON."""

TEMPLATE_1 = os.path.join(BASE_DIR, "GWS_Quote_-_1_Page.xlsx")
TEMPLATE_2 = os.path.join(BASE_DIR, "GWS_Quote_-_2_Page_.xlsx")

def sentence_case(text):
    if not text:
        return text
    return text[0].upper() + text[1:]

def split_description(description, max_chars=80):
    words = description.split()
    lines = []
    current = ""
    for word in words:
        test = (current + " " + word).strip() if current else word
        if len(test) <= max_chars:
            current = test
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines

def parse_with_claude(dictation):
    import json
    client = anthropic.Anthropic()
    response = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=2000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": dictation}]
    )
    raw = response.content[0].text.strip()
    raw = re.sub(r"^```json\s*|^```\s*|\s*```$", "", raw, flags=re.MULTILINE).strip()
    return json.loads(raw)

def calculate_rows_needed(items):
    row = 9
    prev_was_item = False
    for item in items:
        if item["type"] == "subheading":
            if prev_was_item:
                row += 1
            row += 1
            prev_was_item = False
        else:
            lines = split_description(item["description"])
            row += len(lines)
            row += 1
            prev_was_item = True
    return row - 1

def write_quote_to_excel(data):
    rows_needed = calculate_rows_needed(data["items"])
    use_2page = rows_needed > 34
    template_path = TEMPLATE_2 if use_2page else TEMPLATE_1
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    guarantee_row = 91 if use_2page else 40
    notes_row = 92 if use_2page else 41
    ws["A1"] = f"Estimator :  {sentence_case(data['estimator'])}"
    ws["C3"] = data["date"]
    ws["C4"] = sentence_case(data["customer_name"])
    ws["C5"] = sentence_case(data["address"])
    ws["C6"] = data["postcode"].upper()
    current_row = 9
    prev_was_item = False
    for item in data["items"]:
        if item["type"] == "subheading":
            if prev_was_item:
                current_row += 1
            cell = ws.cell(row=current_row, column=2)
            cell.value = sentence_case(item["description"])
            cell.font = Font(bold=True, name=cell.font.name or "Arial")
            current_row += 1
            prev_was_item = False
        else:
            lines = split_description(item["description"])
            for i, line in enumerate(lines):
                if i == 0:
                    ws.cell(row=current_row, column=1).value = item["number"]
                ws.cell(row=current_row, column=2).value = sentence_case(line) if i == 0 else line
                if i == len(lines) - 1:
                    ws.cell(row=current_row, column=7).value = item["cost"]
                current_row += 1
            current_row += 1
            prev_was_item = True
    if data.get("guarantee"):
        ws.cell(row=guarantee_row, column=3).value = sentence_case(data["guarantee"])
    if data.get("notes"):
        ws.cell(row=notes_row, column=3).value = sentence_case(data["notes"])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), use_2page

def build_preview(data):
    lines = []
    lines.append(f"Estimator : {data['estimator']}")
    lines.append(f"Date      : {data['date']}")
    lines.append(f"Customer  : {data['customer_name']}")
    lines.append(f"Address   : {data['address']}")
    lines.append(f"Postcode  : {data['postcode']}")
    if data.get("guarantee"):
        lines.append(f"Guarantee : {data['guarantee']}")
    if data.get("notes"):
        lines.append(f"Notes     : {data['notes']}")
    lines.append("")
    lines.append("-" * 50)
    for item in data["items"]:
        if item["type"] == "subheading":
            lines.append(f"\n  > {item['description'].upper()}")
        else:
            cost_str = f"£{item['cost']:,.2f}" if item["cost"] else "NO COST"
            lines.append(f"  Item {item['number']} - {item['description']} - {cost_str}")
    return "\n".join(lines)

for key in ["parsed_data", "preview_text", "excel_bytes", "filename", "error"]:
    if key not in st.session_state:
        st.session_state[key] = None

left_col, right_col = st.columns([1, 1], gap="large")

with left_col:
    st.markdown('<div class="section-label">Quote Details</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="info-box">Select the estimator, then begin your dictation with '
        '<strong>New quote</strong>. Include date, customer name, address, postcode and items. '
        'Costs must include £ or the word "pounds".</div>',
        unsafe_allow_html=True
    )
    st.markdown('<div class="section-label">Estimator</div>', unsafe_allow_html=True)
    estimator_name = st.selectbox(
        "Estimator",
        ["Gary Sparrowhawk", "Joe Sparrowhawk", "Gary Dolling", "Sam Baldwin"],
        label_visibility="collapsed"
    )
    st.markdown('<div class="section-label" style="margin-top:0.8rem;">Dictation</div>', unsafe_allow_html=True)
    dictation = st.text_area(
        "Dictation",
        height=260,
        placeholder=(
            "New quote\n"
            "Date 17th February 2026\n"
            "Name of customer Mr A Jones\n"
            "Address 14 High Street, Guildford\n"
            "Postcode GU1 3AA\n\n"
            "Item 1 Strip and re-tile front roof slope in concrete interlocking tiles £2500\n"
            "Next item Replace all lead flashings to chimney stack £850"
        ),
        label_visibility="collapsed"
    )
    btn_col1, btn_col2 = st.columns([3, 1])
    with btn_col1:
        read_btn = st.button("Process with AI", use_container_width=True)
    with btn_col2:
        if st.button("Reset", use_container_width=True):
            for key in ["parsed_data", "preview_text", "excel_bytes", "filename", "error"]:
                st.session_state[key] = None
            st.rerun()
    if st.session_state.error:
        st.markdown(f'<div class="error-box">Warning: {st.session_state.error}</div>', unsafe_allow_html=True)

with right_col:
    st.markdown('<div class="section-label">Preview</div>', unsafe_allow_html=True)
    if st.session_state.preview_text:
        st.markdown(f'<div class="preview-panel">{st.session_state.preview_text}</div>', unsafe_allow_html=True)
        if st.session_state.excel_bytes:
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label=f"Download {st.session_state.filename}",
                data=st.session_state.excel_bytes,
                file_name=st.session_state.filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.markdown('<div class="preview-placeholder">Your quote preview will appear here</div>', unsafe_allow_html=True)

if read_btn and dictation.strip():
    with st.spinner("Processing with AI..."):
        try:
            data = parse_with_claude(dictation)
            if "error" in data:
                st.session_state.error = f"Missing cost for item: {data.get('item', 'unknown')}"
                st.session_state.parsed_data = None
            else:
                data["estimator"] = estimator_name
                missing = [f for f in ["date", "customer_name", "address", "postcode"] if not data.get(f)]
                if missing:
                    st.session_state.error = f"Missing required fields: {', '.join(missing)}"
                else:
                    st.session_state.parsed_data = data
                    st.session_state.preview_text = build_preview(data)
                    st.session_state.error = None
                    excel_bytes, used_2page = write_quote_to_excel(data)
                    address = data.get("address", "Quote")
                    safe_addr = re.sub(r"[^\w\s-]", "", address).strip()
                    safe_addr = re.sub(r"\s+", " ", safe_addr)
                    st.session_state.filename = f"GWS Quote {safe_addr}.xlsx"
                    st.session_state.excel_bytes = excel_bytes
        except Exception as e:
            st.session_state.error = f"Error: {str(e)}"
    st.rerun()
