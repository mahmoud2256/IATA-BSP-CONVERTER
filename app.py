import html
import io
import os

import streamlit as st

st.set_page_config(
    page_title="IATA BSP Converter PRO - Developed by Mahmoud Amin",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

CSS = """
<style>
header[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"],
#MainMenu, footer, [data-testid="InputInstructions"] {display:none !important;}
.stApp {background:#0F172A; font-family:"Segoe UI",sans-serif;}
.block-container {padding:0 !important; max-width:100% !important;}
.hdr {background:#1E3A8A; height:120px; text-align:center; padding-top:18px; box-sizing:border-box;}
.hdr h1 {color:#fff; font-size:28px; font-weight:700; margin:0; padding:0; line-height:1.3;}
.hdr .s {color:#BFDBFE; font-size:13px; margin-top:2px;}
.hdr .d {color:#93C5FD; font-size:11px; font-style:italic; margin-top:2px;}
div[data-testid="stHorizontalBlock"] {padding:30px 40px; gap:20px; align-items:stretch !important;}
div[data-testid="stColumn"], div[data-testid="column"] {background:#1E293B; border:1px solid #475569;}
div[data-testid="stColumn"] > div, div[data-testid="column"] > div {padding:30px;}
.step {font-size:12px; font-weight:700; margin-top:6px;}
.ttl {font-size:19px; font-weight:700; color:#F1F5F9; margin:0 0 15px;}
.status {color:#94A3B8; font-size:14px; margin-top:6px; word-break:break-word;}
.ftitle {color:#3B82F6; font-size:13px; font-weight:700; margin-bottom:15px;}
.info {background:#334155; color:#F1F5F9; font-family:Consolas,monospace; font-size:13px;
       padding:14px; min-height:440px; white-space:pre-wrap;}
.foot {text-align:center; color:#94A3B8; font-size:12px; padding:0 0 24px;}
/* inputs */
div[data-baseweb="input"], div[data-baseweb="base-input"] {background:#334155 !important; border:0 !important; border-radius:0 !important;}
.stTextInput input {background:#334155 !important; color:#fff !important; padding:14px 12px !important; font-size:14px;}
/* uploader */
[data-testid="stFileUploader"] section {background:#334155 !important; border:0 !important; border-radius:0 !important; padding:12px;}
[data-testid="stFileUploader"] section * {color:#94A3B8 !important;}
[data-testid="stFileUploader"] section button {background:#3B82F6 !important; border:0 !important; border-radius:0 !important;
    font-size:0 !important; padding:10px 22px !important;}
[data-testid="stFileUploader"] section button::after {content:"IMPORT PDF"; color:#fff; font-size:14px; font-weight:700;}
[data-testid="stFileUploaderFile"] *, [data-testid="stFileUploader"] small {color:#F1F5F9 !important;}
/* buttons */
[data-testid="stButton"] button, .stButton > button {background:#10B981 !important; border:0 !important; border-radius:0 !important;
    width:100%; padding:18px 0 !important; margin:10px 0 14px;}
[data-testid="stButton"] button:hover {background:#059669 !important;}
[data-testid="stButton"] button p, .stButton > button p {color:#fff !important; font-size:18px; font-weight:700;}
[data-testid="stDownloadButton"] button, .stDownloadButton > button {background:#3B82F6 !important; border:0 !important;
    border-radius:0 !important; width:100%; padding:12px 0 !important;}
[data-testid="stDownloadButton"] button:hover {background:#2563EB !important;}
[data-testid="stDownloadButton"] button p {color:#fff !important; font-size:15px; font-weight:700;}
.stSpinner, .stSpinner * {color:#94A3B8 !important;}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

st.markdown(
    '<div class="hdr"><h1>IATA / BSP CONVERTER</h1>'
    '<div class="s">FCAGBILLDET PDF → EXCEL • PROFESSIONAL EDITION</div>'
    '<div class="d">Developed by Mahmoud Amin</div></div>',
    unsafe_allow_html=True,
)

DEFAULT_INFO = (
    "No file selected.\n\nSupported file:\n• EG_FCAGBILLDET_*.PDF\n• Any BSP Billing PDF\n\n"
    "Output:\n• Excel with 1 sheet\n• IATA (all transactions)\n\n© 2026 Mahmoud Amin"
)
S = st.session_state
S.setdefault("status", "Ready - Select your FCAGBILLDET PDF file")
S.setdefault("result", None)
S.setdefault("last_file", None)
S.setdefault("out_name", "")


def step(n, color, title):
    st.markdown(f'<div class="step" style="color:{color}">STEP {n}</div><div class="ttl">{title}</div>',
                unsafe_allow_html=True)


left, right = st.columns([2.7, 1])

with left:
    step(1, "#3B82F6", "Select PDF File")
    up = st.file_uploader("pdf", type=["pdf"], label_visibility="collapsed")
    key = (up.name, up.size) if up else None
    if key != S.last_file:
        S.last_file, S.result = key, None
        if up:
            S.out_name = os.path.splitext(up.name)[0] + ".xlsx"
            S.status = f"✓ Selected: {up.name} | Size: {up.size // 1024} KB"
        else:
            S.out_name = ""
            S.status = "Ready - Select your FCAGBILLDET PDF file"

    step(2, "#10B981", "Save Excel As")
    st.text_input("out", key="out_name", placeholder="BSP_Report.xlsx", label_visibility="collapsed")

    if st.button("⚡ GENERATE EXCEL NOW", key="go"):
        if up is None:
            S.status = "❌ Please select a PDF file first!"
        else:
            out = (S.out_name.strip() or "BSP_Report.xlsx")
            if not out.lower().endswith(".xlsx"):
                out += ".xlsx"
            with st.spinner("⏳ Converting PDF to Excel... Please wait, this may take a few seconds"):
                try:
                    import pdfplumber  # heavy imports only when converting
                    import iata_to_excel as conv
                    with pdfplumber.open(io.BytesIO(up.getvalue())) as pdf:
                        raw = conv.parse_detail_pages(pdf)
                    rows = conv.finalize_rows(raw)
                    buf = io.BytesIO()
                    conv.write_excel(rows, [], buf)
                    S.result = {"rows": len(rows), "xlsx": buf.getvalue(), "name": out,
                                "txn": sum(r["Transaction Amount"] or 0 for r in rows),
                                "bal": sum(r["Balance Payable"] or 0 for r in rows)}
                    S.status = f"✅ Success! {len(rows)} tickets converted - {out}"
                except Exception as e:
                    S.result = None
                    S.status = f"❌ Error: {e}"

    st.markdown(f'<div class="status">{html.escape(S.status)}</div>', unsafe_allow_html=True)
    R = S.result
    if R:
        st.download_button("⬇️ DOWNLOAD EXCEL", R["xlsx"], file_name=R["name"],
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with right:
    if up is None:
        info = DEFAULT_INFO
    else:
        info = f"File: {up.name}\nSize: {up.size // 1024} KB\n\n"
        R = S.result
        if R:
            info += (f"Rows: {R['rows']}\nTransaction Amount: {R['txn']:,.2f}\n"
                     f"Balance Payable: {R['bal']:,.2f}\n\nOutput:\n• Excel with 1 sheet (IATA)")
        else:
            info += "Ready to convert!\n\nClick GENERATE to create Excel file with all billing details."
    st.markdown(f'<div class="ftitle">FILE INFO</div><div class="info">{html.escape(info)}</div>',
                unsafe_allow_html=True)

st.markdown('<div class="foot">© 2026 Mahmoud Amin - IATA BSP Converter PRO v2.0 - All Rights Reserved</div>',
            unsafe_allow_html=True)
