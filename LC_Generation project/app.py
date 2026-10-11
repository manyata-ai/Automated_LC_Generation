"""
Leaving Certificate (LC) Generation System
-------------------------------------------
Single-page Streamlit app for school office staff.

Flow:
    1. Upload a photo of a filled admission form
    2. Auto-detect fields (YOLOv8) + extract text (EasyOCR, Marathi + English)
    3. Review / correct the extracted values in an editable form
    4. Generate the Leaving Certificate as a .docx (exact school template)
"""

import io
import os
import re
import shutil
import subprocess
import tempfile
import streamlit as st

# ---------------------------------------------------------------------------
# Page config (must be the first Streamlit call)
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="LC Generation System | ज्ञानदीप विद्यालय",
    page_icon="📄",
    layout="centered",
)

# ---------------------------------------------------------------------------
# Field definitions — order, keys and Marathi labels (matches the LC template)
# ---------------------------------------------------------------------------
FIELDS = [
    ("student_name", "विद्यार्थ्याचे पूर्ण नाव"),
    ("dob", "जन्मतारीख"),
    ("gender", "लिंग"),
    ("class", "इयत्ता"),
    ("admission_no", "प्रवेश क्रमांक"),
    ("admission_date", "प्रवेश दिनांक"),
    ("blood_group", "रक्तगट"),
    ("mother_tongue", "मातृभाषा"),
    ("father_name", "वडिलांचे नाव"),
    ("mother_name", "आईचे नाव"),
    ("parents_mobile_no", "पालकांचा मोबाईल क्रमांक"),
    ("address", "पूर्ण पत्ता"),
    ("previous_school_name", "मागील शाळेचे नाव"),
    ("previous_class", "मागील इयत्ता"),
    ("reason_for_leaving", "शाळा सोडण्याचे कारण"),
]

DATE_NUMBER_FIELDS = {"dob", "admission_date", "class", "previous_class"}
ADMISSION_FIELDS = {"admission_no"}
BLOOD_FIELDS = {"blood_group"}

CROP_PAD = 5
CONF_THRESHOLD = 0.5  # below this (or empty) -> flagged for staff to double check

MODEL_PATH = os.environ.get("LC_MODEL_PATH", os.path.join(os.path.dirname(__file__), "models", "best.pt"))

# ---------------------------------------------------------------------------
# Styling — light background, single accent color, Devanagari-friendly font
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Devanagari:wght@400;600;700&family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"]  {
        font-family: 'Inter', 'Noto Sans Devanagari', sans-serif;
    }

    .main .block-container {
        max-width: 760px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    /* Header */
    .lc-header {
        text-align: center;
        margin-bottom: 0.25rem;
    }
    .lc-header h1 {
        font-size: 1.5rem;
        font-weight: 700;
        color: #1C1C1C;
        margin-bottom: 0;
    }
    .lc-header p {
        color: #6B6B63;
        font-size: 0.95rem;
        margin-top: 0.15rem;
    }

    /* Stepper */
    .stepper {
        display: flex;
        justify-content: center;
        gap: 0;
        margin: 1.75rem 0 2rem 0;
    }
    .step {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        padding: 0.4rem 1rem;
        font-size: 0.85rem;
        font-weight: 600;
        color: #A8A89E;
        border-bottom: 3px solid #E5E5DE;
    }
    .step.active {
        color: #1F6F54;
        border-bottom: 3px solid #1F6F54;
    }
    .step.done {
        color: #1F6F54;
        border-bottom: 3px solid #B9DBCB;
    }
    .step-num {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 20px;
        height: 20px;
        border-radius: 50%;
        background: #E5E5DE;
        color: #6B6B63;
        font-size: 0.75rem;
    }
    .step.active .step-num, .step.done .step-num {
        background: #1F6F54;
        color: white;
    }

    /* Flag for uncertain / empty fields */
    .flag-warning {
        color: #B3261E;
        font-size: 0.8rem;
        font-weight: 500;
        margin-bottom: -0.6rem;
    }
    div[data-testid="stTextInput"] input.flagged-input {
        border-color: #B3261E !important;
    }

    /* Success banner */
    .success-box {
        background: #EAF5EF;
        border: 1px solid #B9DBCB;
        border-radius: 8px;
        padding: 1.25rem;
        text-align: center;
        margin-bottom: 1.25rem;
    }
    .success-box h3 {
        color: #1F6F54;
        margin-bottom: 0.25rem;
    }

    footer {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="lc-header">
        <h1>ज्ञानदीप विद्यालय, पुणे</h1>
        <p>Leaving Certificate Generation System</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------
if "step" not in st.session_state:
    st.session_state.step = 1  # 1=upload, 2=review, 3=download
if "extracted" not in st.session_state:
    st.session_state.extracted = {}
if "flags" not in st.session_state:
    st.session_state.flags = {}
if "docx_bytes" not in st.session_state:
    st.session_state.docx_bytes = None
if "pdf_bytes" not in st.session_state:
    st.session_state.pdf_bytes = None
if "pdf_error" not in st.session_state:
    st.session_state.pdf_error = None
if "uploaded_image_bytes" not in st.session_state:
    st.session_state.uploaded_image_bytes = None
if "pipeline_error" not in st.session_state:
    st.session_state.pipeline_error = None


def render_stepper():
    labels = ["Upload", "Review / Edit", "Download"]
    parts = []
    for i, label in enumerate(labels, start=1):
        cls = "step"
        if i == st.session_state.step:
            cls += " active"
        elif i < st.session_state.step:
            cls += " done"
        parts.append(
            f'<div class="{cls}"><span class="step-num">{i}</span>{label}</div>'
        )
    st.markdown(f'<div class="stepper">{"".join(parts)}</div>', unsafe_allow_html=True)


render_stepper()

# ---------------------------------------------------------------------------
# Model / OCR loading (cached so it only loads once per server process)
# ---------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def load_yolo_model(model_path: str):
    from ultralytics import YOLO
    if not os.path.exists(model_path):
        raise FileNotFoundError(
            f"Trained YOLOv8 weights not found at '{model_path}'. "
            "Place the trained best.pt file in the models/ folder."
        )
    return YOLO(model_path)


@st.cache_resource(show_spinner=False)
def load_ocr_reader():
    import easyocr
    return easyocr.Reader(["mr", "en"])


def clean_text(text: str, strip_latin_prefix: bool = False, strip_label_fragment: bool = False) -> str:
    text = re.sub(r"^[०.:ः\s]+", "", text)
    if strip_latin_prefix:
        # OCR sometimes misreads a stray mark/table-border artifact as a
        # leading Latin letter or digit (e.g. "O विद्यार्थ्याचे नाव..." ->
        # actual name). Marathi name/address/etc. fields should start with
        # Devanagari text, so drop any such junk prefix.
        text = re.sub(r"^[A-Za-z0-9.\-:]+(?=\s*[\u0900-\u097F])", "", text).strip()
    if strip_label_fragment:
        # Crop padding sometimes pulls in the tail end of the printed label
        # to the left of the value (e.g. "नाव : अभिज्ञान..." or "पत्ता : सदनिका...").
        # If a colon-like character shows up early in the text, treat
        # everything before it as leftover label and keep only what follows.
        match = re.match(r"^.{0,18}?[:ः]\s*(.+)$", text)
        if match and match.group(1).strip():
            text = match.group(1).strip()
    return text.strip()


def extract_pattern(text: str, pattern: str) -> str:
    """Pulls out the substring matching `pattern` from `text`, discarding
    any stray leading/trailing tokens picked up by wider crop padding
    (e.g. "५ A२०२७०५२७" -> "A२०२७०५२७"). Falls back to the original
    (whitespace-collapsed) text if no match is found, so nothing is lost."""
    match = re.search(pattern, text)
    if match:
        return match.group(0)
    return text


def fix_date_format(date_str: str) -> str:
    parts = date_str.split("-")
    if len(parts) == 3:
        year, month, day = parts
        marathi_to_eng = str.maketrans("०१२३४५६७८९", "0123456789")
        if len(month.translate(marathi_to_eng)) == 1:
            month = "०" + month
        if len(day.translate(marathi_to_eng)) == 1:
            day = "०" + day
        return f"{year}-{month}-{day}"
    return date_str


def run_pipeline(image_path: str):
    """Runs YOLOv8 detection + EasyOCR extraction on the uploaded image.

    Returns (extracted_data, flags) where flags[field] = True if the value
    is empty or low-confidence and should be double-checked by staff.
    """
    import cv2

    model = load_yolo_model(MODEL_PATH)
    reader = load_ocr_reader()

    # Lower confidence catches small/narrow fields (blood_group, previous_class,
    # admission_no) more reliably; the OCR-confidence flag afterwards still
    # tells staff which values to double-check, so this is a safe trade-off.
    results = model.predict(source=image_path, conf=0.15, verbose=False)
    img = cv2.imread(image_path)
    if img is None:
        raise ValueError("Could not read the uploaded image.")

    extracted_data = {key: "" for key, _ in FIELDS}
    flags = {key: True for key, _ in FIELDS}  # default: flagged until detected

    # If YOLO fires multiple boxes for the same field, keep only the
    # highest-confidence one instead of overwriting with whichever came last.
    best_conf_seen = {}

    for box in results[0].boxes:
        cls_id = int(box.cls[0])
        field_name = model.names[cls_id]
        if field_name not in extracted_data:
            continue

        det_conf = float(box.conf[0])
        if field_name in best_conf_seen and det_conf <= best_conf_seen[field_name]:
            continue

        x1, y1, x2, y2 = map(int, box.xyxy[0])
        # Keep left padding minimal -- the printed label sits immediately to
        # the left of the value on these forms, so widening left padding
        # pulls in leftover label text (e.g. "नाव :", "पत्ता :"). Right/top/
        # bottom can afford a bit more room for narrow boxes (blood group,
        # short numbers) without that risk.
        pad_left = CROP_PAD
        pad_right = max(CROP_PAD, int(0.12 * (x2 - x1)))
        pad_y = max(CROP_PAD, int(0.10 * (y2 - y1)))
        x1, y1 = max(0, x1 - pad_left), max(0, y1 - pad_y)
        x2, y2 = min(img.shape[1], x2 + pad_right), min(img.shape[0], y2 + pad_y)
        cropped = img[y1:y2, x1:x2]
        if cropped.size == 0:
            continue
        resized = cv2.resize(cropped, None, fx=8, fy=8, interpolation=cv2.INTER_CUBIC)

        if field_name in DATE_NUMBER_FIELDS:
            raw = reader.readtext(
                resized, detail=1, allowlist="०१२३४५६७८९-",
                decoder="beamsearch", beamWidth=10,
            )
        elif field_name in ADMISSION_FIELDS:
            raw = reader.readtext(
                resized, detail=1, allowlist="A०१२३४५६७८९-",
                decoder="beamsearch", beamWidth=10,
            )
        elif field_name in BLOOD_FIELDS:
            raw = reader.readtext(
                resized, detail=1, allowlist="ABO+-",
                decoder="beamsearch", beamWidth=10,
            )
        else:
            raw = reader.readtext(resized, detail=1, decoder="beamsearch", beamWidth=10)

        raw.sort(key=lambda r: r[0][0][0])
        text = " ".join([r[1] for r in raw])
        avg_conf = (sum(r[2] for r in raw) / len(raw)) if raw else 0.0

        if field_name in ("dob", "admission_date"):
            value = extract_pattern(text, r"[०-९]{1,4}-[०-९]{1,2}-[०-९]{1,2}")
        elif field_name in ("class", "previous_class"):
            value = extract_pattern(text, r"[०-९]{1,2}")
        elif field_name in ADMISSION_FIELDS:
            value = extract_pattern(text, r"A?[०-९]{4,}")
        elif field_name in BLOOD_FIELDS:
            value = extract_pattern(text, r"(AB|A|B|O)[+-]")
        else:
            value = clean_text(text, strip_latin_prefix=True, strip_label_fragment=True)

        extracted_data[field_name] = value.strip()
        flags[field_name] = (not text.strip()) or (avg_conf < CONF_THRESHOLD)
        best_conf_seen[field_name] = det_conf

    if extracted_data.get("admission_date"):
        extracted_data["admission_date"] = fix_date_format(
            re.sub(r"-{2,}", "-", extracted_data["admission_date"])
        )
    if extracted_data.get("dob"):
        extracted_data["dob"] = fix_date_format(re.sub(r"-{2,}", "-", extracted_data["dob"]))

    return extracted_data, flags


# ---------------------------------------------------------------------------
# DOCX generation — exact school template, unchanged wording / order
# ---------------------------------------------------------------------------
def generate_lc(data: dict) -> bytes:
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()

    heading = doc.add_heading("ज्ञानदीप विद्यालय, पुणे", level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph("शाळा सोडल्याचा दाखला (Leaving Certificate)")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.runs[0].bold = True
    sub.runs[0].font.size = Pt(14)

    doc.add_paragraph("")  # spacing

    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"

    for key, label in FIELDS:
        row = table.add_row()
        row.cells[0].text = label
        row.cells[1].text = data.get(key, "")
        row.cells[0].paragraphs[0].runs[0].bold = True

    doc.add_paragraph("")
    doc.add_paragraph("")

    sign_para = doc.add_paragraph()
    sign_para.add_run("पालकांची सही" + " " * 50 + "मुख्याध्यापकांची सही")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ---------------------------------------------------------------------------
# PDF generation — converts the already-correct DOCX to PDF using Word or
# LibreOffice, whichever is installed. We deliberately do NOT draw the PDF
# from scratch (e.g. with reportlab): general-purpose PDF drawing libraries
# don't perform the complex-script shaping Devanagari needs (reordering
# vowel signs / matras around consonants), so text like "विद्यालय" comes
# out garbled as "वदि्यालय". Word and LibreOffice already shape Marathi
# text correctly, so converting the finished .docx sidesteps the problem.
# ---------------------------------------------------------------------------
def generate_lc_pdf(docx_bytes: bytes) -> bytes:
    tmp_dir = tempfile.mkdtemp(prefix="lc_pdf_")
    docx_path = os.path.join(tmp_dir, "lc.docx")
    pdf_path = os.path.join(tmp_dir, "lc.pdf")

    try:
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)

        converted = False

        # Preferred on Windows/Mac: MS Word via docx2pdf (needs Word installed)
        try:
            from docx2pdf import convert as docx2pdf_convert
            docx2pdf_convert(docx_path, pdf_path)
            converted = os.path.exists(pdf_path)
        except Exception:
            converted = False

        # Fallback: LibreOffice headless, if installed
        if not converted:
            soffice = shutil.which("soffice") or shutil.which("soffice.exe")
            if not soffice:
                raise RuntimeError(
                    "PDF conversion needs Microsoft Word or LibreOffice installed on this "
                    "computer, and neither could be found. Please install one of them, or "
                    "use the .docx download and save/export it as PDF from Word directly."
                )
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp_dir, docx_path],
                check=True, timeout=120, capture_output=True,
            )
            converted = os.path.exists(pdf_path)

        if not converted:
            raise RuntimeError("PDF file was not created during conversion.")

        with open(pdf_path, "rb") as f:
            return f.read()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# STEP 1 — Upload
# ---------------------------------------------------------------------------
if st.session_state.step == 1:
    st.write("Upload a clear photo of the filled admission form to auto-fill the certificate.")

    uploaded_file = st.file_uploader(
        "प्रवेश अर्जाचा फोटो अपलोड करा (Upload admission form photo)",
        type=["jpg", "jpeg", "png"],
    )

    if uploaded_file is not None:
        st.session_state.uploaded_image_bytes = uploaded_file.getvalue()
        st.image(st.session_state.uploaded_image_bytes, caption="Uploaded form", use_container_width=True)

        if st.button("Detect & Extract", type="primary", use_container_width=True):
            with st.spinner("Detecting fields and reading text — this may take a moment..."):
                tmp_path = None
                try:
                    suffix = os.path.splitext(uploaded_file.name)[1] or ".jpg"
                    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                        tmp.write(st.session_state.uploaded_image_bytes)
                        tmp_path = tmp.name

                    extracted, flags = run_pipeline(tmp_path)
                    st.session_state.extracted = extracted
                    st.session_state.flags = flags
                    st.session_state.pipeline_error = None
                    st.session_state.step = 2
                except Exception as exc:
                    st.session_state.pipeline_error = str(exc)
                    # Fall back to a blank editable form so staff can still proceed
                    st.session_state.extracted = {key: "" for key, _ in FIELDS}
                    st.session_state.flags = {key: True for key, _ in FIELDS}
                    st.session_state.step = 2
                finally:
                    if tmp_path and os.path.exists(tmp_path):
                        os.remove(tmp_path)
            st.rerun()

# ---------------------------------------------------------------------------
# STEP 2 — Review / Edit
# ---------------------------------------------------------------------------
elif st.session_state.step == 2:
    if st.session_state.pipeline_error:
        st.warning(
            "आम्ही फोटोमधून माहिती वाचू शकलो नाही. कृपया अधिक स्पष्ट फोटो अपलोड करा, "
            "किंवा खालील माहिती स्वतः भरा.\n\n"
            "We couldn't read the photo clearly. Please re-upload a clearer image, "
            "or fill in the details below manually."
        )
        with st.expander("Technical details (for troubleshooting)"):
            st.code(st.session_state.pipeline_error)

    st.write("Please review the extracted details below and correct anything that looks wrong. "
             "Fields outlined in red were empty or uncertain.")

    with st.form("review_form"):
        values = {}
        for key, label in FIELDS:
            current_val = st.session_state.extracted.get(key, "")
            is_flagged = st.session_state.flags.get(key, False)
            if is_flagged:
                st.markdown(f'<div class="flag-warning">⚠️ {label} — please verify</div>', unsafe_allow_html=True)

            if key in ("reason_for_leaving", "address"):
                values[key] = st.text_area(label, value=current_val, height=70, key=f"field_{key}")
            else:
                values[key] = st.text_input(label, value=current_val, key=f"field_{key}")

        col1, col2 = st.columns([1, 1])
        with col1:
            back = st.form_submit_button("← Upload a different photo", use_container_width=True)
        with col2:
            submitted = st.form_submit_button("Generate LC", type="primary", use_container_width=True)

    if back:
        st.session_state.step = 1
        st.session_state.pipeline_error = None
        st.rerun()

    if submitted:
        st.session_state.extracted = values
        st.session_state.docx_bytes = generate_lc(values)
        try:
            st.session_state.pdf_bytes = generate_lc_pdf(st.session_state.docx_bytes)
            st.session_state.pdf_error = None
        except Exception as exc:
            st.session_state.pdf_bytes = None
            st.session_state.pdf_error = str(exc)
        st.session_state.step = 3
        st.rerun()

# ---------------------------------------------------------------------------
# STEP 3 — Download
# ---------------------------------------------------------------------------
elif st.session_state.step == 3:
    st.markdown(
        """
        <div class="success-box">
            <h3>✅ दाखला तयार आहे</h3>
            <p>The Leaving Certificate has been generated successfully.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    student_name = st.session_state.extracted.get("student_name", "student").strip() or "student"
    safe_name = re.sub(r"[^\w\-]+", "_", student_name)

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="⬇ Download .docx",
            data=st.session_state.docx_bytes,
            file_name=f"LC_{safe_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
    with col2:
        if st.session_state.pdf_bytes:
            st.download_button(
                label="⬇ Download .pdf",
                data=st.session_state.pdf_bytes,
                file_name=f"LC_{safe_name}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True,
            )
        else:
            st.button("⬇ Download .pdf", disabled=True, use_container_width=True)

    if st.session_state.pdf_error:
        with st.expander("PDF could not be generated (for troubleshooting)"):
            st.code(st.session_state.pdf_error)

    if st.button("Start a new certificate", use_container_width=True):
        st.session_state.step = 1
        st.session_state.extracted = {}
        st.session_state.flags = {}
        st.session_state.docx_bytes = None
        st.session_state.pdf_bytes = None
        st.session_state.pdf_error = None
        st.session_state.uploaded_image_bytes = None
        st.session_state.pipeline_error = None
        st.rerun()
