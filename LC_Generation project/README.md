# LC (Leaving Certificate) Generation System

Single-page Streamlit app for ज्ञानदीप विद्यालय, पुणे office staff.

**Flow:** Upload a photo of a filled admission form → YOLOv8 detects the
15 form fields → EasyOCR (Marathi + English) reads each field → staff
review/correct the values → a Word (.docx) Leaving Certificate is generated
in the school's exact template.

## Setup

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Add your trained YOLOv8 weights:
   - Copy your `best.pt` (trained on the `lc_field_detection` Roboflow
     dataset, per the provided training notebook) into `models/best.pt`.
   - Or set an environment variable to point elsewhere:
     ```bash
     export LC_MODEL_PATH=/path/to/best.pt
     ```

3. Run the app:
   ```bash
   streamlit run app.py
   ```

## Notes

- If the model weights are missing or a photo can't be processed, the app
  shows a friendly bilingual message and falls back to a blank editable
  form so staff can still generate the certificate manually — the app
  never gets stuck.
- Fields that came back empty or with low OCR confidence are flagged with
  a red "please verify" note above the field, so staff know what to
  double-check before generating the certificate.
- The certificate template (school name, subheading, field order/labels,
  table layout, and signature line) is hard-coded to match the school's
  existing format exactly and should not be edited without updating the
  underlying paperwork too.
- `models/` and any uploaded photos are not stored anywhere else — each
  upload is processed in a temporary file that is deleted right after use.
