import streamlit as st
from pathlib import Path
import json
import io
from api_client import extract_invoice
from config import BASE_DIR, JSON_RESULTS_DIR, OUTPUT_DIR
from reports import lenovo_report, meyer_report


UPLOADS_DIR = BASE_DIR / "uploads"
UPLOADS_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


def save_uploaded_file(uploaded) -> Path:
    dest = UPLOADS_DIR / uploaded.name
    with open(dest, "wb") as f:
        f.write(uploaded.getbuffer())
    return dest


def offer_download(path: Path, label: str):
    if not path.exists():
        st.error(f"File not found: {path}")
        return
    with open(path, "rb") as f:
        data = f.read()
    st.download_button(label, data, file_name=path.name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def run_lenovo(json_path: Path):
    with open(json_path, "r", encoding="utf-8") as f:
        json_data = json.load(f)

    rows = lenovo_report.parse_json_to_rows(json_data)
    if not rows:
        st.warning("No rows extracted for Lenovo report.")
        return None

    data_block = json_data.get("result", {}).get("data", [{}])[0]
    original_filename = data_block.get("filename", json_path.stem)
    pdf_name = Path(original_filename).stem
    output_file = OUTPUT_DIR / f"{pdf_name}_lenovo_report.xlsx"
    lenovo_report.create_excel(rows, output_file)
    return output_file


def run_meyer(json_path: Path):
    data = meyer_report.load_json(json_path)

    reference_data = meyer_report.load_reference_excel(meyer_report.REFERENCE_EXCEL_PATH)

    data_block = data.get("result", {}).get("data", [{}])[0]
    original_filename = data_block.get("filename", json_path.stem)
    pdf_name = Path(original_filename).stem

    extracted_data = data_block.get("extracted_data", {})
    root = extracted_data.get("invoice", {}) if "invoice" in extracted_data else extracted_data

    service_locations = root.get("service_locations", [])
    if not service_locations:
        st.warning("No service locations found for Meyer report.")
        return None

    invoice_number = service_locations[0].get("Invoice_number", "")
    json_grand_total = root.get("grand_total")

    location_totals, location_meta = meyer_report.aggregate_by_location(service_locations, reference_data)

    if json_grand_total is None:
        json_grand_total = sum(location_totals.values())

    output_file = OUTPUT_DIR / f"{pdf_name}_meyer_report.xlsx"
    meyer_report.create_excel(location_totals, location_meta, json_grand_total, invoice_number, output_file)
    return output_file


def main():
    st.title("Invoice Pipeline — Extract & Reports")

    st.write("Upload a PDF invoice; the app will extract JSON and generate reports.")

    uploaded = st.file_uploader("Upload invoice PDF", type=["pdf"], accept_multiple_files=False)

    report_choice = st.radio("Report to generate", ("Lenovo", "Meyer", "Both"))

    if uploaded is not None:
        pdf_path = save_uploaded_file(uploaded)
        st.success(f"Saved uploaded PDF: {pdf_path.name}")

        if st.button("Run Extraction and Generate Report"):
            try:
                with st.spinner("Extracting invoice (this may take a while)..."):
                    json_path = extract_invoice(pdf_path)

                st.success(f"Extraction complete: {json_path.name}")

                outputs = []

                if report_choice in ("Lenovo", "Both"):
                    with st.spinner("Generating Lenovo report..."):
                        out = run_lenovo(json_path)
                    if out:
                        st.success(f"Lenovo report created: {out.name}")
                        outputs.append(("Lenovo", out))

                if report_choice in ("Meyer", "Both"):
                    with st.spinner("Generating Meyer report..."):
                        out = run_meyer(json_path)
                    if out:
                        st.success(f"Meyer report created: {out.name}")
                        outputs.append(("Meyer", out))

                if outputs:
                    st.subheader("Download Reports")
                    for label, path in outputs:
                        offer_download(path, f"Download {label} Excel")
                else:
                    st.warning("No reports were generated.")

            except Exception as e:
                st.error(f"Error: {e}")

                


if __name__ == "__main__":
    main()
