import argparse
import os
import re
import warnings
import pandas as pd
from pypdf import PdfReader, PdfWriter




# Build cross-platform paths relative to this file
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.normpath(
    os.path.join(BASE_DIR, "..", "excel_input", "cross_registration_data.xlsx")
)
TEMPLATE_DIR = os.path.normpath(
    os.path.join(BASE_DIR, "..", "pche_template")
)
OUTPUT_DIR = os.path.normpath(
    os.path.join(BASE_DIR, "..", "Filled_PDFs")
)
os.makedirs(OUTPUT_DIR, exist_ok=True)
warnings.filterwarnings("ignore")




def read_template_pdf(home_school: str) -> PdfReader:
    """Return the PDF template corresponding to the student's home school."""
    if home_school == "pitt":
        return PdfReader(os.path.join(TEMPLATE_DIR, "pche-cross-reg_pitt.pdf"))
    if home_school == "cmu":
        return PdfReader(os.path.join(TEMPLATE_DIR, "pche-cross-reg_cmu.pdf"))
    raise ValueError("No homeschool provided.")


def populate_fields_from_row(row: pd.Series, fields: dict[str, str]) -> None:
    """Populate PDF form fields from a single student row."""
    fields["Home School Student ID"] = row["home_school_id"]
    fields["Sex"] = row["sex"]
    fields["Preferred First Name"] = row["preferred_name"]
    fields["Legal Last Name"] = row["legal_last_name"]
    fields["Legal First Name"] = row["legal_first_name"]
    fields["Middle Initial"] = row["middle_initial"]
    fields["Host School Student ID"] = row["home_school_email"]
    fields["Home School's Email address"] = row["host_school_id"]
    fields["Course Title - Lecture A"] = row["course_code"]
    fields["Course Subject Code - Lecture A"] = row["course_title"]
    fields["Credit/Units - Lecture A"] = row["credits_units"]


def build_writer_with_fields(reader: PdfReader, fields: dict[str, str]) -> PdfWriter:
    """Create a writer, copy pages, and apply field updates."""
    writer = PdfWriter()
    writer.append(reader)
    writer.update_page_form_field_values(writer.pages[0], fields)
    return writer


def sanitize_filename_component(text: str) -> str:
    """Normalize a string for safe filenames."""
    safe = re.sub(r"\s+", "_", text.strip().lower())
    safe = re.sub(r"[^a-z0-9._-]", "", safe)
    return safe or "unknown"


def unique_pdf_name(base_name: str, output_dir: str) -> str:
    """Return a unique PDF filename, appending a counter on collisions."""
    candidate = f"{base_name}.pdf"
    counter = 1
    while os.path.exists(os.path.join(output_dir, candidate)):
        candidate = f"{base_name}-{counter}.pdf"
        counter += 1
    return candidate


def main(output_dir: str = OUTPUT_DIR, excel_path: str = EXCEL_PATH) -> None:
    """Load student data and generate filled PDFs."""
    os.makedirs(output_dir, exist_ok=True)

    data_df = pd.read_excel(excel_path)
    errors: list[tuple[object, str, str, Exception]] = []

    for index, row in data_df.iterrows():
        try:
            # Read the PDF template for this student
            reader = read_template_pdf(row["home_school"])

            # Get the fields from the PDF
            fields = reader.get_form_text_fields()

            # Update the fields
            populate_fields_from_row(row, fields)

            # Create the PDF writer with updated form fields
            writer = build_writer_with_fields(reader, fields)

            # Create the Filled PDF
            base_name = (
                f"{sanitize_filename_component(row['legal_last_name'])}-"
                f"{sanitize_filename_component(row['preferred_name'])}-"
                f"{sanitize_filename_component(row['home_school'])}"
            )
            file_name = unique_pdf_name(base_name, output_dir)
            with open(os.path.join(output_dir, file_name), "wb") as f:
                writer.write(f)
        except Exception as exc:  # noqa: BLE001 - bubble errors with context
            errors.append(
                (
                    index,
                    sanitize_filename_component(row.get("legal_last_name", "")),
                    sanitize_filename_component(row.get("preferred_name", "")),
                    exc,
                )
            )

    if errors:
        print("\nCompleted with errors:")
        for idx, last, pref, exc in errors:
            label = f"Row {idx}"
            if last or pref:
                label += f" ({last}, {pref})"
            print(f"  {label}: {exc}")
        print(f"\nFinished {len(data_df) - len(errors)}/{len(data_df)} rows successfully.")
    else:
        print(f"\nFinished all {len(data_df)} rows successfully.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Generate filled cross-registration PDFs."
    )
    parser.add_argument(
        "--output-dir",
        default=OUTPUT_DIR,
        help="Directory where filled PDFs will be written (default: Filled_PDFs)",
    )
    parser.add_argument(
        "--excel-path",
        default=EXCEL_PATH,
        help=(
            "Path to the cross registration Excel input file "
            "(default: excel_input/cross_registration_data.xlsx)"
        ),
    )
    args = parser.parse_args()
    main(
        output_dir=os.path.normpath(os.path.abspath(args.output_dir)),
        excel_path=os.path.normpath(os.path.abspath(args.excel_path)),
    )
