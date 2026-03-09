import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from io import BytesIO

# === DEFAULT CONSTANTS ===
DEFAULT_WIDTH_MM = 50
DEFAULT_HEIGHT_MM = 30
FONT_ADJUSTMENT = 2
LABEL_MARGIN = 4

# Built-in fonts
AVAILABLE_FONTS = [
    "Helvetica",
    "Helvetica-Bold",
    "Times-Roman",
    "Times-Bold",
    "Courier",
    "Courier-Bold"
]


# === TEXT WRAPPING ===
def wrap_text(text, font_name, font_size, max_width):
    words = text.split()
    lines = []
    current_line = ""

    for word in words:

        test_line = f"{current_line} {word}".strip()

        width = pdfmetrics.stringWidth(test_line, font_name, font_size)

        if width <= max_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            current_line = word

    if current_line:
        lines.append(current_line)

    return lines


# === FAST FONT SIZE SEARCH (BINARY SEARCH) ===
def find_max_font_size(text, max_width, max_height, font_name):

    low = 1
    high = 200
    best = 1

    while low <= high:

        mid = (low + high) // 2

        lines = wrap_text(text, font_name, mid, max_width - LABEL_MARGIN)

        if not lines:
            return 1

        max_line_width = max(
            pdfmetrics.stringWidth(line, font_name, mid) for line in lines
        )

        total_height = len(lines) * mid + (len(lines) - 1) * 2

        if max_line_width <= (max_width - LABEL_MARGIN) and total_height <= (
            max_height - LABEL_MARGIN
        ):
            best = mid
            low = mid + 1
        else:
            high = mid - 1

    return best


# === DRAW LABEL ===
def draw_label_pdf(c, text, font_name, width, height, font_override=0):

    raw_font_size = find_max_font_size(text, width, height, font_name)

    font_size = max(raw_font_size - FONT_ADJUSTMENT + font_override, 1)

    lines = wrap_text(text, font_name, font_size, width - LABEL_MARGIN)

    c.setFont(font_name, font_size)

    total_height = len(lines) * font_size + (len(lines) - 1) * 2

    start_y = (height - total_height) / 2

    for i, line in enumerate(lines):

        line_width = pdfmetrics.stringWidth(line, font_name, font_size)

        x = (width - line_width) / 2

        y = start_y + (len(lines) - i - 1) * (font_size + 2)

        c.drawString(x, y, line)


# === CREATE MULTI PAGE PDF ===
def create_pdf(data_list, font_name, width, height, font_override=0):

    buffer = BytesIO()

    c = canvas.Canvas(buffer, pagesize=(width, height))

    for value in data_list:

        text = str(value).strip()

        if not text or text.lower() == "nan":
            continue

        draw_label_pdf(c, text, font_name, width, height, font_override)

        c.showPage()

    c.save()

    buffer.seek(0)

    return buffer


# === PREVIEW GENERATOR ===
def generate_preview(text, font_name, width, height, font_override=0):

    buffer = BytesIO()

    c = canvas.Canvas(buffer, pagesize=(width, height))

    draw_label_pdf(c, text, font_name, width, height, font_override)

    c.save()

    buffer.seek(0)

    return buffer


# === STREAMLIT UI ===

st.title("Excel/CSV to Label PDF Generator")

st.write("Generate multi-page PDF labels with custom settings.")

# --- USER SETTINGS ---

selected_font = st.selectbox("Select font", AVAILABLE_FONTS, index=1)

font_override = st.slider(
    "Font size override (+/- points)", min_value=-5, max_value=5, value=0
)

width_mm = st.number_input(
    "Label width (mm)", min_value=10, max_value=500, value=DEFAULT_WIDTH_MM
)

height_mm = st.number_input(
    "Label height (mm)", min_value=10, max_value=500, value=DEFAULT_HEIGHT_MM
)

remove_duplicates = st.checkbox("Remove duplicate values", value=True)

# === LABEL PREVIEW ===

st.subheader("Label Preview")

preview_text = st.text_input("Preview text", "Sample Label")

if st.button("Generate Preview"):

    preview_pdf = generate_preview(
        preview_text,
        selected_font,
        width_mm * mm,
        height_mm * mm,
        font_override,
    )

    st.download_button(
        label="Download Preview PDF",
        data=preview_pdf,
        file_name="preview_label.pdf",
        mime="application/pdf",
    )

# --- FILE UPLOAD ---

uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

df = None

if uploaded_file:

    try:

        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)

        else:
            df = pd.read_excel(uploaded_file, engine="openpyxl")

        st.success("File loaded successfully!")

    except Exception as e:

        st.error(f"Error reading file: {e}")

# --- COLUMN SELECTION ---

if df is not None:

    st.write("Preview of data:")

    st.dataframe(df)

    selected_columns = st.multiselect(
        "Select columns to generate labels",
        options=df.columns.tolist(),
        default=df.columns.tolist(),
    )

    if selected_columns:

        cell_values = df[selected_columns].values.flatten()

        cell_values = [
            str(val).strip()
            for val in cell_values
            if pd.notnull(val) and str(val).strip() != ""
        ]

        if remove_duplicates:
            cell_values = list(dict.fromkeys(cell_values))

        st.info(f"Labels to generate: {len(cell_values)}")

        if st.button("Generate PDF"):

            if not cell_values:

                st.warning("No valid data found!")

            else:

                pdf_buffer = create_pdf(
                    cell_values,
                    selected_font,
                    width_mm * mm,
                    height_mm * mm,
                    font_override,
                )

                st.download_button(
                    label="Download PDF",
                    data=pdf_buffer,
                    file_name="labels.pdf",
                    mime="application/pdf",
                )
