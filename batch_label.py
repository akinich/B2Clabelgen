import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from io import BytesIO

# === DEFAULT CONSTANTS ===
DEFAULT_WIDTH_MM = 50
DEFAULT_HEIGHT_MM = 30
LABEL_MARGIN = 4

AVAILABLE_FONTS = [
    "Helvetica",
    "Helvetica-Bold",
    "Times-Roman",
    "Times-Bold",
    "Courier",
    "Courier-Bold"
]


# === TEXT WRAP ===
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


# === FAST FONT SIZE SEARCH ===
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

        if max_line_width <= (max_width - LABEL_MARGIN) and total_height <= (max_height):
            best = mid
            low = mid + 1
        else:
            high = mid - 1

    return best


# === DRAW STANDARDIZED LABEL ===
def draw_label_pdf(c, batch_number, product_text, font_name, width, height):

    batch_text = f"BN/{batch_number}"

    # --- Batch Font (smaller) ---
    batch_font_size = int(height * 0.18)

    c.setFont(font_name, batch_font_size)

    batch_width = pdfmetrics.stringWidth(batch_text, font_name, batch_font_size)

    batch_x = (width - batch_width) / 2
    batch_y = height - batch_font_size - 4

    c.drawString(batch_x, batch_y, batch_text)

    # --- Product Text Area ---
    product_area_height = height - batch_font_size - 10

    font_size = find_max_font_size(
        product_text,
        width,
        product_area_height,
        font_name
    )

    lines = wrap_text(product_text, font_name, font_size, width - LABEL_MARGIN)

    c.setFont(font_name, font_size)

    total_height = len(lines) * font_size + (len(lines) - 1) * 2

    start_y = (product_area_height - total_height) / 2

    for i, line in enumerate(lines):

        line_width = pdfmetrics.stringWidth(line, font_name, font_size)

        x = (width - line_width) / 2

        y = start_y + (len(lines) - i - 1) * (font_size + 2)

        c.drawString(x, y, line)


# === CREATE PDF ===
def create_pdf(labels, batch_number, font_name, width, height):

    buffer = BytesIO()

    c = canvas.Canvas(buffer, pagesize=(width, height))

    for text in labels:

        draw_label_pdf(c, batch_number, text, font_name, width, height)

        c.showPage()

    c.save()

    buffer.seek(0)

    return buffer


# === STREAMLIT UI ===

st.title("Batch Label Generator")

selected_font = st.selectbox("Font", AVAILABLE_FONTS, index=1)

width_mm = st.number_input("Label width (mm)", min_value=10, max_value=200, value=50)
height_mm = st.number_input("Label height (mm)", min_value=10, max_value=200, value=30)

uploaded_file = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])

if uploaded_file:

    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.dataframe(df)

    # === FIND BATCH NUMBER ===

    batch_number = None

    for col in df.columns:

        if "batch" in col.lower():

            batch_number = str(df[col].dropna().iloc[0])

            break

    if not batch_number:
        st.error("Batch number not found")
        st.stop()

    # === VALIDATE COLUMNS ===

    if "Name" not in df.columns or "Weight" not in df.columns:
        st.error("Columns 'Name' and 'Weight' are required.")
        st.stop()

    labels = []

    for _, row in df.iterrows():

        name = str(row["Name"]).strip()
        weight = str(row["Weight"]).strip()

        if not name or name.lower() == "nan":
            continue

        if weight.lower() == "nan":
            weight = ""

        product_text = f"{name} {weight}".strip()

        labels.append(product_text)

    # remove duplicates
    labels = list(dict.fromkeys(labels))

    st.info(f"Labels to generate: {len(labels)}")

    if st.button("Generate Labels"):

        pdf_buffer = create_pdf(
            labels,
            batch_number,
            selected_font,
            width_mm * mm,
            height_mm * mm
        )

        st.download_button(
            label="Download PDF",
            data=pdf_buffer,
            file_name="labels.pdf",
            mime="application/pdf"
        )
