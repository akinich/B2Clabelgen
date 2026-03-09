import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from io import BytesIO

# === CONSTANTS ===

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

        if max_line_width <= (max_width - LABEL_MARGIN) and total_height <= max_height:
            best = mid
            low = mid + 1
        else:
            high = mid - 1

    return best


# === DRAW LABEL ===

def draw_label_pdf(c, batch_number, product_text, font_name, width, height):

    batch_text = f"BN/{batch_number}"

    # Batch line (smaller)
    batch_font_size = int(height * 0.18)

    c.setFont(font_name, batch_font_size)

    batch_width = pdfmetrics.stringWidth(batch_text, font_name, batch_font_size)

    batch_x = (width - batch_width) / 2
    batch_y = height - batch_font_size - 4

    c.drawString(batch_x, batch_y, batch_text)

    # Product text area
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

width_mm = st.number_input(
    "Label width (mm)",
    min_value=10,
    max_value=200,
    value=DEFAULT_WIDTH_MM
)

height_mm = st.number_input(
    "Label height (mm)",
    min_value=10,
    max_value=200,
    value=DEFAULT_HEIGHT_MM
)

uploaded_file = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])


# === FILE PROCESSING ===

if uploaded_file:

    if uploaded_file.name.endswith(".csv"):
        df_raw = pd.read_csv(uploaded_file, header=None)
    else:
        df_raw = pd.read_excel(uploaded_file, header=None)

    st.write("Raw file preview")
    st.dataframe(df_raw)

    # Remove empty columns
    df_raw = df_raw.dropna(axis=1, how="all")

    # === BATCH NUMBER FROM C2 ===
    try:
        batch_number = str(df_raw.iloc[0,2]).strip()
    except:
        st.error("Could not read batch number from cell C2.")
        st.stop()

    if batch_number.lower() == "nan" or batch_number == "":
        st.error("Batch number missing in C2.")
        st.stop()

    # === CREATE DATA TABLE ===
    df = df_raw.iloc[1:].reset_index(drop=True)
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)

    st.write("Parsed data")
    st.dataframe(df)

    if "Name" not in df.columns or "Weight" not in df.columns:
        st.error("Columns 'Name' and 'Weight' not found.")
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

    # Remove duplicates
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
