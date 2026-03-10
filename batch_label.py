import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from io import BytesIO

# ==============================
# SETTINGS
# ==============================

DEFAULT_WIDTH_MM = 50
DEFAULT_HEIGHT_MM = 30

FONT_SIZE = 16

AVAILABLE_FONTS = [
    "Helvetica",
    "Helvetica-Bold",
    "Times-Roman",
    "Times-Bold",
    "Courier",
    "Courier-Bold"
]

# ==============================
# DRAW LABEL
# ==============================

def draw_label_pdf(c, label_text, font_name, width, height):

    lines = label_text.split("\n")

    c.setFont(font_name, FONT_SIZE)

    total_height = len(lines) * FONT_SIZE + (len(lines) - 1) * 4

    start_y = (height - total_height) / 2

    for i, line in enumerate(lines):

        line_width = pdfmetrics.stringWidth(line, font_name, FONT_SIZE)

        x = (width - line_width) / 2

        y = start_y + (len(lines) - i - 1) * (FONT_SIZE + 4)

        c.drawString(x, y, line)


# ==============================
# CREATE PDF
# ==============================

def create_pdf(labels, font_name, width, height):

    buffer = BytesIO()

    c = canvas.Canvas(buffer, pagesize=(width, height))

    for text in labels:

        draw_label_pdf(c, text, font_name, width, height)
        c.showPage()

    c.save()

    buffer.seek(0)

    return buffer


# ==============================
# STREAMLIT UI
# ==============================

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


# ==============================
# FILE PROCESSING
# ==============================

if uploaded_file:

    if uploaded_file.name.endswith(".csv"):
        df_raw = pd.read_csv(uploaded_file, header=None)
    else:
        df_raw = pd.read_excel(uploaded_file, header=None)

    st.write("Raw file preview")
    st.dataframe(df_raw)

    df_raw = df_raw.dropna(axis=1, how="all")

    # ==============================
    # FIND BATCH NUMBER
    # ==============================

    batch_number = None

    for i in range(len(df_raw)):
        for j in range(len(df_raw.columns)):

            cell = str(df_raw.iloc[i, j]).strip().lower()

            if cell == "batch number":

                batch_number = str(df_raw.iloc[i, j + 1]).strip()
                break

        if batch_number:
            break

    if not batch_number:

        st.error("Batch number not found.")
        st.stop()

    # ==============================
    # FIND NAME + WEIGHT HEADER
    # ==============================

    header_row = None
    name_col = None
    weight_col = None

    for i in range(len(df_raw)):

        row_values = [str(x).strip().lower() for x in df_raw.iloc[i]]

        if "name" in row_values and "weight" in row_values:

            header_row = i
            name_col = row_values.index("name")
            weight_col = row_values.index("weight")

            break

    if header_row is None:

        st.error("Could not find Name and Weight headers.")
        st.stop()

    # ==============================
    # EXTRACT DATA
    # ==============================

    df = df_raw.iloc[header_row + 1:].reset_index(drop=True)

    labels = []

    for _, row in df.iterrows():

        name = str(row.iloc[name_col]).strip()
        weight = str(row.iloc[weight_col]).strip()

        if name.lower() == "nan" or name == "":
            continue

        if weight.lower() == "nan":
            weight = ""

        label_text = f"BN/{batch_number}\n{name} {weight}".strip()

        labels.append(label_text)

    labels = list(dict.fromkeys(labels))

    st.info(f"Labels to generate: {len(labels)}")

    if st.button("Generate Labels"):

        pdf_buffer = create_pdf(
            labels,
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
