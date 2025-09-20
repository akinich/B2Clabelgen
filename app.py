import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from io import BytesIO
import requests

# === DEFAULT CONSTANTS ===
DEFAULT_WIDTH_MM = 50
DEFAULT_HEIGHT_MM = 30
FONT_ADJUSTMENT = 2  # for printer safety

# Built-in fonts
AVAILABLE_FONTS = [
    "Helvetica",
    "Helvetica-Bold",
    "Times-Roman",
    "Times-Bold",
    "Courier",
    "Courier-Bold"
]

# === HELPER FUNCTIONS ===
def find_max_font_size_for_multiline(lines, max_width, max_height, font_name):
    font_size = 1
    while True:
        max_line_width = max(pdfmetrics.stringWidth(line, font_name, font_size) for line in lines)
        total_height = len(lines) * font_size + (len(lines) - 1) * 2
        if max_line_width > (max_width - 4) or total_height > (max_height - 4):
            return max(font_size - 1, 1)
        font_size += 1

def draw_label(c, text, font_name, width, height, font_override=0):
    lines = text.split()
    raw_font_size = find_max_font_size_for_multiline(lines, width, height, font_name)
    font_size = max(raw_font_size - FONT_ADJUSTMENT + font_override, 1)
    c.setFont(font_name, font_size)

    total_height = len(lines) * font_size + (len(lines) - 1) * 2
    start_y = (height - total_height) / 2

    for i, line in enumerate(lines):
        line_width = pdfmetrics.stringWidth(line, font_name, font_size)
        x = (width - line_width) / 2
        y = start_y + (len(lines) - i - 1) * (font_size + 2)
        c.drawString(x, y, line)

def create_pdf(data_list, font_name, width, height, font_override=0):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=(width, height))
    for value in data_list:
        text = str(value).strip()
        if not text or text.lower() == "nan":
            continue
        draw_label(c, text, font_name, width, height, font_override)
        c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

# === STREAMLIT UI ===
st.title("Excel/CSV/Google Sheet to Label PDF Generator")
st.write("""
Generate multi-page PDF labels with custom settings.
""")

# --- User Inputs ---
selected_font = st.selectbox("Select font", AVAILABLE_FONTS, index=1)
font_override = st.slider("Font size override (+/- points)", min_value=-5, max_value=5, value=0)

width_mm = st.number_input("Label width (mm)", min_value=10, max_value=500, value=DEFAULT_WIDTH_MM)
height_mm = st.number_input("Label height (mm)", min_value=10, max_value=500, value=DEFAULT_HEIGHT_MM)
page_width = width_mm * mm
page_height = height_mm * mm

remove_duplicates = st.checkbox("Remove duplicate values", value=True)

google_sheet_url = st.text_input("Or paste Google Sheet CSV URL (optional)")

uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

# --- Load Data ---
df = None
if google_sheet_url:
    try:
        response = requests.get(google_sheet_url)
        response.raise_for_status()
        df = pd.read_csv(BytesIO(response.content))
        st.success("Google Sheet loaded successfully!")
    except Exception as e:
        st.warning(f"Could not load Google Sheet: {e}")

elif uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
        st.success("File loaded successfully!")
    except Exception as e:
        st.error(f"Error reading file: {e}")

# --- Column Selection ---
if df is not None:
    st.write("Preview of data:")
    st.dataframe(df)
    
    selected_columns = st.multiselect(
        "Select columns to generate labels",
        options=df.columns.tolist(),
        default=df.columns.tolist()
    )
    
    if not selected_columns:
        st.warning("Please select at least one column.")
    else:
        # Flatten selected columns
        cell_values = df[selected_columns].values.flatten()
        # Remove empty / NaN
        cell_values = [str(val).strip() for val in cell_values if pd.notnull(val) and str(val).strip() != ""]
        if remove_duplicates:
            cell_values = list(dict.fromkeys(cell_values))  # keep order, remove duplicates
        
        # Preview first label text
        if cell_values:
            st.subheader("Preview of first label (text only)")
            st.markdown(f"**{cell_values[0]}** in {selected_font}")

        # Generate PDF
        if st.button("Generate PDF"):
            if not cell_values:
                st.warning("No valid data found!")
            else:
                pdf_buffer = create_pdf(cell_values, selected_font, page_width, page_height, font_override)
                st.download_button(
                    label="Download PDF",
                    data=pdf_buffer,
                    file_name="labels.pdf",
                    mime="application/pdf"
                )
