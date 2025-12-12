import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

# ==========================================
# 1. CONFIGURATION & STYLES
# ==========================================
# Fixed logo path (As per your request to fix it in code)
FIXED_LOGO_PATH = "Image.png"

# Color constants based on your images
COLOR_HEADER_BLUE = colors.HexColor("#8ea9db")
COLOR_TABLE_ORANGE = colors.HexColor("#f4b084")

# ReportLab Styles
styles = getSampleStyleSheet()
style_normal = styles["Normal"]
style_center = ParagraphStyle(name='Center', parent=styles['Normal'], alignment=TA_CENTER, fontSize=9)
style_left = ParagraphStyle(name='Left', parent=styles['Normal'], alignment=TA_LEFT, fontSize=9)
style_bold_center = ParagraphStyle(name='BoldCenter', parent=styles['Normal'], fontName='Helvetica-Bold', alignment=TA_CENTER, fontSize=12)
style_bold_left = ParagraphStyle(name='BoldLeft', parent=styles['Normal'], fontName='Helvetica-Bold', alignment=TA_LEFT, fontSize=12)

# ==========================================
# 2. PDF GENERATION LOGIC
# ==========================================
def generate_trolley_pdf(df, top_logo_stream):
    buffer = io.BytesIO()
    
    # Page Setup: A4 Landscape
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=landscape(A4),
        leftMargin=1*cm, 
        rightMargin=1*cm, 
        topMargin=1*cm, 
        bottomMargin=1*cm
    )
    
    elements = []
    
    # --- PRE-PROCESSING DATA ---
    # 1. Create Trolley No Key: Combine RACK + RACK NO (1st) + RACK NO (2nd) -> e.g., TL-0-1
    # We clean the data first to ensure no float/nan issues
    df = df.fillna("")
    
    # Helper to clean strings
    def clean_str(val):
        return str(val).replace('.0', '') if val != "" else ""

    # Check if necessary columns exist for Trolley construction, else fallback
    if all(col in df.columns for col in ['RACK', 'RACK NO (1st digit)', 'RACK NO (2nd digit)']):
        df['Calculated_Trolley'] = df.apply(
            lambda x: f"{clean_str(x['RACK'])}-{clean_str(x['RACK NO (1st digit)'])}{clean_str(x['RACK NO (2nd digit)'])}", 
            axis=1
        )
    elif 'TROLLEY NO' in df.columns:
         df['Calculated_Trolley'] = df['TROLLEY NO']
    else:
         df['Calculated_Trolley'] = "UNKNOWN"

    # Ensure Station Name exists
    if 'STATION NAME' not in df.columns:
        df['STATION NAME'] = ""

    # --- GROUPING LOGIC ---
    # Group by: Station No, Trolley No, Bus Model
    # (Optional: include Station Name in grouping if it varies)
    group_cols = ['STATION NO', 'Calculated_Trolley', 'BUS MODEL', 'STATION NAME']
    # Filter out columns that might not exist to avoid errors
    actual_group_cols = [c for c in group_cols if c in df.columns]
    
    # Sort to keep the PDF orderly
    df.sort_values(by=actual_group_cols, inplace=True)
    
    grouped = df.groupby(actual_group_cols)
    
    # Iterate through each group to create a page/section
    for name, group in grouped:
        # Unpack grouping keys. Note: 'name' is a tuple if multiple columns
        # Map values based on the order of actual_group_cols
        station_no = str(name[0]) if len(actual_group_cols) > 0 else ""
        trolley_no = str(name[1]) if len(actual_group_cols) > 1 else ""
        model = str(name[2]) if len(actual_group_cols) > 2 else ""
        station_name = str(name[3]) if len(actual_group_cols) > 3 else ""
        
        # --- HEADER SECTION ---
        
        # Top Row: "Document Ref No" (Left) and Top Logo (Right)
        top_logo_img = ""
        if top_logo_stream:
            # Create a fresh stream pointer for the image for every page reuse
            img_io = io.BytesIO(top_logo_stream.getvalue())
            top_logo_img = RLImage(img_io, width=4*cm, height=1.2*cm)
            top_logo_img.hAlign = 'RIGHT'
        
        header_meta_data = [
            [Paragraph("<b>Document Ref No.:</b>", style_left), "", top_logo_img]
        ]
        t_meta = Table(header_meta_data, colWidths=[6*cm, 16*cm, 5*cm])
        t_meta.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('ALIGN', (-1,0), (-1,0), 'RIGHT'),
        ]))
        elements.append(t_meta)
        elements.append(Spacer(1, 0.2*cm))
        
        # Main Blue Header Table
        # Layout:
        # Station Name | Station No
        # Model        | Trolley No
        
        blue_header_data = [
            [Paragraph(f"STATION NAME: {station_name}", style_bold_left), "", 
             Paragraph(f"STATION NO: {station_no}", style_bold_left), ""],
            [Paragraph(f"MODEL: {model}", style_bold_left), "", 
             Paragraph(f"TROLLEY NO: {trolley_no}", style_bold_left), ""]
        ]
        
        # Available width is approx 27.7cm (A4 Land - margins)
        # We merge cells to get the look: Col 1 spans, Col 3 spans
        t_header = Table(blue_header_data, colWidths=[8*cm, 4*cm, 8*cm, 7.7*cm], rowHeights=[1*cm, 1*cm])
        t_header.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('BACKGROUND', (0,0), (-1,-1), COLOR_HEADER_BLUE),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('SPAN', (0,0), (1,0)), # Span Station Name
            ('SPAN', (2,0), (3,0)), # Span Station No
            ('SPAN', (0,1), (1,1)), # Span Model
            ('SPAN', (2,1), (3,1)), # Span Trolley No
            ('LEFTPADDING', (0,0), (-1,-1), 6),
            ('TEXTCOLOR', (0,0), (-1,-1), colors.black),
            ('UPPERCASE', (0,0), (-1,-1), True),
        ]))
        elements.append(t_header)
        
        # --- PART LIST TABLE ---
        
        # Columns based on your Image 1: 
        # S.No | PART NO | DESCRIPTION | Qty/Veh | MAX SIZE | QTY / TROLLEY | LOCATION
        
        headers = ["S. No", "PART NO", "DESCRIPTION", "Qty/ Veh", "MAX SIZE", "QTY / TROLLEY", "LOCATION"]
        
        data_rows = [headers]
        
        # Iterate rows in this group
        for idx, row in enumerate(group.to_dict('records')):
            data_rows.append([
                str(idx + 1),
                str(row.get('PARTNO', '')),
                Paragraph(str(row.get('PART DESCRIPTION', '')), style_left),
                clean_str(row.get('Qty / Veh', '')),
                clean_str(row.get('Max Size', '')),
                clean_str(row.get('Qty /Trolley', '')), # Check Excel col name carefully
                str(row.get('LOCATION', ''))
            ])
            
        # Column widths
        cw = [1.5*cm, 4*cm, 9*cm, 1.9*cm, 2.5*cm, 3.3*cm, 5.5*cm]
        
        t_data = Table(data_rows, colWidths=cw, repeatRows=1)
        
        # Styling
        tbl_style = [
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('BACKGROUND', (0,0), (-1,0), COLOR_TABLE_ORANGE), # Orange Header
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTSIZE', (0,0), (-1,-1), 12),
            ('BOTTOMPADDING', (0,0), (-1,0), 6),
            ('TOPPADDING', (0,0), (-1,0), 6),
        ]
        t_data.setStyle(TableStyle(tbl_style))
        elements.append(t_data)
        
        # Push Footer to bottom (Simple approach: Add a large spacer or just append at end of flow)
        # Note: In SimpleDocTemplate, elements flow. To guarantee bottom position requires a Frame or PageTemplate.
        # For simplicity in this script, we append it after the table.
        elements.append(Spacer(1, 1*cm))
        
        # --- FOOTER SECTION ---
        
        # Left side: Date, Verified By
        # Right side: Designed By Logo
        
        creation_date = datetime.now().strftime("%d-%b-%Y")
        
        footer_left_content = [
            [Paragraph(f"<i>Creation Date: {creation_date}</i>", style_left)],
            [Spacer(1, 0.5*cm)],
            [Paragraph("<b>Verified By:</b>", style_left)],
            [Paragraph("Name: ____________________", style_left)],
            [Paragraph("Signature: _________________", style_left)]
        ]
        
        # Load Fixed Logo (Agilomatrix)
        fixed_logo_img = Paragraph("<b>[Agilomatrix Logo Missing]</b>", style_left)
        if os.path.exists(FIXED_LOGO_PATH):
            try:
                # Requirement: width=4.3*cm, height=1.5*cm
                fixed_logo_img = RLImage(FIXED_LOGO_PATH, width=4.3*cm, height=1.5*cm)
            except:
                pass

        footer_right_content = [
            [Paragraph("Designed By:", ParagraphStyle(name='RightAlign', parent=styles['Normal'], alignment=TA_RIGHT))],
            [fixed_logo_img]
        ]
        
        # Embed these in a table to align Left vs Right
        t_footer_left = Table(footer_left_content, colWidths=[10*cm])
        t_footer_left.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'LEFT')]))
        
        t_footer_right = Table(footer_right_content, colWidths=[17*cm])
        t_footer_right.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
            ('VALIGN', (0,0), (-1,-1), 'BOTTOM')
        ]))
        
        t_footer_main = Table([[t_footer_left, t_footer_right]], colWidths=[10*cm, 17.7*cm])
        t_footer_main.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'BOTTOM')]))
        
        elements.append(t_footer_main)
        
        # Page Break after every Trolley List
        elements.append(PageBreak())

    # Build PDF
    doc.build(elements)
    buffer.seek(0)
    return buffer

# ==========================================
# 3. STREAMLIT UI
# ==========================================
st.set_page_config(page_title="Trolley List Generator", layout="wide")

st.title("🏭 Trolley Part List Generator")
st.markdown("""
This tool extracts data from the production Excel sheet and creates formatted **Trolley Part Lists**.
Data is grouped by **Station No**, **Trolley No**, and **Model**.
""")

col1, col2 = st.columns([1, 1])

with col1:
    uploaded_file = st.file_uploader("📂 Upload Excel Data", type=["xlsx", "xls"])

with col2:
    top_logo_file = st.file_uploader("🖼️ Upload Client Logo (Top Right)", type=["png", "jpg", "jpeg"])

# Warning/Info about the Fixed Logo
if not os.path.exists(FIXED_LOGO_PATH):
    st.warning(f"⚠️ The fixed logo file `{FIXED_LOGO_PATH}` was not found in the directory. A text placeholder will be used.")
else:
    st.success(f"✅ Fixed logo `{FIXED_LOGO_PATH}` found.")

if uploaded_file is not None:
    try:
        # Read Excel
        df = pd.read_excel(uploaded_file)
        
        st.subheader("Data Preview")
        st.dataframe(df.head())
        
        # Required Columns Check based on your images
        req_cols = ['STATION NO', 'BUS MODEL', 'PARTNO', 'PART DESCRIPTION', 'LOCATION']
        missing_cols = [c for c in req_cols if c not in df.columns]
        
        if missing_cols:
            st.error(f"❌ The uploaded Excel is missing required columns: {', '.join(missing_cols)}")
        else:
            if st.button("Generate Trolley List PDF"):
                with st.spinner("Processing..."):
                    pdf_data = generate_trolley_pdf(df, top_logo_file)
                    
                    st.success("PDF Generated Successfully!")
                    st.download_button(
                        label="⬇️ Download Trolley Part List.pdf",
                        data=pdf_data,
                        file_name=f"Trolley_List_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                        mime="application/pdf"
                    )

    except Exception as e:
        st.error(f"An error occurred during processing: {e}")
