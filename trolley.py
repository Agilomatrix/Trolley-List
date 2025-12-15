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
# Fixed logo path
FIXED_LOGO_PATH = "Image.png"

# Color constants
COLOR_HEADER_BLUE = colors.HexColor("#8ea9db")
COLOR_TABLE_ORANGE = colors.HexColor("#f4b084")

# ReportLab Styles
styles = getSampleStyleSheet()
style_normal = styles["Normal"]
style_center = ParagraphStyle(name='Center', parent=styles['Normal'], alignment=TA_CENTER, fontSize=9)
style_left = ParagraphStyle(name='Left', parent=styles['Normal'], alignment=TA_LEFT, fontSize=9)
style_bold_center = ParagraphStyle(name='BoldCenter', parent=styles['Normal'], fontName='Helvetica-Bold', alignment=TA_CENTER, fontSize=14)
style_bold_left = ParagraphStyle(name='BoldLeft', parent=styles['Normal'], fontName='Helvetica-Bold', alignment=TA_LEFT, fontSize=16)

# ==========================================
# 2. PDF GENERATION LOGIC
# ==========================================
def generate_trolley_pdf(df, top_logo_stream, logo_w, logo_h):
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
    df = df.fillna("")
    
    # Helper to clean strings
    def clean_str(val):
        return str(val).replace('.0', '') if val != "" else ""

    # Calculate Trolley ID
    if all(col in df.columns for col in ['RACK', 'RACK NO (1st digit)', 'RACK NO (2nd digit)']):
        df['Calculated_Trolley'] = df.apply(
            lambda x: f"{clean_str(x['RACK'])} - {clean_str(x['RACK NO (1st digit)'])}{clean_str(x['RACK NO (2nd digit)'])}", 
            axis=1
        )
    elif 'TROLLEY NO' in df.columns:
         df['Calculated_Trolley'] = df['TROLLEY NO']
    else:
         df['Calculated_Trolley'] = "UNKNOWN"

    if 'STATION NAME' not in df.columns:
        df['STATION NAME'] = ""

    # Grouping
    group_cols = ['STATION NO', 'Calculated_Trolley', 'BUS MODEL', 'STATION NAME']
    actual_group_cols = [c for c in group_cols if c in df.columns]
    
    df.sort_values(by=actual_group_cols, inplace=True)
    grouped = df.groupby(actual_group_cols)
    
    # Iterate groups
    for name, group in grouped:
        station_no = str(name[0]) if len(actual_group_cols) > 0 else ""
        trolley_no = str(name[1]) if len(actual_group_cols) > 1 else ""
        model = str(name[2]) if len(actual_group_cols) > 2 else ""
        station_name = str(name[3]) if len(actual_group_cols) > 3 else ""
        
        # --- TOP HEADER (Ref No & Logo) ---
        top_logo_img = ""
        if top_logo_stream:
            try:
                # Create fresh stream for every loop to avoid closed file errors
                img_io = io.BytesIO(top_logo_stream.getvalue())
                # Use User Defined Dimensions (logo_w, logo_h)
                top_logo_img = RLImage(img_io, width=logo_w*cm, height=logo_h*cm)
                top_logo_img.hAlign = 'RIGHT'
            except Exception:
                pass
        
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
        
        # --- BLUE HEADER TABLE ---
        blue_header_data = [
            [
                Paragraph("<b>STATION NAME</b>", style_bold_left), 
                Paragraph(station_name, style_bold_left),
                Paragraph("<b>STATION NO</b>", style_bold_left), 
                Paragraph(station_no, style_bold_left)
            ],
            [
                Paragraph("<b>MODEL</b>", style_bold_left), 
                Paragraph(model, style_bold_left),
                Paragraph("<b>TROLLEY NO</b>", style_bold_left), 
                Paragraph(trolley_no, style_bold_left)
            ]
        ]
        
        header_col_widths = [5*cm, 9.9*cm, 4.8*cm, 8*cm]
        t_header = Table(blue_header_data, colWidths=header_col_widths, rowHeights=[1*cm, 1*cm])
        
        t_header.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('BACKGROUND', (0,0), (-1,-1), COLOR_HEADER_BLUE),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('LEFTPADDING', (0,0), (-1,-1), 6),
            ('TEXTCOLOR', (0,0), (-1,-1), colors.black),
            ('UPPERCASE', (0,0), (-1,-1), True),
        ]))
        elements.append(t_header)
        
        # --- PART LIST TABLE ---
        headers = ["S. No", "PART NO", "DESCRIPTION", "Qty/ Veh", "MAX SIZE", "QTY / TROLLEY", "LOCATION"]
        data_rows = [headers]
        
        for idx, row in enumerate(group.to_dict('records')):
            data_rows.append([
                str(idx + 1),
                str(row.get('PARTNO', '')),
                Paragraph(str(row.get('PART DESCRIPTION', '')), style_left),
                clean_str(row.get('Qty / Veh', '')),
                clean_str(row.get('Max Size', '')),
                clean_str(row.get('Qty /Trolley', '')),
                str(row.get('LOCATION', ''))
            ])
            
        cw = [1.5*cm, 4*cm, 9*cm, 1.9*cm, 2.5*cm, 3.3*cm, 5.5*cm]
        t_data = Table(data_rows, colWidths=cw, repeatRows=1)
        
        tbl_style = [
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('BACKGROUND', (0,0), (-1,0), COLOR_TABLE_ORANGE),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTSIZE', (0,0), (-1,-1), 12),
            ('BOTTOMPADDING', (0,0), (-1,0), 6),
            ('TOPPADDING', (0,0), (-1,0), 6),
        ]
        t_data.setStyle(TableStyle(tbl_style))
        elements.append(t_data)
        elements.append(Spacer(1, 0.5*cm))
        
        # --- FOOTER SECTION ---
        creation_date = datetime.now().strftime("%d-%b-%Y")
        
        footer_left_content = [
            [Paragraph(f"<i>Creation Date: {creation_date}</i>", style_left)],
            [Spacer(1, 0.1*cm)],
            [Paragraph("<b>Verified By:</b>", style_left)],
            [Paragraph("Name: ____________________", style_left)],
            [Paragraph("Signature: _________________", style_left)]
        ]
        
        # Load Fixed Logo (Agilomatrix) with check
        fixed_logo_img = Paragraph("<b>[Agilomatrix Logo Missing]</b>", rl_cell_left_style)
        if os.path.exists(fixed_logo_path):
            try:
                 fixed_logo_img = RLImage(fixed_logo_path, width=4.3*cm, height=1.5*cm)
            except:
                 pass
        
        left_content = [
            Paragraph(f"<i>Creation Date: {today_date}</i>", rl_cell_left_style),
            Spacer(1, 0.2*cm),
            Paragraph("<b>Verified by:</b>", ParagraphStyle('BoldFooter', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT)),
            Paragraph("Name: ___________________", rl_cell_left_style),
            Paragraph("Signature: _______________", rl_cell_left_style)
        ]

        designed_by_text = Paragraph("Designed by:", ParagraphStyle('DesignedBy', fontName='Helvetica', fontSize=10, alignment=TA_RIGHT))
        right_inner_table = Table([[designed_by_text, fixed_logo_img]], colWidths=[3*cm, 4.5*cm])
        right_inner_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ]))

        footer_table = Table([[left_content, right_inner_table]], colWidths=[20*cm, 7.7*cm])
        footer_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'BOTTOM'),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ]))
        
        elements.append(footer_table)
        elements.append(PageBreak())

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
    
    # --- ADDED: Inputs for Top Logo Dimensions ---
    top_logo_w = 3.0 # Default
    top_logo_h = 2.8 # Default
    if top_logo_file:
        st.caption("Define Logo Dimensions (cm):")
        c_w, c_h = st.columns(2)
        top_logo_w = c_w.number_input("Width (cm)", min_value=0.5, max_value=10.0, value=3.0, step=0.1)
        top_logo_h = c_h.number_input("Height (cm)", min_value=0.5, max_value=5.0, value=2.8, step=0.1)

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
                    # Pass the user defined width and height to the function
                    pdf_data = generate_trolley_pdf(df, top_logo_file, top_logo_w, top_logo_h)
                    
                    st.success("PDF Generated Successfully!")
                    st.download_button(
                        label="⬇️ Download Trolley Part List.pdf",
                        data=pdf_data,
                        file_name=f"Trolley_List_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                        mime="application/pdf"
                    )

    except Exception as e:
        st.error(f"An error occurred during processing: {e}")
