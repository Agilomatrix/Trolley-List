import streamlit as st
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm, inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import io
import base64
from datetime import datetime

# ==========================================
# 1. FIXED ASSETS (AGILOMATRIX LOGO)
# ==========================================
# Since you want the logo fixed in the code, we use a Base64 string.
# NOTE: Replace this long string with the Base64 of your actual Agilomatrix logo 
# if you want the high-res version embedded without a file. 
# For now, this is a placeholder 1x1 pixel image to prevent errors if no file is present.
# Ideally, place 'agilomatrix_logo.png' in the same folder and the code below will use it.
AGILO_LOGO_PATH = "agilomatrix_logo.png" 

# ==========================================
# 2. PDF GENERATION LOGIC
# ==========================================

class TrolleyPDF:
    def __init__(self, buffer, top_logo_byte_stream):
        self.buffer = buffer
        self.top_logo_stream = top_logo_byte_stream
        self.width, self.height = landscape(A4)
        self.styles = getSampleStyleSheet()
        
        # Current Page Context (updated during loop)
        self.current_station_name = ""
        self.current_model = ""
        self.current_station_no = ""
        self.current_trolley_no = ""

    def header_footer(self, canvas, doc):
        canvas.saveState()
        
        # --- TOP RIGHT LOGO (User Uploaded) ---
        if self.top_logo_stream:
            try:
                # Reset pointer
                self.top_logo_stream.seek(0)
                # Draw Image: x, y, width, height (Approx adjust to top right)
                canvas.drawImage(
                    self.top_logo_stream, 
                    doc.width + doc.leftMargin - 4*cm, 
                    doc.height + doc.topMargin - 1.5*cm, 
                    width=4*cm, height=1.2*cm, 
                    preserveAspectRatio=True, mask='auto'
                )
            except Exception as e:
                print(f"Error loading top logo: {e}")

        # --- DOCUMENT REF NO ---
        canvas.setFont('Helvetica-Bold', 10)
        canvas.drawString(doc.leftMargin, doc.height + doc.topMargin - 0.5*cm, "Document Ref No.:")

        # --- HEADER TABLE (Blue/Orange) ---
        # We draw this manually using a Table because it needs specific positioning
        header_data = [
            [f"STATION NAME: {self.current_station_name}", "", f"STATION NO: {self.current_station_no}", ""],
            [f"MODEL: {self.current_model}", "", f"TROLLEY NO: {self.current_trolley_no}", ""]
        ]
        
        # Widths: The total width should match the page body width
        total_w = doc.width
        col_widths = [total_w * 0.35, total_w * 0.15, total_w * 0.35, total_w * 0.15]
        
        t = Table(header_data, colWidths=col_widths, rowHeights=[0.8*cm, 0.8*cm])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#8ea9db")), # Blueish color
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            # Merge cells if needed based on the image (Station Name spans 2 cols?)
            # Based on image: Station Name (col 1), Empty (col 2), Station No (col 3), Empty (col 4)
            # Adjusting to match image visually:
            ('SPAN', (0,0), (1,0)), # Station Name spans
            ('SPAN', (2,0), (3,0)), # Station No spans
            ('SPAN', (0,1), (1,1)), # Model spans
            ('SPAN', (2,1), (3,1)), # Trolley No spans
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('UPPERCASE', (0,0), (-1,-1), True),
        ]))
        
        w, h = t.wrap(doc.width, doc.topMargin)
        t.drawOn(canvas, doc.leftMargin, doc.height + doc.topMargin - 3.0*cm)

        # --- FOOTER ---
        # Creation Date
        date_str = datetime.now().strftime("%d-%b-%Y")
        canvas.setFont('Helvetica-Oblique', 10)
        canvas.drawString(doc.leftMargin, 2.5*cm, f"Creation Date: {date_str}")

        # Verified By
        canvas.setFont('Helvetica-Bold', 10)
        canvas.drawString(doc.leftMargin, 1.8*cm, "Verified By:")
        canvas.setFont('Helvetica', 10)
        canvas.drawString(doc.leftMargin, 1.3*cm, "Name:")
        canvas.drawString(doc.leftMargin, 0.8*cm, "Signature:")

        # Designed By (Bottom Right)
        canvas.drawRightString(doc.width + doc.leftMargin - 4.5*cm, 1.0*cm, "Designed By:")
        
        # Fixed Agilomatrix Logo
        # Tries to load from file 'agilomatrix_logo.png' if it exists
        # Dimensions requested: w=4.3cm, h=1.5cm
        try:
            logo_x = doc.width + doc.leftMargin - 4.3*cm
            logo_y = 0.5*cm # Bottom margin buffer
            canvas.drawImage(AGILO_LOGO_PATH, logo_x, logo_y, width=4.3*cm, height=1.5*cm, mask='auto', preserveAspectRatio=True)
        except:
            # Fallback if image not found on server
            canvas.setFont('Helvetica-Oblique', 8)
            canvas.drawString(logo_x, logo_y + 0.5*cm, "[Logo Placeholder]")

        canvas.restoreState()

    def create_pdf(self, grouped_data):
        doc = SimpleDocTemplate(
            self.buffer,
            pagesize=landscape(A4),
            rightMargin=0.5*inch, leftMargin=0.5*inch,
            topMargin=1.8*inch, bottomMargin=1.5*inch # Margins accommodate header/footer
        )

        elements = []
        
        # Loop through each group (Station/Trolley combo)
        for group_key, df_group in grouped_data.items():
            
            # Update Context for Header
            # group_key = (StationNo, TrolleyNo, StationName, Model)
            self.current_station_no = group_key[0]
            self.current_trolley_no = group_key[1]
            self.current_station_name = group_key[2]
            self.current_model = group_key[3]

            # --- MAIN DATA TABLE ---
            # Columns based on image: S.No, Part No, Description, Qty/Veh, Max Size, Qty/Trolley, Location
            
            table_data = []
            # Table Header Row (Orange)
            headers = ['S. No', 'PART NO', 'DESCRIPTION', 'Qty/ Veh', 'MAX SIZE', 'QTY / TROLLEY', 'LOCATION']
            table_data.append(headers)

            # Table Rows
            row_idx = 1
            for index, row in df_group.iterrows():
                table_data.append([
                    str(row_idx),
                    str(row.get('PARTNO', '')),
                    str(row.get('PART DESCRIPTION', '')),
                    str(row.get('Qty / Veh', '')),
                    str(row.get('Max Size', '')), # Assuming column exists or empty
                    str(row.get('Qty /Trolley', '')), # Assuming column exists or empty
                    str(row.get('LOCATION', ''))
                ])
                row_idx += 1

            # Column Widths
            avail_width = doc.width
            # Adjust these ratios based on content length
            cw = [avail_width*0.05, avail_width*0.15, avail_width*0.35, avail_width*0.08, avail_width*0.08, avail_width*0.10, avail_width*0.19]

            t = Table(table_data, colWidths=cw, repeatRows=1)
            
            # Style the table
            style = [
                # Header Row Style (Orange)
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#f4b084")),
                ('ALIGN', (0,0), (-1,0), 'CENTER'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,0), 10),
                ('BOTTOMPADDING', (0,0), (-1,0), 6),
                
                # Body Style
                ('GRID', (0,0), (-1,-1), 1, colors.black),
                ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
                ('FONTSIZE', (0,1), (-1,-1), 9),
                ('ALIGN', (0,1), (0,-1), 'CENTER'), # S.No Center
                ('ALIGN', (3,1), (5,-1), 'CENTER'), # Qtys Center
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('LEFTPADDING', (0,0), (-1,-1), 3),
                ('RIGHTPADDING', (0,0), (-1,-1), 3),
            ]
            t.setStyle(TableStyle(style))
            
            elements.append(t)
            
            # Page Break after every group
            elements.append(PageBreak())

        # Build PDF
        doc.build(elements, onFirstPage=self.header_footer, onLaterPages=self.header_footer)

# ==========================================
# 3. STREAMLIT UI
# ==========================================

st.set_page_config(page_title="Trolley List Generator", layout="wide")

st.title("🏭 Trolley Part List Generator")
st.markdown("""
Upload the production Excel sheet. The tool will group parts by **Station**, **Trolley**, and **Model** 
and generate a formatted PDF.
""")

col1, col2 = st.columns(2)

with col1:
    uploaded_file = st.file_uploader("Upload Excel Data", type=["xlsx", "xls"])

with col2:
    top_logo_file = st.file_uploader("Upload Company Logo (Top Right)", type=["png", "jpg", "jpeg"])

# Information about Fixed Logo
st.info(f"ℹ️ The bottom-right 'Designed By' logo expects a file named `{AGILO_LOGO_PATH}` in the root folder. If running locally, please ensure this file exists.")

if uploaded_file is not None and top_logo_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        
        # --- DATA PREPROCESSING ---
        # 1. Ensure required columns exist (Mapping Excel header to Code logic)
        # Based on snippet provided: 
        # Needs: STATION NO, BUS MODEL, RACK, RACK NO (1st digit), RACK NO (2nd digit), PARTNO, PART DESCRIPTION, Qty / Veh, LOCATION
        
        # Check specific column names from your image
        required_cols = ['STATION NO', 'BUS MODEL', 'PARTNO', 'PART DESCRIPTION', 'Qty / Veh', 'LOCATION']
        missing = [c for c in required_cols if c not in df.columns]
        
        if missing:
            st.error(f"Missing columns in Excel: {', '.join(missing)}")
        else:
            # 2. Construct 'TROLLEY NO'
            # Logic: Combine 'RACK' + 'RACK NO (1st)' + 'RACK NO (2nd)' -> e.g., TL-01
            # Assuming columns exist based on your screenshot
            try:
                # Helper to format numbers with leading zero if needed
                df['RACK'] = df['RACK'].astype(str)
                df['R1'] = df['RACK NO (1st digit)'].astype(str)
                df['R2'] = df['RACK NO (2nd digit)'].astype(str)
                df['TROLLEY_ID'] = df['RACK'] + "-" + df['R1'] + df['R2']
            except KeyError:
                 # Fallback if Rack columns missing, maybe user has a 'Trolley No' column
                 if 'Trolley No' in df.columns:
                     df['TROLLEY_ID'] = df['Trolley No']
                 else:
                     df['TROLLEY_ID'] = "Unknown"
            
            # 3. Handle 'STATION NAME'
            # The prompt says "Station Name: UnderBody". If this column isn't in Excel, allow manual input or default.
            if 'STATION NAME' not in df.columns:
                # Create default based on Station No or Empty
                df['STATION NAME'] = "" 
            
            # 4. Fill NaNs for display
            df.fillna("", inplace=True)

            # 5. Grouping Logic
            # "Data will insert from Same Station No, Same Trolley No"
            # We group by: Station No, Trolley ID, Station Name, Model
            grouped = {}
            # Sort to keep order tidy
            df.sort_values(by=['STATION NO', 'TROLLEY_ID'], inplace=True)
            
            groups = df.groupby(['STATION NO', 'TROLLEY_ID', 'STATION NAME', 'BUS MODEL'])
            
            for name, group in groups:
                # name is a tuple: (STATION NO, TROLLEY_ID, STATION NAME, BUS MODEL)
                grouped[name] = group

            if st.button("Generate PDF"):
                # Generate PDF
                buffer = io.BytesIO()
                
                # Convert UploadedFile to BytesIO for ReportLab
                logo_stream = io.BytesIO(top_logo_file.getvalue())
                
                generator = TrolleyPDF(buffer, logo_stream)
                generator.create_pdf(grouped)
                
                buffer.seek(0)
                
                st.success("PDF Generated Successfully!")
                st.download_button(
                    label="⬇️ Download Trolley Part List PDF",
                    data=buffer,
                    file_name="Trolley_Part_List.pdf",
                    mime="application/pdf"
                )

    except Exception as e:
        st.error(f"An error occurred: {e}")
