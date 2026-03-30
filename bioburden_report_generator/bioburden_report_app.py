# bioburden_report_final.py
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from datetime import datetime
import streamlit as st
from io import BytesIO
import os

class BioburdenReportGenerator:
    def __init__(self, logo_path="SASOmed_Logo.png"):
        self.wb = openpyxl.Workbook()
        
        # Create two worksheets
        self.ws_cover = self.wb.active
        self.ws_cover.title = "Cover Page"
        self.ws_content = self.wb.create_sheet("Test Report")
        
        self.logo_path = logo_path
        
        # Define styles
        self.styles = self._define_styles()
        
        # Set A4 page setup for both sheets
        self.set_a4_page_setup()
        
    def set_a4_page_setup(self):
        """Configure A4 page setup for both worksheets"""
        for ws in [self.ws_cover, self.ws_content]:
            # Page setup for A4
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.page_setup.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            
            # Set page margins (in inches)
            ws.page_margins = PageMargins(
                left=0.7,
                right=0.7,
                top=0.75,
                bottom=0.75,
                header=0.3,
                footer=0.3
            )
            
            # Set print area to cover all content
            ws.print_options.horizontalCentered = True
            
            # Set column widths for A4
            ws.column_dimensions['A'].width = 12
            ws.column_dimensions['B'].width = 20
            ws.column_dimensions['C'].width = 20
            ws.column_dimensions['D'].width = 18
            ws.column_dimensions['E'].width = 18
            
            # Set default row height
            ws.row_dimensions[1].height = 80  # Header row
    
    def _define_styles(self):
        """Define all styles according to specifications"""
        return {
            'title_main': {
                'font': Font(name='Times New Roman', bold=True, size=12),
                'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True)
            },
            'cover_note': {
                'font': Font(name='Times New Roman', bold=True, size=12),
                'alignment': Alignment(horizontal='center', vertical='center'),
                'fill': PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            },
            'info_label': {
                'font': Font(name='Times New Roman', bold=True, size=12),
                'alignment': Alignment(horizontal='left', vertical='center')
            },
            'info_value': {
                'font': Font(name='Times New Roman', size=12),
                'alignment': Alignment(horizontal='left', vertical='center')
            },
            'table_header': {
                'fill': PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"),
                'font': Font(name='Times New Roman', bold=True, size=11),
                'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
                'border': Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
            },
            'table_data': {
                'font': Font(name='Times New Roman', size=11),
                'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
                'border': Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
            },
            'table_data_left': {
                'font': Font(name='Times New Roman', size=11),
                'alignment': Alignment(horizontal='left', vertical='center', wrap_text=True),
                'border': Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
            },
            'underline_title': {
                'font': Font(name='Times New Roman', bold=True, size=12, underline='single'),
                'alignment': Alignment(horizontal='left', vertical='center')
            },
            'reference': {
                'font': Font(name='Times New Roman', size=11),
                'alignment': Alignment(horizontal='left', vertical='center', wrap_text=True)
            },
            'method_title': {
                'font': Font(name='Traditional Arabic', bold=True, size=11),
                'alignment': Alignment(horizontal='left', vertical='center')
            },
            'method_text': {
                'font': Font(name='Times New Roman', size=10),
                'alignment': Alignment(horizontal='left', vertical='center', wrap_text=True)
            },
            'footer_text': {
                'font': Font(name='Times New Roman', size=10),
                'alignment': Alignment(horizontal='center', vertical='center')
            },
            'arabic_header': {
                'font': Font(name='Times New Roman', size=14),
                'alignment': Alignment(horizontal='left', vertical='center')
            },
            'table_criteria_header': {
                'fill': PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"),
                'font': Font(name='Times New Roman', bold=True, size=11),
                'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
                'border': Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
            },
            'table_criteria_data': {
                'font': Font(name='Times New Roman', bold=True, size=11),
                'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
                'border': Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
            }
        }
    
    def add_header(self, ws):
        """Add header with logo on left and Arabic text on right"""
        # Logo on left side (Column A)
        if os.path.exists(self.logo_path):
            try:
                img = Image(self.logo_path)
                # Resize logo to fit nicely
                img.width = 120
                img.height = 60
                ws.add_image(img, 'A1')
            except Exception as e:
                print(f"Error adding logo: {e}")
                ws['A1'] = "[SASOmed Logo]"
                ws['A1'].font = Font(name='Times New Roman', size=10)
        else:
            ws['A1'] = "[SASOmed Logo]"
            ws['A1'].font = Font(name='Times New Roman', size=10)
        
        # Arabic text on right side (Column E)
        arabic_cell = ws['E1']
        arabic_cell.value = "الاقليمية للصناعات الطبية و المخبرية"
        arabic_cell.font = Font(name='Times New Roman', size=14, bold=True)
        arabic_cell.alignment = Alignment(horizontal='right', vertical='center')
        
        # Merge cells for better layout
        ws.merge_cells('A1:B1')  # Logo area
        ws.merge_cells('D1:E1')  # Arabic text area
        
        # Add a separator line below header
        for col in range(1, 6):
            cell = ws.cell(row=2, column=col)
            cell.border = Border(bottom=Side(style='thin'))
        
        return 3  # Return starting row after header
    
    def add_footer(self, ws, page_num, total_pages):
        """Add footer with technical manager and contact information"""
        # Find last row with content
        last_row = ws.max_row + 2
        
        # Technical Manager line
        ws.merge_cells(f'A{last_row}:B{last_row}')
        tech_manager_cell = ws[f'A{last_row}']
        tech_manager_cell.value = "Technical Manager"
        tech_manager_cell.font = Font(name='Times New Roman', size=10, bold=True)
        tech_manager_cell.alignment = Alignment(horizontal='left')
        
        ws.merge_cells(f'C{last_row}:E{last_row}')
        sign_cell = ws[f'C{last_row}']
        sign_cell.value = "Sign and Date: _____________"
        sign_cell.font = Font(name='Times New Roman', size=10)
        sign_cell.alignment = Alignment(horizontal='right')
        
        last_row += 1
        
        # Page x of y
        ws.merge_cells(f'A{last_row}:E{last_row}')
        page_cell = ws[f'A{last_row}']
        page_cell.value = f"Page {page_num} of {total_pages}"
        page_cell.font = Font(name='Times New Roman', size=10)
        page_cell.alignment = Alignment(horizontal='center')
        
        last_row += 1
        
        # Contact information table (invisible borders except top)
        # First row of contact info
        col_positions = [1, 3, 5]  # A, C, E columns
        contact_texts = ["Amman - Jordan", "Fax: +962 6 5829665", "Tel: +962 6 5829658"]
        
        for col, text in zip(col_positions, contact_texts):
            col_letter = get_column_letter(col)
            cell = ws[f'{col_letter}{last_row}']
            cell.value = text
            cell.font = Font(name='Times New Roman', size=10)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            # Add only top border
            cell.border = Border(top=Side(style='medium'))
        
        last_row += 1
        
        # Second row of contact info - Postal code
        postal_cell = ws[f'A{last_row}']
        postal_cell.value = "Postal code: 11814"
        postal_cell.font = Font(name='Times New Roman', size=10)
        postal_cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Add top border for postal code as well to create separation
        postal_cell.border = Border(top=Side(style='thin'))
        
        return last_row
    
    def create_cover_page(self, data):
        """Create cover page with all specifications"""
        # Add header
        self.add_header(self.ws_cover)
        
        current_row = 4
        
        # Title
        self.ws_cover.merge_cells(f'A{current_row}:E{current_row}')
        cell = self.ws_cover[f'A{current_row}']
        cell.value = "Microbiological Examination of Nonsterile Products: Microbial Enumeration Tests\nDetermines the Total Population of Aerobic Bacteria and Yeast and Molds in the Product (Bioburden test)"
        cell.font = self.styles['title_main']['font']
        cell.alignment = self.styles['title_main']['alignment']
        current_row += 2
        
        # Cover note with grey background
        self.ws_cover.merge_cells(f'A{current_row}:E{current_row}')
        cell = self.ws_cover[f'A{current_row}']
        cell.value = "The cover page is an integral part of this test report"
        cell.font = self.styles['cover_note']['font']
        cell.fill = self.styles['cover_note']['fill']
        cell.alignment = self.styles['cover_note']['alignment']
        current_row += 2
        
        # Sample Information
        info_fields = [
            ("Sample Receiving Date:", data.get('received_date', '')),
            ("Test Performing Date:", data.get('test_performing_date', '')),
            ("Issuing Date:", data.get('issuing_date', '')),
            ("Customer Name:", data.get('customer_name', '')),
            ("Sample condition:", "Accepted")
        ]
        
        for label, value in info_fields:
            # Label
            label_cell = self.ws_cover.cell(row=current_row, column=1)
            label_cell.value = label
            label_cell.font = self.styles['info_label']['font']
            
            # Value
            value_cell = self.ws_cover.cell(row=current_row, column=2)
            value_cell.value = value
            value_cell.font = self.styles['info_value']['font']
            if "Accepted" in value:
                value_cell.font = Font(name='Times New Roman', size=12, color="008000")
            
            current_row += 1
        
        current_row += 1
        
        # Product Information Table
        # Table headers - without side borders
        headers = ["Product Description", "Sample ID", "Product code"]
        for col, header in enumerate(headers, 1):
            cell = self.ws_cover.cell(row=current_row, column=col)
            cell.value = header
            cell.font = self.styles['table_header']['font']
            cell.fill = self.styles['table_header']['fill']
            cell.alignment = self.styles['table_header']['alignment']
            # Border without sides (only top and bottom)
            cell.border = Border(
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        
        current_row += 1
        
        # Table data
        product_code = f"{data.get('sample_batch_no', '')}\n{data.get('reference_no', '')}"
        row_data = [
            data.get('product_description', ''),
            data.get('sample_id', ''),
            product_code
        ]
        
        for col, value in enumerate(row_data, 1):
            cell = self.ws_cover.cell(row=current_row, column=col)
            cell.value = value
            cell.font = self.styles['table_data_left']['font']
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            # Border without sides
            cell.border = Border(
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        
        # Set row height
        self.ws_cover.row_dimensions[current_row].height = 50
        
        # Add footer
        self.add_footer(self.ws_cover, 1, 2)
        
        return current_row
    
    def create_content_page(self, data):
        """Create content page with test results"""
        # Add header
        self.add_header(self.ws_content)
        
        current_row = 4
        
        # Test Results title
        self.ws_content.merge_cells(f'A{current_row}:E{current_row}')
        cell = self.ws_content[f'A{current_row}']
        cell.value = "Test Results:"
        cell.font = self.styles['underline_title']['font']
        current_row += 2
        
        # Test Results Table with full borders
        headers = ["Sample Identification", "Total Aerobic Microbial Count CFU/ml", "Total Combined Yeasts/Molds Count CFU/ml"]
        
        # Add headers with full borders
        for col, header in enumerate(headers, 1):
            cell = self.ws_content.cell(row=current_row, column=col)
            cell.value = header
            cell.font = self.styles['table_header']['font']
            cell.fill = self.styles['table_header']['fill']
            cell.alignment = self.styles['table_header']['alignment']
            cell.border = self.styles['table_header']['border']
        
        current_row += 1
        
        # Results data
        results = data.get('results', [])
        if isinstance(results, dict):
            results = [results]
        
        for result in results:
            tamc = str(result.get('tamc_result', ''))
            tymc = str(result.get('tymc_result', ''))
            batch_no = data.get('sample_batch_no', '')
            
            # Check if result is 0
            if tamc == '0' or tamc == '0.0' or tamc.lower() == 'zero':
                tamc = f"No microbial growth was detected for batch number {batch_no}"
            
            if tymc == '0' or tymc == '0.0' or tymc.lower() == 'zero':
                tymc = f"No microbial growth was detected for batch number {batch_no}"
            
            row_data = [
                result.get('sample_id', ''),
                tamc,
                tymc
            ]
            
            for col, value in enumerate(row_data, 1):
                cell = self.ws_content.cell(row=current_row, column=col)
                cell.value = value
                cell.font = self.styles['table_data']['font']
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = self.styles['table_data']['border']
            
            current_row += 1
        
        current_row += 2
        
        # Acceptance criteria text
        criteria_text = [
            "Acceptance criteria for nonsterile pharmaceutical products based upon the total aerobic microbial count",
            "(TAMC) and the total combined yeasts and molds count (TYMC) are given in Table 1.",
            "",
            "Acceptance criteria are based on 〈1111〉 MICROBIOLOGICAL individual results",
            "",
            "When an acceptance criterion for microbiological quality is prescribed, it is interpreted as follows:",
            "• 10¹ cfu: maximum acceptable count = 20",
            "• 10² cfu: maximum acceptable count = 200",
            "• 10³ cfu: maximum acceptable count = 2000"
        ]
        
        for line in criteria_text:
            if line:
                self.ws_content.merge_cells(f'A{current_row}:E{current_row}')
                cell = self.ws_content[f'A{current_row}']
                cell.value = line
                cell.font = self.styles['reference']['font']
                cell.alignment = Alignment(horizontal='left')
                current_row += 1
            else:
                current_row += 1
        
        current_row += 1
        
        # Table 1 title
        self.ws_content.merge_cells(f'A{current_row}:E{current_row}')
        cell = self.ws_content[f'A{current_row}']
        cell.value = "Table 1: Acceptance Criteria for Microbiological Quality of Nonsterile Substances for Pharmaceutical Use"
        cell.font = Font(name='Times New Roman', size=11)
        cell.alignment = Alignment(horizontal='center')
        current_row += 1
        
        # Table 1 headers
        table_headers = ["", "Total Aerobic Microbial Count (cfu/g or cfu/ml)", "Total Combined Yeasts/Molds Count (cfu/g or cfu/ml)"]
        
        for col, header in enumerate(table_headers, 1):
            cell = self.ws_content.cell(row=current_row, column=col)
            cell.value = header
            cell.font = self.styles['table_criteria_header']['font']
            cell.fill = self.styles['table_criteria_header']['fill']
            cell.alignment = self.styles['table_criteria_header']['alignment']
            cell.border = self.styles['table_criteria_header']['border']
        
        current_row += 1
        
        # Table 1 data
        table_data = ["Substance for Pharmaceutical Use", "10³", "10²"]
        
        for col, value in enumerate(table_data, 1):
            cell = self.ws_content.cell(row=current_row, column=col)
            cell.value = value
            cell.font = self.styles['table_criteria_data']['font']
            cell.alignment = Alignment(horizontal='center')
            cell.border = self.styles['table_criteria_data']['border']
        
        current_row += 2
        
        # Reference
        self.ws_content.merge_cells(f'A{current_row}:E{current_row}')
        cell = self.ws_content[f'A{current_row}']
        cell.value = "Reference: 〈1111〉 Microbiological Examination / General Information"
        cell.font = Font(name='Times New Roman', size=11)
        cell.alignment = Alignment(horizontal='left')
        current_row += 2
        
        # Test Method
        cell = self.ws_content.cell(row=current_row, column=1)
        cell.value = "Test Method:"
        cell.font = Font(name='Traditional Arabic', bold=True, size=11)
        current_row += 1
        
        self.ws_content.merge_cells(f'A{current_row}:E{current_row}')
        cell = self.ws_content[f'A{current_row}']
        cell.value = "ISO 11737-1 Sterilization of health care products – Microbiological methods – Part 1: Determination of the population\nof microorganisms on product, and USP 〈61〉 \"Bioburden\" or \"Microbial Limits\" test."
        cell.font = Font(name='Times New Roman', size=10)
        cell.alignment = Alignment(horizontal='left', wrap_text=True)
        
        # Add footer
        self.add_footer(self.ws_content, 2, 2)
    
    def generate_report(self, data):
        """Generate complete report"""
        self.create_cover_page(data)
        self.create_content_page(data)
        
        return self.wb


# Streamlit Application
def main():
    st.set_page_config(
        page_title="Bioburden Test Report Generator",
        page_icon="🔬",
        layout="wide"
    )
    
    st.title("🔬 Bioburden Test Report Generator")
    st.markdown("Generates professional A4 laboratory reports with SASOmed logo")
    
    # Check for logo file
    logo_exists = os.path.exists("SASOmed_Logo.png")
    if not logo_exists:
        st.warning("⚠️ SASOmed_Logo.png not found in the current directory. Please ensure the logo file is present.")
    
    with st.form("report_form"):
        st.subheader("📅 Sample Information")
        col1, col2 = st.columns(2)
        
        with col1:
            received_date = st.date_input("Sample Receiving Date", datetime.now())
            test_performing_date = st.date_input("Test Performing Date", datetime.now())
            issuing_date = st.date_input("Issuing Date", datetime.now())
        
        with col2:
            customer_name = st.text_input("Customer Name")
        
        st.subheader("🏭 Product Information")
        col3, col4 = st.columns(2)
        
        with col3:
            product_description = st.text_area("Product Description", height=80)
            sample_id = st.text_input("Sample ID")
        
        with col4:
            sample_batch_no = st.text_input("Sample Batch No.")
            reference_no = st.text_input("Reference No.")
        
        st.subheader("🧪 Test Results")
        col5, col6 = st.columns(2)
        
        with col5:
            tamc_result = st.text_input("Total Aerobic Microbial Count (TAMC) CFU/ml", 
                                        placeholder="e.g., <10, 50, 100, or 0")
        with col6:
            tymc_result = st.text_input("Total Combined Yeasts/Molds Count (TYMC) CFU/ml",
                                        placeholder="e.g., <10, 20, 50, or 0")
        
        submitted = st.form_submit_button("Generate Report", use_container_width=True)
    
    if submitted:
        if not product_description or not sample_id:
            st.error("❌ Please fill in Product Description and Sample ID")
            return
        
        # Prepare data
        report_data = {
            'received_date': received_date.strftime('%Y-%m-%d'),
            'test_performing_date': test_performing_date.strftime('%Y-%m-%d'),
            'issuing_date': issuing_date.strftime('%Y-%m-%d'),
            'customer_name': customer_name or "_____________",
            'product_description': product_description,
            'sample_id': sample_id,
            'sample_batch_no': sample_batch_no or "_____________",
            'reference_no': reference_no or "_____________",
            'results': [{
                'sample_id': sample_id,
                'tamc_result': tamc_result or "_____________",
                'tymc_result': tymc_result or "_____________"
            }]
        }
        
        with st.spinner("Generating professional A4 report..."):
            generator = BioburdenReportGenerator("SASOmed_Logo.png")
            workbook = generator.generate_report(report_data)
            
            output = BytesIO()
            workbook.save(output)
            output.seek(0)
        
        st.success("✅ Report generated successfully!")
        
        st.download_button(
            label="📥 Download Excel Report (A4 Format)",
            data=output,
            file_name=f"Bioburden_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        # Show preview
        with st.expander("📄 Report Specifications"):
            st.markdown("""
            **Report Format: A4 Paper Size**
            
            **Header (Both Pages):**
            - Left: SASOmed Logo
            - Right: Arabic text "الاقليمية للصناعات الطبية و المخبرية"
            
            **Footer (Both Pages):**
            - Technical Manager with sign/date field
            - Page numbers (1 of 2, 2 of 2)
            - Contact information with top border only
            
            **Cover Page:**
            - Main title (bold, centered, Times New Roman 12)
            - Grey background note
            - Sample information (bold labels)
            - Product table (no side borders)
            
            **Test Report Page:**
            - Test results with "No microbial growth" for zero results
            - Complete acceptance criteria
            - Table 1 with microbiological limits
            - Test method (Traditional Arabic title, Times New Roman text)
            """)

if __name__ == "__main__":
    main()