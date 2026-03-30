# bioburden_report_exact.py
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.worksheet.header_footer import _HeaderFooterPart
from datetime import datetime
import streamlit as st
from io import BytesIO

class BioburdenReportGenerator:
    def __init__(self):
        self.wb = openpyxl.Workbook()
        
        # Create two worksheets
        self.ws_cover = self.wb.active
        self.ws_cover.title = "Cover Page"
        self.ws_content = self.wb.create_sheet("Test Report")
        
        self.current_row_cover = 1
        self.current_row_content = 1
        
        # Define styles exactly matching Word template
        self.styles = self._define_styles()
        
    def _define_styles(self):
        """Define styles exactly matching the Word template"""
        return {
            'title_main': {
                'font': Font(name='Times New Roman', bold=True, size=16),
                'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True)
            },
            'title_sub': {
                'font': Font(name='Times New Roman', bold=True, size=12),
                'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True)
            },
            'cover_note': {
                'font': Font(name='Times New Roman', italic=True, size=9),
                'alignment': Alignment(horizontal='center', vertical='center')
            },
            'section_header': {
                'font': Font(name='Times New Roman', bold=True, size=12),
                'alignment': Alignment(horizontal='left', vertical='center')
            },
            'label_bold': {
                'font': Font(name='Times New Roman', bold=True, size=10),
                'alignment': Alignment(horizontal='left', vertical='center')
            },
            'label_value': {
                'font': Font(name='Times New Roman', size=10),
                'alignment': Alignment(horizontal='left', vertical='center')
            },
            'table_header': {
                'fill': PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"),
                'font': Font(name='Times New Roman', bold=True, size=10),
                'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
                'border': Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
            },
            'table_data': {
                'font': Font(name='Times New Roman', size=10),
                'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
                'border': Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
            },
            'underline_title': {
                'font': Font(name='Times New Roman', bold=True, size=11, underline='single'),
                'alignment': Alignment(horizontal='left', vertical='center')
            },
            'reference': {
                'font': Font(name='Times New Roman', italic=True, size=9),
                'alignment': Alignment(horizontal='left', vertical='center', wrap_text=True)
            },
            'method_text': {
                'font': Font(name='Times New Roman', size=9),
                'alignment': Alignment(horizontal='left', vertical='center', wrap_text=True)
            }
        }
    
    def set_page_setup(self):
        """Configure page setup for A4 printing"""
        for ws in [self.ws_cover, self.ws_content]:
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.page_setup.topMargin = 0.75
            ws.page_setup.bottomMargin = 0.75
            ws.page_setup.leftMargin = 0.7
            ws.page_setup.rightMargin = 0.7
            
            # Set column widths
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 25
            ws.column_dimensions['C'].width = 25
            ws.column_dimensions['D'].width = 20
    
    def add_exact_header_footer(self, ws, page_type="cover"):
        """Add exact header and footer matching Word template"""
        
        if page_type == "cover":
            # Header for cover page (empty as per Word template)
            ws.header_footer.left_header.text = ""
            ws.header_footer.center_header.text = ""
            ws.header_footer.right_header.text = ""
            
            # Footer for cover page
            ws.header_footer.left_footer.text = ""
            ws.header_footer.center_footer.text = ""
            ws.header_footer.right_footer.text = ""
            
        else:
            # Header for content pages (matching Word template)
            # Left header: empty
            ws.header_footer.left_header.text = ""
            
            # Center header: empty
            ws.header_footer.center_header.text = ""
            
            # Right header: empty
            ws.header_footer.right_header.text = ""
            
            # Footer for content pages
            # Left footer: empty
            ws.header_footer.left_footer.text = ""
            
            # Center footer: empty
            ws.header_footer.center_footer.text = ""
            
            # Right footer: empty (Word template doesn't have visible footer)
            ws.header_footer.right_footer.text = ""
    
    def create_cover_page(self, data):
        """Create cover page exactly matching Word template"""
        self.set_page_setup()
        self.add_exact_header_footer(self.ws_cover, "cover")
        
        current_row = 1
        
        # Main Title (exactly as in Word)
        self.ws_cover.merge_cells(f'A{current_row}:D{current_row}')
        cell = self.ws_cover[f'A{current_row}']
        cell.value = "Microbiological Examination of Nonsterile Products: Microbial\nEnumeration Tests"
        cell.font = self.styles['title_main']['font']
        cell.alignment = self.styles['title_main']['alignment']
        current_row += 2
        
        # Subtitle
        self.ws_cover.merge_cells(f'A{current_row}:D{current_row}')
        cell = self.ws_cover[f'A{current_row}']
        cell.value = "Determines the Total Population of Aerobic Bacteria and Yeast and\nMolds in the Product (Bioburden test)"
        cell.font = self.styles['title_sub']['font']
        cell.alignment = self.styles['title_sub']['alignment']
        current_row += 2
        
        # Cover page note
        self.ws_cover.merge_cells(f'A{current_row}:D{current_row}')
        cell = self.ws_cover[f'A{current_row}']
        cell.value = "The cover page is an integral part of this test report"
        cell.font = self.styles['cover_note']['font']
        cell.alignment = Alignment(horizontal='center')
        current_row += 3
        
        # Sample Information Section
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
            label_cell.font = self.styles['label_bold']['font']
            
            # Value
            value_cell = self.ws_cover.cell(row=current_row, column=2)
            value_cell.value = value
            value_cell.font = Font(name='Times New Roman', size=10)
            
            # For "Accepted", add specific formatting as per Word
            if "Accepted" in value:
                value_cell.font = Font(name='Times New Roman', size=10)
                # Add checkmark style as in Word template
                value_cell.value = "[] Accepted"
            
            current_row += 1
        
        current_row += 1
        
        # Product Information Table (exactly as in Word template)
        # Create table with borders matching Word
        table_start_row = current_row
        
        # Table headers
        headers = ["Product Description", "Sample ID", "Product code"]
        for col, header in enumerate(headers, 1):
            cell = self.ws_cover.cell(row=current_row, column=col)
            cell.value = header
            cell.font = self.styles['table_header']['font']
            cell.fill = self.styles['table_header']['fill']
            cell.alignment = self.styles['table_header']['alignment']
            cell.border = self.styles['table_header']['border']
        
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
            cell.font = self.styles['table_data']['font']
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            cell.border = self.styles['table_data']['border']
        
        # Add thick border around table (matching Word style)
        for row in range(table_start_row, current_row + 1):
            for col in range(1, 4):
                border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin') if row == table_start_row else Side(style='thin'),
                    bottom=Side(style='thin') if row == current_row else Side(style='thin')
                )
                self.ws_cover.cell(row=row, column=col).border = border
        
        return current_row
    
    def create_content_page(self, data):
        """Create content page exactly matching Word template"""
        self.add_exact_header_footer(self.ws_content, "content")
        
        current_row = 1
        
        # Test Results title (underlined as in Word)
        self.ws_content.merge_cells(f'A{current_row}:D{current_row}')
        cell = self.ws_content[f'A{current_row}']
        cell.value = "Test Results:"
        cell.font = self.styles['underline_title']['font']
        current_row += 2
        
        # Test Results Table (exactly as in Word template)
        headers = ["Sample Identification", "Total Aerobic Microbial Count CFU/ml", "Total Combined Yeasts/Molds Count CFU/ml"]
        
        # Add headers
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
            row_data = [
                result.get('sample_id', ''),
                result.get('tamc_result', ''),
                result.get('tymc_result', '')
            ]
            
            for col, value in enumerate(row_data, 1):
                cell = self.ws_content.cell(row=current_row, column=col)
                cell.value = value
                cell.font = self.styles['table_data']['font']
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = self.styles['table_data']['border']
            
            current_row += 1
        
        current_row += 2
        
        # Acceptance criteria text (exactly as in Word)
        criteria_text = [
            "Acceptance criteria for nonsterile pharmaceutical products based upon the total aerobic microbial count (TAMC) and the total combined yeasts and molds count (TYMC) are given in Table 1.",
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
                self.ws_content.merge_cells(f'A{current_row}:D{current_row}')
                cell = self.ws_content[f'A{current_row}']
                cell.value = line
                cell.font = self.styles['reference']['font']
                cell.alignment = Alignment(horizontal='left')
                current_row += 1
            else:
                current_row += 1
        
        current_row += 1
        
        # Table 1: Acceptance Criteria (exactly as in Word)
        self.ws_content.merge_cells(f'A{current_row}:D{current_row}')
        cell = self.ws_content[f'A{current_row}']
        cell.value = "Table 1: Acceptance Criteria for Microbiological Quality of Nonsterile Substances for Pharmaceutical Use"
        cell.font = Font(name='Times New Roman', bold=True, size=10)
        cell.alignment = Alignment(horizontal='center')
        current_row += 1
        
        # Table 1 headers
        table_headers = ["", "Total Aerobic Microbial Count (cfu/g or cfu/ml)", "Total Combined Yeasts/Molds Count (cfu/g or cfu/ml)"]
        
        for col, header in enumerate(table_headers, 1):
            cell = self.ws_content.cell(row=current_row, column=col)
            cell.value = header
            cell.font = self.styles['table_header']['font']
            cell.fill = self.styles['table_header']['fill']
            cell.alignment = self.styles['table_header']['alignment']
            cell.border = self.styles['table_header']['border']
        
        current_row += 1
        
        # Table 1 data
        table_data = ["Substance for Pharmaceutical Use", "10³", "10²"]
        
        for col, value in enumerate(table_data, 1):
            cell = self.ws_content.cell(row=current_row, column=col)
            cell.value = value
            cell.font = self.styles['table_data']['font']
            cell.alignment = Alignment(horizontal='center')
            cell.border = self.styles['table_data']['border']
        
        current_row += 2
        
        # Reference (exactly as in Word)
        self.ws_content.merge_cells(f'A{current_row}:D{current_row}')
        cell = self.ws_content[f'A{current_row}']
        cell.value = "Reference: 〈1111〉 Microbiological Examination / General Information"
        cell.font = Font(name='Times New Roman', italic=True, size=9)
        cell.alignment = Alignment(horizontal='left')
        current_row += 2
        
        # Test Method (exactly as in Word)
        self.ws_content.merge_cells(f'A{current_row}:D{current_row}')
        cell = self.ws_content[f'A{current_row}']
        cell.value = "Test Method:"
        cell.font = Font(name='Times New Roman', bold=True, size=10)
        current_row += 1
        
        self.ws_content.merge_cells(f'A{current_row}:D{current_row}')
        cell = self.ws_content[f'A{current_row}']
        cell.value = "ISO 11737-1 Sterilization of health care products -- Microbiological methods -- Part 1: Determination of the population of microorganisms on product, and USP 〈61〉 \"Bioburden\" or \"Microbial Limits\" test."
        cell.font = Font(name='Times New Roman', size=9)
        cell.alignment = Alignment(horizontal='left', wrap_text=True)
        current_row += 3
        
        # End of Report (exactly as in Word)
        self.ws_content.merge_cells(f'A{current_row}:D{current_row}')
        cell = self.ws_content[f'A{current_row}']
        cell.value = "End of Report"
        cell.font = Font(name='Times New Roman', bold=True, italic=True, size=10)
        cell.alignment = Alignment(horizontal='center')
    
    def generate_report(self, data):
        """Generate complete report matching Word template exactly"""
        self.create_cover_page(data)
        self.create_content_page(data)
        return self.wb


# Streamlit Application
def main():
    st.set_page_config(
        page_title="Bioburden Test Report Generator",
        page_icon="📋",
        layout="wide"
    )
    
    st.title("📋 Bioburden Test Report Generator")
    st.markdown("Generates reports exactly matching the Word template format")
    
    with st.form("report_form"):
        st.subheader("Sample Information")
        col1, col2 = st.columns(2)
        
        with col1:
            received_date = st.date_input("Sample Receiving Date", datetime.now())
            test_performing_date = st.date_input("Test Performing Date", datetime.now())
            issuing_date = st.date_input("Issuing Date", datetime.now())
        
        with col2:
            customer_name = st.text_input("Customer Name")
            sample_condition = st.selectbox("Sample Condition", ["Accepted", "Rejected", "Pending"])
        
        st.subheader("Product Information")
        col3, col4 = st.columns(2)
        
        with col3:
            product_description = st.text_area("Product Description", height=80)
            sample_id = st.text_input("Sample ID")
        
        with col4:
            sample_batch_no = st.text_input("Sample Batch No.")
            reference_no = st.text_input("Reference No.")
        
        st.subheader("Test Results")
        col5, col6 = st.columns(2)
        
        with col5:
            tamc_result = st.text_input("Total Aerobic Microbial Count (TAMC) CFU/ml", 
                                        placeholder="e.g., <10, 50, 100")
        with col6:
            tymc_result = st.text_input("Total Combined Yeasts/Molds Count (TYMC) CFU/ml",
                                        placeholder="e.g., <10, 20, 50")
        
        submitted = st.form_submit_button("Generate Report", use_container_width=True)
    
    if submitted:
        if not product_description or not sample_id:
            st.error("Please fill in Product Description and Sample ID")
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
        
        with st.spinner("Generating report..."):
            generator = BioburdenReportGenerator()
            workbook = generator.generate_report(report_data)
            
            output = BytesIO()
            workbook.save(output)
            output.seek(0)
        
        st.success("Report generated successfully!")
        
        st.download_button(
            label="Download Excel Report",
            data=output,
            file_name=f"Bioburden_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        st.info("""
        **Report Features:**
        - ✓ Exactly matches your Word template structure
        - ✓ Professional Times New Roman font
        - ✓ Proper table formatting with borders
        - ✓ Cover page and content page
        - ✓ Acceptance criteria and Table 1
        - ✓ Test method documentation
        - ✓ Print-ready A4 format
        """)

if __name__ == "__main__":
    main()