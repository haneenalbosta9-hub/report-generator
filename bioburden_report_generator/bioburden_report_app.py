import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import streamlit as st
from io import BytesIO

class BioburdenReportGenerator:
    def __init__(self):
        self.wb = openpyxl.Workbook()
        self.ws = self.wb.active
        self.ws.title = "Bioburden Test Report"
        self.current_row = 1
        
        # Define styles
        self.styles = self._define_styles()
        
    def _define_styles(self):
        """Define all styles matching the Word template"""
        return {
            'title': {
                'font': Font(name='Arial', bold=True, size=14),
                'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True)
            },
            'subtitle': {
                'font': Font(name='Arial', bold=True, size=11),
                'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True)
            },
            'header': {
                'fill': PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"),
                'font': Font(name='Arial', bold=True, size=10),
                'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
                'border': Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
            },
            'data': {
                'font': Font(name='Arial', size=10),
                'alignment': Alignment(horizontal='left', vertical='center', wrap_text=True),
                'border': Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
            },
            'center_data': {
                'font': Font(name='Arial', size=10),
                'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
                'border': Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
            },
            'label': {
                'font': Font(name='Arial', bold=True, size=10),
                'alignment': Alignment(horizontal='left', vertical='center'),
                'fill': PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            },
            'underline': {
                'font': Font(name='Arial', size=10, underline='single'),
                'alignment': Alignment(horizontal='left', vertical='center')
            },
            'reference': {
                'font': Font(name='Arial', italic=True, size=9),
                'alignment': Alignment(horizontal='left', vertical='center', wrap_text=True)
            },
            'bold': {
                'font': Font(name='Arial', bold=True, size=10),
                'alignment': Alignment(horizontal='left', vertical='center')
            }
        }
    
    def set_column_widths(self):
        """Set column widths for professional layout"""
        self.ws.column_dimensions['A'].width = 15
        self.ws.column_dimensions['B'].width = 30
        self.ws.column_dimensions['C'].width = 30
        self.ws.column_dimensions['D'].width = 25
    
    def add_title(self):
        """Add main title"""
        self.ws.merge_cells('A1:D1')
        cell = self.ws['A1']
        cell.value = "Microbiological Examination of Nonsterile Products: Microbial Enumeration Tests"
        cell.font = self.styles['title']['font']
        cell.alignment = self.styles['title']['alignment']
        
        self.current_row = 2
        
        self.ws.merge_cells('A2:D2')
        cell = self.ws['A2']
        cell.value = "Determines the Total Population of Aerobic Bacteria and Yeast and Molds in the Product (Bioburden test)"
        cell.font = self.styles['subtitle']['font']
        cell.alignment = self.styles['subtitle']['alignment']
        
        self.current_row = 3
        
        self.ws.merge_cells('A3:D3')
        cell = self.ws['A3']
        cell.value = "The cover page is an integral part of this test report"
        cell.font = Font(name='Arial', italic=True, size=9)
        cell.alignment = Alignment(horizontal='center')
        
        self.current_row = 5
    
    def add_sample_info(self, data):
        """Add sample information section"""
        # Sample Receiving Date
        self.ws.cell(row=self.current_row, column=1, value="Sample Receiving Date:")
        self.ws.cell(row=self.current_row, column=1).font = self.styles['bold']['font']
        self.ws.merge_cells(f'B{self.current_row}:D{self.current_row}')
        self.ws.cell(row=self.current_row, column=2, value=data.get('received_date', ''))
        self.ws.cell(row=self.current_row, column=2).font = self.styles['underline']['font']
        self.current_row += 1
        
        # Test Performing Date
        self.ws.cell(row=self.current_row, column=1, value="Test Performing Date:")
        self.ws.cell(row=self.current_row, column=1).font = self.styles['bold']['font']
        self.ws.merge_cells(f'B{self.current_row}:D{self.current_row}')
        self.ws.cell(row=self.current_row, column=2, value=data.get('test_performing_date', ''))
        self.ws.cell(row=self.current_row, column=2).font = self.styles['underline']['font']
        self.current_row += 1
        
        # Issuing Date
        self.ws.cell(row=self.current_row, column=1, value="Issuing Date:")
        self.ws.cell(row=self.current_row, column=1).font = self.styles['bold']['font']
        self.ws.merge_cells(f'B{self.current_row}:D{self.current_row}')
        self.ws.cell(row=self.current_row, column=2, value=data.get('issuing_date', ''))
        self.ws.cell(row=self.current_row, column=2).font = self.styles['underline']['font']
        self.current_row += 1
        
        # Customer Name
        self.ws.cell(row=self.current_row, column=1, value="Customer Name:")
        self.ws.cell(row=self.current_row, column=1).font = self.styles['bold']['font']
        self.ws.merge_cells(f'B{self.current_row}:D{self.current_row}')
        self.ws.cell(row=self.current_row, column=2, value=data.get('customer_name', ''))
        self.ws.cell(row=self.current_row, column=2).font = self.styles['underline']['font']
        self.current_row += 1
        
        # Sample Condition
        self.ws.cell(row=self.current_row, column=1, value="Sample condition:")
        self.ws.cell(row=self.current_row, column=1).font = self.styles['bold']['font']
        self.ws.cell(row=self.current_row, column=2, value="Accepted")
        self.ws.cell(row=self.current_row, column=2).font = Font(name='Arial', bold=True, size=10, color="008000")
        self.current_row += 2
    
    def add_product_info_table(self, data):
        """Add product information table"""
        headers = ["Product Description", "Sample ID", "Product code"]
        
        # Add headers
        for col, header in enumerate(headers, 1):
            cell = self.ws.cell(row=self.current_row, column=col)
            cell.value = header
            cell.font = self.styles['header']['font']
            cell.fill = self.styles['header']['fill']
            cell.alignment = self.styles['header']['alignment']
            cell.border = self.styles['header']['border']
        
        self.current_row += 1
        
        # Add data row
        product_desc = data.get('product_description', '')
        sample_id = data.get('sample_id', '')
        product_code = f"{data.get('sample_batch_no', '')}\n{data.get('reference_no', '')}"
        
        row_data = [product_desc, sample_id, product_code]
        
        for col, value in enumerate(row_data, 1):
            cell = self.ws.cell(row=self.current_row, column=col)
            cell.value = value
            cell.font = self.styles['data']['font']
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            cell.border = self.styles['data']['border']
        
        self.current_row += 2
    
    def add_test_results_table(self, results_data):
        """Add test results table"""
        # Title
        self.ws.merge_cells(f'A{self.current_row}:D{self.current_row}')
        cell = self.ws[f'A{self.current_row}']
        cell.value = "Test Results:"
        cell.font = Font(name='Arial', bold=True, size=11, underline='single')
        self.current_row += 1
        
        # Headers
        headers = ["Sample Identification", "Total Aerobic Microbial Count CFU/ml", "Total Combined Yeasts/Molds Count CFU/ml"]
        
        for col, header in enumerate(headers, 1):
            cell = self.ws.cell(row=self.current_row, column=col)
            cell.value = header
            cell.font = self.styles['header']['font']
            cell.fill = self.styles['header']['fill']
            cell.alignment = self.styles['header']['alignment']
            cell.border = self.styles['header']['border']
        
        self.current_row += 1
        
        # Add result rows
        for result in results_data:
            row_data = [
                result.get('sample_id', ''),
                result.get('tamc_result', ''),
                result.get('tymc_result', '')
            ]
            
            for col, value in enumerate(row_data, 1):
                cell = self.ws.cell(row=self.current_row, column=col)
                cell.value = value
                cell.font = self.styles['data']['font']
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = self.styles['data']['border']
            
            self.current_row += 1
        
        self.current_row += 1
    
    def add_acceptance_criteria(self):
        """Add acceptance criteria information"""
        # Criteria text
        text_lines = [
            "Acceptance criteria for nonsterile pharmaceutical products based upon the total aerobic microbial count (TAMC) and the total combined yeasts and molds count (TYMC) are given in Table 1.",
            "",
            "Acceptance criteria are based on 〈1111〉 MICROBIOLOGICAL individual results",
            "",
            "When an acceptance criterion for microbiological quality is prescribed, it is interpreted as follows:",
            "• 10¹ cfu: maximum acceptable count = 20",
            "• 10² cfu: maximum acceptable count = 200",
            "• 10³ cfu: maximum acceptable count = 2000"
        ]
        
        for line in text_lines:
            if line:
                self.ws.merge_cells(f'A{self.current_row}:D{self.current_row}')
                cell = self.ws[f'A{self.current_row}']
                cell.value = line
                cell.font = self.styles['reference']['font']
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                self.current_row += 1
            else:
                self.current_row += 1
        
        self.current_row += 1
    
    def add_table_1(self):
        """Add Table 1: Acceptance Criteria"""
        # Table title
        self.ws.merge_cells(f'A{self.current_row}:D{self.current_row}')
        cell = self.ws[f'A{self.current_row}']
        cell.value = "Table 1: Acceptance Criteria for Microbiological Quality of Nonsterile Substances for Pharmaceutical Use"
        cell.font = Font(name='Arial', bold=True, size=10)
        cell.alignment = Alignment(horizontal='center')
        self.current_row += 1
        
        # Table headers
        headers = ["", "Total Aerobic Microbial Count (cfu/g or cfu/ml)", "Total Combined Yeasts/Molds Count (cfu/g or cfu/ml)"]
        
        for col, header in enumerate(headers, 1):
            cell = self.ws.cell(row=self.current_row, column=col)
            cell.value = header
            cell.font = self.styles['header']['font']
            cell.fill = self.styles['header']['fill']
            cell.alignment = self.styles['header']['alignment']
            cell.border = self.styles['header']['border']
        
        self.current_row += 1
        
        # Table data
        data = ["Substance for Pharmaceutical Use", "10³", "10²"]
        
        for col, value in enumerate(data, 1):
            cell = self.ws.cell(row=self.current_row, column=col)
            cell.value = value
            cell.font = self.styles['data']['font']
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = self.styles['data']['border']
        
        self.current_row += 2
        
        # Reference
        self.ws.merge_cells(f'A{self.current_row}:D{self.current_row}')
        cell = self.ws[f'A{self.current_row}']
        cell.value = "Reference: 〈1111〉 Microbiological Examination / General Information"
        cell.font = Font(name='Arial', italic=True, size=9)
        cell.alignment = Alignment(horizontal='left')
        self.current_row += 2
    
    def add_test_method(self):
        """Add test method section"""
        self.ws.merge_cells(f'A{self.current_row}:D{self.current_row}')
        cell = self.ws[f'A{self.current_row}']
        cell.value = "Test Method:"
        cell.font = Font(name='Arial', bold=True, size=10)
        self.current_row += 1
        
        self.ws.merge_cells(f'A{self.current_row}:D{self.current_row}')
        cell = self.ws[f'A{self.current_row}']
        cell.value = "ISO 11737-1 Sterilization of health care products -- Microbiological methods -- Part 1: Determination of the population of microorganisms on product, and USP 〈61〉 \"Bioburden\" or \"Microbial Limits\" test."
        cell.font = Font(name='Arial', size=9)
        cell.alignment = Alignment(horizontal='left', wrap_text=True)
        self.current_row += 2
    
    def add_end_of_report(self):
        """Add end of report marker"""
        self.ws.merge_cells(f'A{self.current_row}:D{self.current_row}')
        cell = self.ws[f'A{self.current_row}']
        cell.value = "End of Report"
        cell.font = Font(name='Arial', bold=True, size=10, italic=True)
        cell.alignment = Alignment(horizontal='center')
    
    def generate_report(self, data):
        """Generate complete report"""
        self.set_column_widths()
        self.add_title()
        self.add_sample_info(data)
        self.add_product_info_table(data)
        
        # Convert results data to list format if it's a single result
        results = data.get('results', [])
        if isinstance(results, dict):
            results = [results]
        
        self.add_test_results_table(results)
        self.add_acceptance_criteria()
        self.add_table_1()
        self.add_test_method()
        self.add_end_of_report()
        
        return self.wb

# Streamlit Application
def main():
    st.title("Bioburden Test Report Generator")
    st.write("Generate professional Excel reports from microbiological test data")
    
    with st.form("report_form"):
        st.subheader("Sample Information")
        
        # Sample Information
        received_date = st.date_input("Sample Receiving Date", datetime.now())
        test_performing_date = st.date_input("Test Performing Date", datetime.now())
        issuing_date = st.date_input("Issuing Date", datetime.now())
        customer_name = st.text_input("Customer Name")
        
        st.subheader("Product Information")
        product_description = st.text_area("Product Description")
        sample_id = st.text_input("Sample ID")
        sample_batch_no = st.text_input("Sample Batch No.")
        reference_no = st.text_input("Reference No.")
        
        st.subheader("Test Results")
        st.write("Add test results for each sample")
        
        # Initialize session state for multiple results
        if 'result_count' not in st.session_state:
            st.session_state.result_count = 1
            st.session_state.results = []
        
        # Dynamic form for results
        results = []
        for i in range(st.session_state.result_count):
            st.write(f"**Sample {i+1}**")
            col1, col2, col3 = st.columns(3)
            with col1:
                sample_id_result = st.text_input(f"Sample ID", key=f"sample_id_{i}")
            with col2:
                tamc_result = st.text_input(f"TAMC Result (CFU/ml)", key=f"tamc_{i}")
            with col3:
                tymc_result = st.text_input(f"TYMC Result (CFU/ml)", key=f"tymc_{i}")
            
            results.append({
                'sample_id': sample_id_result,
                'tamc_result': tamc_result,
                'tymc_result': tymc_result
            })
        
        col1, col2 = st.columns(2)
        with col1:
            if st.form_submit_button("Add Another Sample"):
                st.session_state.result_count += 1
                st.experimental_rerun()
        
        submit = st.form_submit_button("Generate Report")
    
    if submit:
        # Prepare data
        report_data = {
            'received_date': received_date.strftime('%Y-%m-%d'),
            'test_performing_date': test_performing_date.strftime('%Y-%m-%d'),
            'issuing_date': issuing_date.strftime('%Y-%m-%d'),
            'customer_name': customer_name,
            'product_description': product_description,
            'sample_id': sample_id,
            'sample_batch_no': sample_batch_no,
            'reference_no': reference_no,
            'results': results
        }
        
        # Generate report
        generator = BioburdenReportGenerator()
        workbook = generator.generate_report(report_data)
        
        # Save to BytesIO
        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        
        # Download button
        st.success("Report generated successfully!")
        st.download_button(
            label="Download Excel Report",
            data=output,
            file_name=f"Bioburden_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Preview option
        st.subheader("Report Preview")
        st.info("The Excel report includes:")
        st.write("✓ Complete test methodology section")
        st.write("✓ Sample and product information")
        st.write("✓ Test results table with TAMC and TYMC values")
        st.write("✓ Acceptance criteria and reference tables")
        st.write("✓ Professional formatting matching the original Word template")

if __name__ == "__main__":
    main()