import io
import json
import unittest
import openpyxl
from app import app

class TestExcelWebEditor(unittest.TestCase):
    def setUp(self):
        self.app = app.test_client()
        self.app.testing = True

    def test_index_page(self):
        """Test if the main editor page renders successfully"""
        response = self.app.get('/')
        self.assertEqual(response.status_code, 200)
        self.assertIn(b'Excel Web Editor', response.data)
        self.assertIn(b'formula-bar', response.data)
        self.assertIn(b'sheet-tabs-bar', response.data)

    def test_excel_upload_and_download_flow(self):
        """Test uploading a dynamically generated Excel file, then exporting/downloading it back"""
        
        # 1. Create a dummy Excel workbook in memory
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "SalesData"
        ws1.cell(row=1, column=1, value="Month")
        ws1.cell(row=1, column=2, value="Revenue")
        ws1.cell(row=2, column=1, value="January")
        ws1.cell(row=2, column=2, value=1500)
        ws1.cell(row=3, column=1, value="February")
        ws1.cell(row=3, column=2, value=2200)
        ws1.cell(row=4, column=1, value="Total")
        ws1.cell(row=4, column=2, value="=SUM(B2:B3)")
        
        # Add another sheet
        ws2 = wb.create_sheet(title="Notes")
        ws2.cell(row=1, column=1, value="Draft note")
        
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)

        # 2. Test upload endpoint
        upload_data = {
            'file': (excel_buffer, 'test_sales.xlsx')
        }
        
        response = self.app.post('/upload', data=upload_data, content_type='multipart/form-data')
        self.assertEqual(response.status_code, 200)
        
        json_data = json.loads(response.data.decode('utf-8'))
        self.assertEqual(json_data['filename'], 'test_sales.xlsx')
        self.assertIn('SalesData', json_data['sheets'])
        self.assertIn('Notes', json_data['sheets'])
        
        # Verify uploaded cell data structure
        sales_sheet = json_data['sheets']['SalesData']
        self.assertEqual(sales_sheet[0][0], "Month")
        self.assertEqual(sales_sheet[0][1], "Revenue")
        self.assertEqual(sales_sheet[1][1], 1500)
        self.assertEqual(sales_sheet[3][1], "=SUM(B2:B3)")
        
        # Verify Notes sheet
        notes_sheet = json_data['sheets']['Notes']
        self.assertEqual(notes_sheet[0][0], "Draft note")

        # 3. Modify data simulate client edits
        # Update February revenue to 3000, and set March data in Row 4
        sales_sheet[2][1] = "3000"  # Row 3 (index 2)
        sales_sheet[3][0] = "March"  # Row 4 (index 3)
        sales_sheet[3][1] = "4000"
        # Set Total formula in Row 5 (index 4)
        sales_sheet[4][0] = "Total"
        sales_sheet[4][1] = "=SUM(B2:B4)"
        
        # 4. Test download/export endpoint
        export_payload = {
            'filename': 'updated_sales.xlsx',
            'sheets': json_data['sheets']
        }
        
        download_response = self.app.post('/download', json=export_payload, content_type='application/json')
        self.assertEqual(download_response.status_code, 200)
        self.assertEqual(download_response.mimetype, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
        # 5. Load returned bytes as workbook to verify content
        download_buffer = io.BytesIO(download_response.data)
        out_wb = openpyxl.load_workbook(download_buffer, data_only=False)
        
        out_ws = out_wb['SalesData']
        
        # Verify content conversions and formula retention
        self.assertEqual(out_ws.cell(row=1, column=1).value, "Month")
        self.assertEqual(out_ws.cell(row=2, column=2).value, 1500) # kept as integer
        self.assertEqual(out_ws.cell(row=3, column=2).value, 3000) # parsed to integer
        self.assertEqual(out_ws.cell(row=4, column=2).value, 4000) # parsed to integer
        self.assertEqual(out_ws.cell(row=5, column=2).value, "=SUM(B2:B4)") # formula saved correctly
        
        out_notes = out_wb['Notes']
        self.assertEqual(out_notes.cell(row=1, column=1).value, "Draft note")
        
        print("SUCCESS: All backend verification tests passed successfully!")

if __name__ == '__main__':
    unittest.main()
