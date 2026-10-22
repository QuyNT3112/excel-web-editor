import io
import json
import unittest
import openpyxl
from openpyxl.styles import PatternFill
from app import app

class TestExcelExportRigorous(unittest.TestCase):
    def setUp(self):
        self.app = app.test_client()
        self.app.testing = True

    def test_complete_xlsx_export_and_inspection(self):
        """Rigorous test for Excel XLSX export: verifying sheet structure, formulas, data types, styles and readability."""
        
        # 1. Create a rich spreadsheet payload
        payload = {
            'filename': 'Kiem_Tra_Xuat_Excel_2026.xlsx',
            'sheets': {
                'DoanhThu': [
                    ['Tháng', 'Doanh Thu', 'Chi Phí', 'Lợi Nhuận', 'Đánh Giá'],
                    ['Tháng 1', 1500, 900, '=B2-C2', '=IF(D2>500,"Tốt","Kém")'],
                    ['Tháng 2', 2200, 1100, '=B3-C3', '=IF(D3>500,"Tốt","Kém")'],
                    ['Tháng 3', 3400, 1500, '=B4-C4', '=IF(D4>500,"Tốt","Kém")'],
                    ['TỔNG CỘNG', '=SUM(B2:B4)', '=SUM(C2:C4)', '=SUM(D2:D4)', 'HOÀN THÀNH']
                ],
                'GhiChu': [
                    ['Mục', 'Nội Dung', 'Người Lập'],
                    ['Kế toán', 'Số liệu đã đối soát ngân hàng', 'Nguyễn Văn A'],
                    ['Kiểm toán', 'Không phát hiện sai lệch', 'Trần Thị B']
                ]
            },
            'styles': {
                'DoanhThu': {
                    '0,0': {'bold': True, 'bg': '#107C41', 'color': '#FFFFFF', 'align': 'center'},
                    '0,1': {'bold': True, 'bg': '#107C41', 'color': '#FFFFFF', 'align': 'center'},
                    '0,2': {'bold': True, 'bg': '#107C41', 'color': '#FFFFFF', 'align': 'center'},
                    '0,3': {'bold': True, 'bg': '#107C41', 'color': '#FFFFFF', 'align': 'center'},
                    '0,4': {'bold': True, 'bg': '#107C41', 'color': '#FFFFFF', 'align': 'center'},
                    '4,0': {'bold': True, 'color': '#107C41'},
                    '4,1': {'bold': True},
                    '4,2': {'bold': True},
                    '4,3': {'bold': True}
                }
            },
            'colWidths': {
                'DoanhThu': {
                    'A': 15.0,
                    'B': 18.0,
                    'C': 18.0,
                    'D': 18.0,
                    'E': 20.0
                }
            }
        }

        # 2. Perform POST request to /download
        response = self.app.post('/download', json=payload, content_type='application/json')
        self.assertEqual(response.status_code, 200, f"Download failed with status: {response.status_code}")
        
        # Verify headers
        self.assertEqual(response.mimetype, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        self.assertIn('attachment', response.headers.get('Content-Disposition', ''))
        self.assertIn('Kiem_Tra_Xuat_Excel_2026.xlsx', response.headers.get('Content-Disposition', ''))

        # 3. Read exported bytes back into OpenPyXL
        buffer = io.BytesIO(response.data)
        wb = openpyxl.load_workbook(buffer, data_only=False)

        # Verify sheets
        self.assertEqual(wb.sheetnames, ['DoanhThu', 'GhiChu'], "Sheet names mismatch in exported Excel file")
        
        # Inspect Sheet 1: DoanhThu
        ws1 = wb['DoanhThu']
        self.assertEqual(ws1.cell(row=1, column=1).value, 'Tháng')
        self.assertEqual(ws1.cell(row=1, column=2).value, 'Doanh Thu')
        
        # Verify Header Styles
        h_cell = ws1.cell(row=1, column=1)
        self.assertTrue(h_cell.font.bold, "Header cell font should be bold")
        self.assertIsNotNone(h_cell.fill, "Header cell should have fill color")
        self.assertEqual(h_cell.alignment.horizontal, 'center')

        # Verify numeric types
        self.assertEqual(ws1.cell(row=2, column=2).value, 1500)
        self.assertIsInstance(ws1.cell(row=2, column=2).value, int, "Numeric value should be parsed as integer")
        self.assertEqual(ws1.cell(row=3, column=2).value, 2200)

        # Verify Formulas are preserved verbatim
        self.assertEqual(ws1.cell(row=2, column=4).value, '=B2-C2', "Cell formula was not preserved")
        self.assertEqual(ws1.cell(row=2, column=5).value, '=IF(D2>500,"Tốt","Kém")', "IF formula was not preserved")
        self.assertEqual(ws1.cell(row=5, column=2).value, '=SUM(B2:B4)', "SUM formula was not preserved")
        self.assertEqual(ws1.cell(row=5, column=3).value, '=SUM(C2:C4)')
        self.assertEqual(ws1.cell(row=5, column=4).value, '=SUM(D2:D4)')

        # Verify Column Dimensions
        self.assertEqual(ws1.column_dimensions['A'].width, 15.0)
        self.assertEqual(ws1.column_dimensions['B'].width, 18.0)

        # Inspect Sheet 2: GhiChu
        ws2 = wb['GhiChu']
        self.assertEqual(ws2.cell(row=1, column=1).value, 'Mục')
        self.assertEqual(ws2.cell(row=2, column=2).value, 'Số liệu đã đối soát ngân hàng')
        self.assertEqual(ws2.cell(row=3, column=3).value, 'Trần Thị B')

        # 4. Save to disk to inspect file size and physical binary validity
        with open('test_output_verified.xlsx', 'wb') as f:
            f.write(response.data)

        # Re-open file from disk to ensure operating system / Excel compatibility
        reopened_wb = openpyxl.load_workbook('test_output_verified.xlsx', data_only=False)
        self.assertIn('DoanhThu', reopened_wb.sheetnames)
        self.assertEqual(reopened_wb['DoanhThu'].cell(row=5, column=2).value, '=SUM(B2:B4)')
        
        print("\n>>> EXPORT TEST PASSED: File generated and validated 100% successfully! <<<")

if __name__ == '__main__':
    unittest.main()
