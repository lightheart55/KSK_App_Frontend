from http.server import BaseHTTPRequestHandler
import json
import urllib.parse
import requests
import openpyxl
import io
import os

GAS_API_URL = "https://script.google.com/macros/s/AKfycbx6cQmMlTmOxsnX8nSbVTp2T3Pa_MYwVocH7FFI7h7jKnj4HL0t8PH5QJyDPtZqWN5llQ/exec"
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "Import_KSK_Khong_PL_Tuoi.xlsm")

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        # Lấy token từ query URL: /api/export?token=xyz
        parsed_path = urllib.parse.urlparse(self.path)
        query = urllib.parse.parse_qs(parsed_path.query)
        token = query.get('token', [''])[0]

        if not token:
            self.send_response(400)
            self.send_header('Content-type', 'text/plain')
            self.end_headers()
            self.wfile.write(b"Missing token")
            return

        # 1. Gọi Google Apps Script để lấy dữ liệu
        try:
            payload = json.dumps({"action": "getData", "token": token})
            res = requests.post(GAS_API_URL, data=payload, headers={'Content-Type': 'text/plain;charset=utf-8'}, allow_redirects=True)
            gas_response = res.json()
            
            if not gas_response.get("success"):
                self.send_response(403)
                self.send_header('Content-type', 'text/plain;charset=utf-8')
                self.end_headers()
                self.wfile.write(f"GAS Error: {gas_response.get('message', 'Unknown')}".encode('utf-8'))
                return
            
            data_rows = gas_response.get("data", [])
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-type', 'text/plain;charset=utf-8')
            self.end_headers()
            self.wfile.write(f"Failed to fetch data from GAS: {str(e)}".encode('utf-8'))
            return

        # 2. Xử lý dữ liệu vào file Excel Template
        try:
            wb = openpyxl.load_workbook(TEMPLATE_PATH, keep_vba=True)
            sheet = wb["Trên 18"]
            
            # Giả định: Dữ liệu trả về từ GAS
            # Bỏ qua dòng tiêu đề nếu cần (ở đây data_rows[0] thường là tiêu đề)
            current_row = 4 # Theo template chuẩn bạn đưa
            count = 1
            
            for row in data_rows[1:]: # Bỏ dòng tiêu đề của sheet Data
                if not row or not row[0]: continue # Dòng trống
                
                sheet[f'A{current_row}'] = count
                sheet[f'B{current_row}'] = row[2] # Mã CSKCB (Giả sử cột 3 trong GSheet Data)
                sheet[f'D{current_row}'] = row[3] # Ngày khám
                sheet[f'E{current_row}'] = row[4] # Họ tên
                sheet[f'F{current_row}'] = row[5] # Giới tính
                sheet[f'G{current_row}'] = row[6] # Ngày sinh
                # ... Map tiếp các trường tương tự form
                
                current_row += 1
                count += 1
            
            # 3. Trả file về cho trình duyệt dưới dạng download
            excel_bytes = io.BytesIO()
            wb.save(excel_bytes)
            excel_bytes.seek(0)
            
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.ms-excel.sheet.macroEnabled.12')
            self.send_header('Content-Disposition', 'attachment; filename="KetQua_KSK.xlsm"')
            self.end_headers()
            self.wfile.write(excel_bytes.read())

        except Exception as e:
            self.send_response(500)
            self.send_header('Content-type', 'text/plain;charset=utf-8')
            self.end_headers()
            self.wfile.write(f"Excel Processing Error: {str(e)}".encode('utf-8'))
            return
