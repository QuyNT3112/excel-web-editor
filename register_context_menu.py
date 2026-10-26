import os
import sys
import winreg

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

def register_association():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    exe_path = os.path.join(script_dir, "dist", "ExcelWebPro.exe")
    
    if not os.path.isfile(exe_path):
        print(f"[Error] File not found: {exe_path}")
        return False
        
    print(f"Registering Excel Web Pro for: {exe_path}")
    exts = [".xlsx", ".xls", ".xlsm", ".csv"]
    
    for ext in exts:
        base_key_path = f"Software\\Classes\\SystemFileAssociations\\{ext}\\Shell\\OpenWithExcelWebPro"
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, base_key_path) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Mở bằng Excel Web Pro")
                winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, f'"{exe_path}"')
                
            cmd_key_path = f"{base_key_path}\\command"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, cmd_key_path) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, f'"{exe_path}" "%1"')
                
            print(f" [OK] Đã đăng ký menu chuột phải cho định dạng {ext}")
        except Exception as e:
            print(f" [Lỗi] Không thể đăng ký cho {ext}: {e}")
            return False

    print("\n[THÀNH CÔNG] Đã tích hợp 'Mở bằng Excel Web Pro' vào menu chuột phải Windows!")
    print("Bạn có thể click chuột phải vào bất kỳ file .xlsx, .xls, .xlsm, .csv và chọn 'Mở bằng Excel Web Pro'.")
    return True

if __name__ == "__main__":
    success = register_association()
    if not success:
        sys.exit(1)
