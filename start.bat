@echo off
chcp 65001 >nul
echo ติดตั้ง dependencies...
pip install -r requirements.txt
echo.
echo เริ่มต้นเซิร์ฟเวอร์ที่ http://localhost:5000
python app.py
pause
