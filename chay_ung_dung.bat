@echo off
cd /d "%~dp0"

python -c "import sys" >nul 2>nul
if errorlevel 1 (
    echo ============================================================
    echo   LOI: May nay chua co Python hoat dong dung cach.
    echo ============================================================
    echo.
    echo   Neu ban vua thay dong chu "Python was not found... install
    echo   from the Microsoft Store" thi day LA NGUYEN NHAN: Windows
    echo   co san mot "python" GIA (App execution alias) - go lenh
    echo   python se chi hien thong bao do chu khong chay Python that,
    echo   KE CA khi may ban da cai Python that roi.
    echo.
    echo   CACH SUA - lam dung theo thu tu:
    echo   1. Mo Settings ^> Apps ^> Advanced app settings ^>
    echo      App execution aliases
    echo   2. TAT (gat OFF) 2 muc "App Installer python.exe" va
    echo      "App Installer python3.exe"
    echo   3. Neu may CHUA cai Python that, cai tai:
    echo      https://www.python.org/downloads/
    echo      Man hinh dau tien luc cai, PHAI TICK CHON o vuong
    echo      "Add python.exe to PATH" TRUOC KHI bam Install Now.
    echo   4. Dong het cac cua so cmd dang mo (hoac khoi dong lai may
    echo      cho chac), roi chay lai file nay.
    echo ============================================================
    pause
    exit /b 1
)

echo Dang cai dat thu vien can thiet (chi mat thoi gian o lan chay dau)...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ============================================================
    echo   LOI: Cai dat thu vien that bai. Kiem tra ket noi mang roi thu lai.
    echo ============================================================
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   DUNG DONG CUA SO NAY khi dang dung ung dung!
echo   Dong cua so nay = tat ung dung ngay lap tuc.
echo ============================================================
python app.py

echo.
echo Ung dung da dung. Nhan phim bat ky de dong cua so nay.
pause
