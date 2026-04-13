@echo off
REM 금형이력카드 처리 프로그램 - EXE 빌드 스크립트
REM PyInstaller를 사용하여 main.py를 단일 EXE 파일로 변환

setlocal enabledelayedexpansion

echo ============================================
echo  금형이력카드 프로그램 EXE 빌드
echo ============================================
echo.

REM 1. PyInstaller 설치 확인
echo [1/4] PyInstaller 설치 확인 중...
pip list | find "PyInstaller" >nul
if errorlevel 1 (
    echo PyInstaller를 설치하고 있습니다...
    pip install PyInstaller==6.1.0
)
echo.

REM 2. 기존 빌드 폴더 정리
echo [2/4] 기존 빌드 파일 정리 중...
if exist "dist" rmdir /s /q "dist" >nul 2>&1
if exist "build" rmdir /s /q "build" >nul 2>&1
if exist "main.spec" del /q "main.spec" >nul 2>&1
echo.

REM 3. PyInstaller로 EXE 생성
echo [3/4] EXE 파일 생성 중... (약 30~60초 소요)
pyinstaller --name "금형이력카드프로그램" ^
    --onefile ^
    --windowed ^
    --icon=app.ico ^
    --add-data "data:data" ^
    --add-data "YES:YES" ^
    --add-data "img:img" ^
    --hidden-import=openpyxl ^
    --hidden-import=olefile ^
    --hidden-import=docx ^
    --hidden-import=PIL ^
    main.py 2>nul

if errorlevel 1 (
    echo EXE 생성 실패. PyInstaller가 제대로 설치되었는지 확인하세요.
    pause
    exit /b 1
)
echo.

REM 4. 빌드 완료
echo [4/4] 빌드 완료!
echo.
echo ============================================
echo  ✓ EXE 파일 생성 완료
echo ============================================
echo.
echo 생성된 파일: dist\금형이력카드프로그램.exe
echo.
echo 다음 단계:
echo  1. dist 폴더에 생성된 EXE 파일 확인
echo  2. 필요한 데이터 폴더 복사 (YES, data, 등)
echo  3. 사용자에게 EXE 파일 배포
echo.
pause
