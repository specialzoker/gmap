@echo off
setlocal enabledelayedexpansion
chcp 65001 > nul
cd /d "%~dp0"

echo ===== G-map 2028 데이터 업데이트 =====
echo.

set "DEFAULT_XLSX=2028 작업 틀_(2차 완료).xlsx"

if "%~1"=="" (
    echo 엑셀 파일명을 입력하세요.
    echo 그냥 Enter를 누르면 기본 파일을 사용합니다:
    echo   %DEFAULT_XLSX%
    echo.
    set /p "XLSX=파일명 (Enter=기본값): "
    if "!XLSX!"=="" set "XLSX=%DEFAULT_XLSX%"
) else (
    set "XLSX=%~1"
)

echo.
echo [1/3] 엑셀 → JSON 변환 중... (%XLSX%)
python convert_to_json.py "%XLSX%"
if errorlevel 1 (
    echo.
    echo 오류 발생! Python 또는 파일을 확인하세요.
    pause
    exit /b 1
)

echo.
echo [2/3] GitHub에 업로드 중...
git add data/
git add index.html
git commit -m "데이터 업데이트 (%date%)"
git push

echo.
echo [3/3] 완료!
echo 1~2분 후 아래 주소에서 확인하세요:
echo   https://specialzoker.github.io/gmap/
echo.
pause
