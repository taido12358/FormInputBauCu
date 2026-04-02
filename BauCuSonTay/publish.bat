@echo off
echo ============================================
echo   BUILD + PUBLISH BauCuSonTay.exe
echo ============================================

cd /d "%~dp0BauCuSonTay"

echo.
echo [1/2] Restoring packages...
dotnet restore

echo.
echo [2/2] Publishing single-file self-contained...
dotnet publish -c Release -r win-x64 --self-contained true ^
  -p:PublishSingleFile=true ^
  -p:IncludeNativeLibrariesForSelfExtract=true ^
  -p:EnableCompressionInSingleFile=true ^
  -o "..\publish_output"

echo.
echo ============================================
if exist "..\publish_output\BauCuSonTay.exe" (
    echo   THANH CONG!
    echo   File EXE: publish_output\BauCuSonTay.exe
) else (
    echo   That bai - kiem tra loi o tren
)
echo ============================================
pause
