@echo off
chcp 65001
@SET EXCEL_FOLDER=%1
@SET JSON_FOLDER=%2
@SET EXE_PATH=%3

@ECHO Converting excel files in folder %EXCEL_FOLDER% ...
for /f "delims=" %%i in ('dir /b /a-d /s %EXCEL_FOLDER%\*.xlsx') do (
    @echo   processing %%~nxi 
    @CALL "%EXE_PATH%" --excel "%%i" --json "%JSON_FOLDER%\%%~ni.json" --header 3 --exclude_prefix ! --cell_json true --array true
)
