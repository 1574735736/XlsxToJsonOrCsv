@echo off
chcp 65001
@SET EXCEL_FOLDER=.\excel
@SET JSON_FOLDER=.\json
@SET EXE=.\excel2json.exe

@ECHO Converting excel files in folder %EXCEL_FOLDER% ...
for /f "delims=" %%i in ('dir /b /a-d /s %EXCEL_FOLDER%\*.xlsx') do (
    @echo   processing %%~nxi 
    @CALL %EXE% --excel %EXCEL_FOLDER%\%%~nxi --json %JSON_FOLDER%\%%~ni.json --header 3 --exclude_prefix ! --cell_json true --array true
)

@ECHO Renaming output files to .text extension...
for /r %JSON_FOLDER% %%x in (*.json) do (
    move /Y "%%x" "%%~dx%%~px%%~nx.text"
)