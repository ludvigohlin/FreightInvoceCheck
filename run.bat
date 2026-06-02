@echo off
cd /d "%~dp0"
echo [%DATE% %TIME%] Starting >> "%~dp003_Logs\task_runner.log"
"C:\Users\LudvigOhlin\AppData\Local\FreightInvoiceControl\venv\Scripts\python.exe" main.py 2>> "%~dp003_Logs\task_runner.log"
echo [%DATE% %TIME%] Exit code: %ERRORLEVEL% >> "%~dp003_Logs\task_runner.log"
