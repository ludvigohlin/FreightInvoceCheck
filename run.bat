@echo off
cd /d "%~dp0"
echo [%DATE% %TIME%] Starting >> "%~dp003_Logs\task_runner.log"
echo [%DATE% %TIME%] Dir: %CD% >> "%~dp003_Logs\task_runner.log"
echo [%DATE% %TIME%] Python: "%~dp0.venv\Scripts\python.exe" >> "%~dp003_Logs\task_runner.log"
"%~dp0.venv\Scripts\python.exe" main.py 2>> "%~dp003_Logs\task_runner.log"
echo [%DATE% %TIME%] Exit code: %ERRORLEVEL% >> "%~dp003_Logs\task_runner.log"
