@echo off
cd /d "%~dp0"

pyinstaller ^
  --onedir ^
  --windowed ^
  --name GrokImageTool ^
  --collect-all tkinterdnd2 ^
  main.py

pause
