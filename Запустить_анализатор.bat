@echo off
chcp 65001 >nul
echo Запуск Анализатора документов...
echo.
cd /d "%~dp0"
python document_analyzer_improved.py
pause 