@echo off
chcp 65001 >nul
echo ========================================
echo    Анализатор документов v2.0.0
echo    Автоматическая установка
echo ========================================
echo.

echo 📦 Проверка Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python не найден!
    echo 📥 Скачайте Python с https://python.org
    echo ✅ Убедитесь, что отмечена опция "Add to PATH"
    pause
    exit /b 1
)
echo ✅ Python найден

echo.
echo 📦 Установка зависимостей Python...
pip install -r requirements.txt
if errorlevel 1 (
    echo ❌ Ошибка установки зависимостей
    pause
    exit /b 1
)
echo ✅ Зависимости установлены

echo.
echo 🔧 Проверка Tesseract OCR...
tesseract --version >nul 2>&1
if errorlevel 1 (
    echo ⚠️  Tesseract не найден
    echo 📥 Запускаю установщик Tesseract...
    call установить_tesseract.bat
) else (
    echo ✅ Tesseract найден
)

echo.
echo 🔧 Проверка Poppler...
pdfinfo -v >nul 2>&1
if errorlevel 1 (
    echo ⚠️  Poppler не найден
    echo 📥 Запускаю установщик Poppler...
    call установить_poppler.bat
) else (
    echo ✅ Poppler найден
)

echo.
echo 🎉 Установка завершена!
echo 🚀 Запуск программы...
echo.
python document_analyzer_improved.py

pause 