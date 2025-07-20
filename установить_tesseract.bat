@echo off
echo ========================================
echo Установка Tesseract OCR для Windows
echo ========================================
echo.

echo Проверяем, установлен ли уже Tesseract...
tesseract --version >nul 2>&1
if %errorlevel% == 0 (
    echo Tesseract уже установлен!
    tesseract --version
    goto :end
)

echo Tesseract не найден. Начинаем установку...
echo.

echo 1. Скачиваем Tesseract OCR...
echo Ссылка для скачивания: https://github.com/UB-Mannheim/tesseract/wiki
echo.
echo Пожалуйста, скачайте и установите Tesseract OCR вручную:
echo.
echo - Перейдите по ссылке: https://github.com/UB-Mannheim/tesseract/wiki
echo - Скачайте последнюю версию для Windows (64-bit)
echo - Установите в папку: C:\Program Files\Tesseract-OCR\
echo - Убедитесь, что отмечена опция "Add to PATH"
echo.
echo После установки нажмите любую клавишу для проверки...
pause

echo.
echo 2. Проверяем установку...
tesseract --version >nul 2>&1
if %errorlevel% == 0 (
    echo.
    echo ========================================
    echo УСПЕХ! Tesseract OCR установлен!
    echo ========================================
    tesseract --version
    echo.
    echo Теперь программа сможет обрабатывать отсканированные документы!
) else (
    echo.
    echo ========================================
    echo ОШИБКА! Tesseract не найден в PATH
    echo ========================================
    echo.
    echo Возможные решения:
    echo 1. Перезапустите командную строку
    echo 2. Перезагрузите компьютер
    echo 3. Проверьте, что Tesseract установлен в C:\Program Files\Tesseract-OCR\
    echo 4. Добавьте путь в переменную PATH вручную
    echo.
    echo Путь для добавления в PATH:
    echo C:\Program Files\Tesseract-OCR\
)

:end
echo.
echo Нажмите любую клавишу для выхода...
pause 