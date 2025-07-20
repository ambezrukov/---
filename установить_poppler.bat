@echo off
echo ========================================
echo Установка Poppler для Windows
echo ========================================
echo.

echo Проверяем, установлен ли уже Poppler...
pdftoppm -h >nul 2>&1
if %errorlevel% == 0 (
    echo Poppler уже установлен!
    pdftoppm -h
    goto :end
)

echo Poppler не найден. Начинаем установку...
echo.

echo 1. Скачиваем Poppler...
echo Ссылка для скачивания: https://github.com/oschwartz10612/poppler-windows/releases
echo.
echo Пожалуйста, скачайте и установите Poppler вручную:
echo.
echo - Перейдите по ссылке: https://github.com/oschwartz10612/poppler-windows/releases
echo - Скачайте последнюю версию (Release-xxx.zip)
echo - Распакуйте в папку: C:\poppler\
echo - Добавьте путь C:\poppler\bin в переменную PATH
echo.
echo После установки нажмите любую клавишу для проверки...
pause

echo.
echo 2. Проверяем установку...
pdftoppm -h >nul 2>&1
if %errorlevel% == 0 (
    echo.
    echo ========================================
    echo УСПЕХ! Poppler установлен!
    echo ========================================
    pdftoppm -h
    echo.
    echo Теперь программа сможет конвертировать PDF в изображения для OCR!
) else (
    echo.
    echo ========================================
    echo ОШИБКА! Poppler не найден в PATH
    echo ========================================
    echo.
    echo Возможные решения:
    echo 1. Перезапустите командную строку
    echo 2. Перезагрузите компьютер
    echo 3. Проверьте, что Poppler распакован в C:\poppler\
    echo 4. Добавьте путь в переменную PATH вручную
    echo.
    echo Путь для добавления в PATH:
    echo C:\poppler\bin
)

:end
echo.
echo Нажмите любую клавишу для выхода...
pause 