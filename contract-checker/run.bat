@echo off
echo ============================================================
echo  Правовая экспертиза договора
echo ============================================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ОШИБКА] Python не найден. Установите Python 3.9+
    pause
    exit /b 1
)

:: Check API key
if "%ANTHROPIC_API_KEY%"=="" (
    echo [ВНИМАНИЕ] Переменная ANTHROPIC_API_KEY не установлена.
    set /p ANTHROPIC_API_KEY="Введите ваш Anthropic API ключ: "
    set ANTHROPIC_API_KEY=%ANTHROPIC_API_KEY%
)

:: Install dependencies
echo Установка зависимостей...
pip install -r requirements.txt -q

echo.
echo  Сервер запущен: http://localhost:5000
echo  Нажмите Ctrl+C для остановки
echo.
python app.py
pause
