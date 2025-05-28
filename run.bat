@echo off
:: =============================================================================
::                      PST File Search Tool - Batch Launcher
:: =============================================================================
::
:: Параметры можно задать через переменные ниже
:: Обязательные параметры: PST_FILE и OUTPUT_DIR
::
:: Пример использования:
::   set PST_FILE="C:\path\to\file.pst"
::   set OUTPUT_DIR="C:\output\folder"
::   search_pst.bat

setlocal enableDelayedExpansion
chcp 65001 > nul

:: -----------------------------------------------------------------------------
:: ОБЯЗАТЕЛЬНЫЕ ПАРАМЕТРЫ (должны быть заданы)
:: -----------------------------------------------------------------------------
@REM Vault
set PST_FILE="Y:\Pst\Sergey.Evsikov.Vault.pst"
set OUTPUT_DIR="R:\Reports\11_Sergey.Evsikov.июль_2024_с_6_по_8"

:: -----------------------------------------------------------------------------
:: НЕОБЯЗАТЕЛЬНЫЕ ПАРАМЕТРЫ ПОИСКА
:: -----------------------------------------------------------------------------
:: Фильтр по отправителю
:: set SENDER=""

:: Фильтр по получателю
:: set RECIPIENT=""

:: Фильтр по теме письма
:: set SUBJECT=""

:: Фильтр по тексту письма
:: set BODY=""

:: Письма, отправленные после указанной даты (формат: YYYY-MM-DD или YYYY-MM-DD HH:MM:SS)
set SENT_AFTER="01.07.2024"

:: Письма, отправленные до указанной даты
set SENT_BEFORE="30.07.2024"

:: Письма, полученные после указанной даты
set RECEIVED_AFTER="01.07.2024"

:: Письма, полученные до указанной даты
set RECEIVED_BEFORE="30.07.2024"

:: Диапазон часов отправки (формат: HH-HH, например 8-17 или 22-6)
@REM set SENT_TIME="17-20"
set SENT_TIME="6-8"

:: Диапазон часов получения
@REM set RECEIVED_TIME="17-20"
set RECEIVED_TIME="6-8"

:: -----------------------------------------------------------------------------
:: ЗАПУСК ПРОГРАММЫ
:: -----------------------------------------------------------------------------
echo Запуск PST File Search Tool с параметрами:
echo PST файл:       %PST_FILE%
echo Выходной каталог: %OUTPUT_DIR%

if "%PST_FILE%"=="" (
    echo ОШИБКА: Не указан PST файл!
    goto :error
)

if "%OUTPUT_DIR%"=="" (
    echo ОШИБКА: Не указан выходной каталог!
    goto :error
)

:: Собираем команду
set "PYTHON_CMD=python main.py %PST_FILE% --output-dir %OUTPUT_DIR%"

:: Добавляем только непустые параметры
if defined SENDER set "PYTHON_CMD=!PYTHON_CMD! --sender !SENDER!"
if defined RECIPIENT set "PYTHON_CMD=!PYTHON_CMD! --recipient !RECIPIENT!"
if defined SUBJECT set "PYTHON_CMD=!PYTHON_CMD! --subject !SUBJECT!"
if defined BODY set "PYTHON_CMD=!PYTHON_CMD! --body !BODY!"
if defined SENT_AFTER set "PYTHON_CMD=!PYTHON_CMD! --sent-after !SENT_AFTER!"
if defined SENT_BEFORE set "PYTHON_CMD=!PYTHON_CMD! --sent-before !SENT_BEFORE!"
if defined RECEIVED_AFTER set "PYTHON_CMD=!PYTHON_CMD! --received-after !RECEIVED_AFTER!"
if defined RECEIVED_BEFORE set "PYTHON_CMD=!PYTHON_CMD! --received-before !RECEIVED_BEFORE!"
if defined SENT_TIME set "PYTHON_CMD=!PYTHON_CMD! --sent-time !SENT_TIME!"
if defined RECEIVED_TIME set "PYTHON_CMD=!PYTHON_CMD! --received-time !RECEIVED_TIME!"

:: Выполняем команду
echo Запускаем команду:
echo !PYTHON_CMD!
!PYTHON_CMD!

goto :eof

:error
echo.
echo Для корректной работы необходимо задать как минимум:
echo   set PST_FILE="путь_к_pst_файлу"
echo   set OUTPUT_DIR="путь_к_выходному_каталогу"
echo.
pause
endlocal
