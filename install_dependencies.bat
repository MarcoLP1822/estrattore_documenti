@echo off
setlocal

echo ================================================
echo   DOCUMENT PROCESSOR - SETUP DIPENDENZE
echo ================================================
echo.

REM Verifica se Python è installato
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python non trovato! Installa Python prima di continuare.
    echo    Scarica da: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo ✅ Python trovato
echo.

REM Verifica se pip è disponibile
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ pip non trovato! Verifica l'installazione di Python.
    pause
    exit /b 1
)

echo ✅ pip trovato
echo.

echo 📦 Installando le dipendenze Python...
echo.

REM Aggiorna pip
echo Aggiornamento pip...
python -m pip install --upgrade pip

echo.
echo Installazione dipendenze dal file requirements.txt...
pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo.
    echo ❌ Errore durante l'installazione delle dipendenze!
    echo Verifica la connessione internet e riprova.
    pause
    exit /b 1
)

echo.
echo ================================================
echo ✅ INSTALLAZIONE COMPLETATA CON SUCCESSO!
echo ================================================
echo.
echo 🚀 Ora puoi eseguire il script con:
echo    python main.py
echo.
echo 💡 Il script convertirà automaticamente:
echo    • File .doc e .docx in PDF
echo    • File .odt in PDF  
echo    • Copierà tutti i file PDF
echo    • Comprimerà PDF oltre 40MB
echo.
echo 📁 Cartella di destinazione:
echo    Desktop\files (creata automaticamente)
echo.
pause