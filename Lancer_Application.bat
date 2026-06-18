@echo off
title Relevés BTCC / DCBR
cd /d "%~dp0"

echo Démarrage de l'application...

REM Vérifier si Python est installé
python --version >nul 2>&1
if errorlevel 1 (
    echo ERREUR : Python n'est pas installé ou introuvable.
    echo Veuillez installer Python depuis https://www.python.org
    pause
    exit /b 1
)

REM Installer les dépendances si nécessaire
pip install -r requirements.txt --quiet

REM Ouvrir le navigateur après 2 secondes
start "" /B cmd /c "timeout /t 2 >nul && start http://localhost:5000"

REM Lancer l'application Flask
python app.py
