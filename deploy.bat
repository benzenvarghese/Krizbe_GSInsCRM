@echo off
title AMAIA CRM Deploy Utility - Nanba Edition

echo ===========================================================
echo                AMAIA CRM DEPLOY TOOL - by Nanba 
echo ===========================================================
echo.
echo 1. Pull latest code from Google Apps Script
echo 2. Push local code to GitHub and Google Apps Script
echo 3. Deploy Standalone Script using Deployment ID
echo.

set /p choice=Choose an option [1/2/3]: 

if "%choice%"=="1" (
    echo Pulling latest from Google Apps Script...
    clasp pull
    echo Pull complete!
    goto end
)

if "%choice%"=="2" (
    echo Checking Git status...
    git status
    echo.
    set /p commitMsg=Enter commit message: 
    if "%commitMsg%"=="" (
        echo ❌ Commit message cannot be empty. Aborting.
        goto end
    )
    git add .
    git commit -m "%commitMsg%"
    echo  Pushing to GitHub...
    git push origin main
    echo Pushing to Google Apps Script...
    clasp push
    echo Code pushed to both GitHub and GAS!
    goto end
)


if "%choice%"=="3" (
    set /p deployId= Enter Deployment ID: 
    echo  Deploying Standalone Script...
    clasp deploy --deploymentId %deployId%
    echo  Deployment complete!
    goto end
)

echo  Invalid option. Please run again.

:end
echo.
echo Press any key to return to the prompt...
pause >nul
