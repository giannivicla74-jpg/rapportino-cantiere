@echo off
title Scegli cartelle - Ponte Rapportino
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0ponte.ps1" -Configura
if errorlevel 1 pause
