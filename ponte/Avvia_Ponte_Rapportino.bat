@echo off
title Ponte Rapportino di Cantiere
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0ponte.ps1"
if errorlevel 1 pause
