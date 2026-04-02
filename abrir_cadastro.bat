@echo off
title Criacao de Usuario - AD + Microsoft 365
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\criar_usuario_gui.ps1"
