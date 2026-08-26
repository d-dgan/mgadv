@echo off
REM Converte para PDF todos os documentos .docx desta pasta, usando o Word
REM instalado no computador. Nada e enviado para a internet.
REM
REM Basta dar dois cliques neste arquivo.
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0converter-pdf.ps1" %*
