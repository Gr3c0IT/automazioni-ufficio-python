@echo off
rem Avvio pianificato del report prodotti (es. da Utilita di pianificazione di Windows).
rem Usa il Python presente nel PATH; il log viene scritto in logs\run_report.log
cd /d "%~dp0"
if not exist logs mkdir logs
python send_agent_product_report.py >> "logs\run_report.log" 2>&1
