@echo off
title KAP Risk Izleme Platformu
echo KAP Risk Izleme Platformu baslatiliyor... Tarayici otomatik acilacak.
echo Kapatmak icin bu pencerede Ctrl+C ya da pencereyi kapatin.
cd /d "%~dp0"
python -m streamlit run kap_risk_app.py
pause
