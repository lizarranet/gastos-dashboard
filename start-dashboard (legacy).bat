@echo off
start "Dashboard Backend" cmd /k "cd /d C:\Proyectos\gastos-dashboard\backend && npm start"
start "Dashboard Frontend LAN" cmd /k "cd /d C:\Proyectos\gastos-dashboard\frontend && npm run dev -- --host 0.0.0.0"