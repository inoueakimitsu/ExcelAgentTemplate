@echo off
REM Specify the path to the Python interpreter of the virtual environment. Users should rewrite this path as needed.
REM Run `where.exe python` with the virtual environment activated to get the path to the Python interpreter of the virtual environment.
set VENV_PATH=C:\path\to\your\venv\Scripts\python.exe

REM Run langchain_fastapi.py.
%VENV% langchain_fastapi.py

REM It does not end until a key is pressed.
pause
