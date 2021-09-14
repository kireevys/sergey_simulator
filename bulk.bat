@ECHO OFF
ECHO Start
call .\venv\Scripts\activate.bat
::python -m pip install -r requirements.txt >nul
python cli.py bulk config.yml %1
pause