@ECHO OFF
ECHO Start
call .\venv\Scripts\activate.bat
::python -m pip install -r requirements.txt
python cli.py close_orders closure config.yml
pause