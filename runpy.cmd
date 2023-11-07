set PY_PIP=python-3.9.0b4-embed-amd64\Scripts
set PY_LIBS=python-3.9.0b4-embed-amd64\Lib;python-3.9.0b4-embed-amd64\Lib\site-packages
cd /d %~dp0
python-3.9.0b4-embed-amd64\python.exe %1
pause