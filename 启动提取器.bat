@echo off
set HTTP_PROXY=
set HTTPS_PROXY=
set http_proxy=
set https_proxy=
set ALL_PROXY=
set all_proxy=

echo =====================================
echo Setting up local environment...
echo =====================================
if not exist "venv" (
    python -m venv venv
)

echo =====================================
echo Installing requirements to local folder...
echo =====================================
venv\Scripts\pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

echo.
echo =====================================
echo Starting Dance Music Extractor...
echo =====================================
echo                _
echo              (o )^>
echo              /   \
echo             ^|  \  ^|
echo             ^|   \ ^|
echo             \_____/
echo              ^^^^  ^^^^
echo       MADE BYYYYYYY PENGUIN !?
echo.
venv\Scripts\python app.py