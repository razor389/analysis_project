@echo on
echo Setting up the project environment...

:: Check if Python is installed
where python >nul 2>nul
if errorlevel 1 (
    echo Error: Python is not installed or not added to PATH. Please install Python 3.7 or newer and try again.
    pause
    exit /b 1
)

:: Create a virtual environment
echo Creating a virtual environment...
python -m venv venv
if errorlevel 1 (
    echo Error: Failed to create virtual environment. Ensure Python is properly installed.
    pause
    exit /b 1
)

:: Activate the virtual environment
echo Activating the virtual environment...
call venv\Scripts\activate
if errorlevel 1 (
    echo Error: Failed to activate virtual environment.
    pause
    exit /b 1
)

:: Check if requirements.txt exists and is not empty
if not exist requirements.txt (
    echo Warning: requirements.txt not found. Skipping dependency installation.
    pause
    exit /b 1
)
for /f %%A in ('find /c /v "" ^< requirements.txt') do set count=%%A
if %count%==0 (
    echo Warning: requirements.txt is empty. Skipping dependency installation.
    pause
    exit /b 1
)

:: Upgrade pip and install dependencies
echo Installing dependencies...
pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
    echo Error: Failed to install dependencies. Check the requirements.txt file and your internet connection.
    pause
    exit /b 1
)

:: Copy .env.example to .env if it doesn't exist
if not exist .env (
    if exist .env.example (
        copy /Y .env.example .env
        if errorlevel 1 (
            echo Error: Failed to copy .env.example to .env.
            pause
            exit /b 1
        )
        echo .env file created. Please edit it with your OpenAI API key.
    ) else (
        echo Error: .env.example does not exist. Please create a .env file manually.
        pause
        exit /b 1
    )
) else (
    echo .env file already exists.
)

echo Setup complete. To run the script, activate the virtual environment:
echo call venv\Scripts\activate
echo Then run: python acm_analysis.py <ticker> <start_year>
pause
