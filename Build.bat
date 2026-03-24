rmdir /s /q build
rmdir /s /q dist
del main.spec

python -m PyInstaller --onefile --noconsole --add-data "Logo.png;." main.py

pause
