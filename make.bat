recycle-bin.exe build
recycle-bin.exe dist
pyinstaller --noconfirm --onefile --console read_csv.py
move dist\read_csv.exe read_csv.exe
pause