@echo off
if not exist "dist\" make dist > nul 2>&1
pyinstaller -F main.py
copy 9224.wav dist\
copy bg_1920x1080.jpg dist\
copy gc_cz.png dist\
copy name_file.xls dist\
copy simhei.ttf dist\