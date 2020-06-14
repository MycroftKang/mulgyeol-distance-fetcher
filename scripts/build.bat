rmdir /s /q build\app build\update
cd src
pyinstaller main.spec
pyinstaller msu.spec
move dist\main ../build/app
move dist\msu ../build/update