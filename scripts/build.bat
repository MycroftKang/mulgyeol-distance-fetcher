rmdir /s /q build\app build\update
copy src\version_info.txt src\version_info_temp.txt
python package\vc.py -lb
cd src
pyinstaller main.spec
pyinstaller msu.spec
xcopy /I /Y /E dist\msu\* dist\main
move dist\main ..\build\app
rmdir /s /q dist\msu
xcopy /I /Y /E ..\package\data ..\build\data
xcopy /I /Y ..\package\info ..\build\info
copy ..\product.json ..\build
move version_info_temp.txt version_info.txt