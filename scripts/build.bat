rmdir /s /q build\app build\update
copy src\version_info.txt src\version_info_temp.txt
python package\vc.py -lb
cd src
pyinstaller main.spec
pyinstaller msu.spec
move dist\main ../build/app
move dist\msu ../build/update
move version_info_temp.txt version_info.txt