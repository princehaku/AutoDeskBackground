@echo off
echo "删除程序中...."
@taskkill /f /im core.exe
@echo y|rmdir /s .\
echo "删除成功"
pause