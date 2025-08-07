@echo off
REM 激活conda环境
call conda activate visa

REM 运行python脚本并保存日志
python n9918a_frontend.py > log.txt 2>&1

REM 或者如果你希望看到输出同时保存在文本和窗口，可以用powershell:
:: powershell -Command "python n9918a_frontend.py | Tee-Object -FilePath log.txt"

pause