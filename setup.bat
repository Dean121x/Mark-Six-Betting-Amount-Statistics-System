@echo off
chcp 65001 >nul
echo ==========================================
echo   六合彩金额统计系统 - 安装 / 打包
echo ==========================================
echo.
echo [1] 直接运行 (需 Python)
echo [2] 打包成独立 exe (推荐，可发给任何人用)
echo.
set /p choice="请选择 (1 或 2): "

if "%choice%"=="1" goto run
if "%choice%"=="2" goto build
goto end

:run
echo.
echo 安装依赖...
pip install openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple 2>nul || pip install openpyxl
echo 启动程序...
python main.py
goto end

:build
echo.
echo 安装依赖...
pip install openpyxl pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple 2>nul || pip install openpyxl pyinstaller
echo.
echo 开始打包 (约 1-2 分钟)...
pyinstaller --noconsole --onefile --name "六合彩金额统计系统" main.py
echo.
echo ==========================================
echo   打包完成！
echo   exe 位置: dist\六合彩金额统计系统.exe
echo   把这个 exe 发给任何人，双击即可使用
echo ==========================================

:end
pause
