@echo off
mode con cols=100 lines=30 & color 1f

>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
rem --> If error flag set, we do not have admin.  
if '%errorlevel%' NEQ '0' (  
    echo 获取管理员权限中,如果UAC弹窗,请选择允许...
    goto UACPrompt  
) else ( goto gotAdmin )  
:UACPrompt  
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"  
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"   
    "%temp%\getadmin.vbs"
    exit /B  
:gotAdmin  
    if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )  
    pushd "%CD%"  
    CD /D "%~dp0"

cls
:start
echo\
echo	请输入数字选择(请以管理员模式运行):
echo\
echo	1：激活 Windows 10 专业版    2：激活 Windows 8 专业版     3：激活 Windows 7 专业版    
echo	4：激活 Office 2016    5：查看 Windows 激活状态  
set /P var=":"
if %var%==1 goto 10
if %var%==2 goto 8
if %var%==3 goto 7
if %var%==4 goto office
if %var%==5 goto look

:10
cls
echo\
echo	正在安装产品密钥，如有弹窗请按[确定]继续...
echo\
slmgr -ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
echo\
echo	正在设置KMS服务器地址，如有弹窗请按[确定]继续...
echo\
slmgr -skms kms.garybear.tk
echo\
echo	正在自动激活...
echo\
slmgr -ato
goto exit


:office
cls
cd "C:\Program Files (x86)\Microsoft Office\Office16"
echo\
echo	正在安装产品密钥，如有弹窗请按[确定]继续...
echo\
cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 
echo\
echo	正在设置KMS服务器地址，如有弹窗请按[确定]继续...
echo\
cscript ospp.vbs /sethst:kms.garybear.tk
echo\
echo	正在自动激活...
echo\
cscript ospp.vbs /act
goto exit


:8
cls
echo\
echo	正在安装产品密钥，如有弹窗请按[确定]继续...
echo\
slmgr -ipk NG4HW-VH26C-733KW-K6F98-J8CK4
echo\
echo	正在设置KMS服务器地址，如有弹窗请按[确定]继续...
echo\
slmgr -skms kms.garybear.tk
echo\
echo	正在自动激活...
echo\
slmgr -ato
goto exit


:7
cls
echo\
echo	正在安装产品密钥，如有弹窗请按[确定]继续...
echo\
slmgr -ipk FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4
echo\
echo	正在设置KMS服务器地址，如有弹窗请按[确定]继续...
echo\
slmgr -skms kms.garybear.tk
echo\
echo	正在自动激活...
echo\
slmgr -ato
goto exit



:look
cls
echo\
echo	请注意查看弹窗内容
slmgr.vbs -dlv
goto exit


:exit
echo\
echo\
echo	按任意键退出
pause>nul 
start http://blog.garybear.cn/ & exit