@echo off
color ff
mode con: cols=100 lines=36
title Windows 10 优化小工具

set TempFile_Name=%SystemRoot%\System32\BatTestUACin_SysRt%Random%.batemp
( echo "BAT Test UAC in Temp" >%TempFile_Name% ) 1>nul 2>nul
if exist %TempFile_Name% (
del %TempFile_Name% 1>nul 2>nul
goto menu
) else (
goto failed
)

:menu
cls
echo.
echo.
echo                   请选择需要执行的操作
echo.
echo                     1.系统设置
echo.
echo                     2.Appx管理
echo.
echo                     3.系统激活
echo.
echo                     4.office激活
echo.
echo                     E.退出
echo.
echo                                          By. YX_life  整理制作
echo.     
echo.
set /p bg= 请输入选择的对应的编号:
if %bg%==1 goto settings
if %bg%==2 goto appx
if %bg%==3 goto activation
if %bg%==4 goto office
if %bg%==E exit
goto error

:settings
cls
echo.
title 系统设置
echo   =====================================================================
echo                        Windows 10 系统设置                      
echo   =====================================================================
echo.
echo  【1】去除桌面小箭头、后缀和盾牌图标
echo  【2】隐藏导航栏OneDrive
echo  【3】隐藏导航栏可移动设备（U盘）
echo  【R】返回                            【E】退出
echo.
set /p set=请输入需要设置对应的编号:
if %set%==1 goto set1
if %set%==2 goto set2
if %set%==3 goto set3
if /i %set%==R goto menu
if /i %set%==E exit
goto error

:set1
cls
echo.
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons" /v 29 /d "%systemroot%\system32\imageres.dll,197" /t reg_sz /f
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons" /v 77 /d "%systemroot%\system32\imageres.dll,197" /t reg_sz /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer" /v link /d "00000000" /t REG_BINARY /f
taskkill /f /im explorer.exe
start explorer
echo.
pause
goto settings

:set2
cls
echo.
reg add "HKEY_CLASSES_ROOT\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\ShellFolder" /v Attributes /t REG_DWORD /d 0xf090004d /f
reg add "HKEY_CLASSES_ROOT\WOW6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}\ShellFolder" /v Attributes /t REG_DWORD /d 0xf090004d /f
taskkill /f /im explorer.exe
start explorer
echo.
pause
goto settings

:set3
cls
echo.
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\DelegateFolders\{F5FB2C77-0E2F-4A16-A381-3E560C68BC83}" /f
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\DelegateFolders\{F5FB2C77-0E2F-4A16-A381-3E560C68BC83}" /f
taskkill /f /im explorer.exe
start explorer
echo.
pause
goto settings

:appx
cls
echo.
title Appx管理
echo   =====================================================================
echo                        Windows 10 Appx管理                     
echo   =====================================================================
echo.
echo   PS:一键卸载：人脉、OneNote、Sticky Notes、3D Builder、Skype预览版
echo                语音录音机、闹钟和时钟、消息
echo.
echo  【1】一键卸载
echo.
echo   以下选用：
echo.
echo  【A】Groove音乐                      【B】电影与电视
echo  【C】计算器                          【D】地图
echo.
echo  【F】重装所有内置应用
echo.
echo  【R】返回                            【E】退出
echo.
set /p app= 请输入选择的对应的编号:
if /i %app%==show goto show
if /i %app%==del goto del
if %app%==1 goto app1
if /i %app%==A goto appA
if /i %app%==B goto appB
if /i %app%==C goto appC
if /i %app%==D goto appD
if /i %app%==F goto appF
if /i %app%==R goto menu
if /i %app%==E exit
goto error

:app1
cls
echo.
powershell "Get-AppxPackage *people* | Remove-AppxPackage"
powershell "Get-AppxPackage *OneNote* | Remove-AppxPackage"
powershell "Get-AppxPackage *Sticky* | Remove-AppxPackage"
powershell "Get-AppxPackage *sky* | Remove-AppxPackage"
powershell "Get-AppxPackage *Alarm* | Remove-AppxPackage"
powershell "Get-AppxPackage *soundrec* | Remove-AppxPackage"
powershell "Get-AppxPackage *solit* | Remove-AppxPackage"
powershell "Get-AppxPackage *messaging* | Remove-AppxPackage"
echo 执行完成，按任意键返上一级菜单
pause >nul
goto appx

:appA
cls
echo.
powershell "Get-AppxPackage *ZuneMusic* | Remove-AppxPackage"
echo 执行完成，按任意键返上一级菜单
pause >nul
goto appx

:appB
cls
echo.
powershell "Get-AppxPackage *ZuneVideo* | Remove-AppxPackage"
echo 执行完成，按任意键返上一级菜单
pause >nul
goto appx

:appC
cls
echo.
powershell "Get-AppxPackage *calc* | Remove-AppxPackage"
echo 执行完成，按任意键返上一级菜单
pause >nul
goto appx

:appD
cls
echo.
powershell "Get-AppxPackage *Maps* | Remove-AppxPackage"
echo 执行完成，按任意键返上一级菜单
pause >nul
goto appx

:appF
cls
echo.
powershell "Get-AppxPackage -AllUsers| Foreach {Add-AppxPackage -DisableDevelopmentMode -Register “$($_.InstallLocation)\AppXManifest.xml”}"
echo 执行完成，按任意键返上一级菜单
pause >nul
goto appx

:activation
cls
echo.
title WIN10专业版激活
echo   =====================================================================
echo               Windows 10 专业版 一键自动激活(需联网)                
echo   =====================================================================
echo.
echo  【1】查看系统激活状态                【2】一键永久激活(提示确定即可)
echo.
echo  【R】返回                            【E】退出
echo.
set /p atc= 请输入选择的对应的编号:
if %atc%==1 goto status
if %atc%==2 goto activate
if /i %atc%==R goto menu
if /i %atc%==E exit
goto error

:status
cls
slmgr.vbs -dlv
slmgr.vbs -dli
slmgr.vbs -xpr
echo 执行完成，按任意键返上一级菜单
pause>nul
goto activation

:activate
slmgr /upk
slmgr -ipk D9NK9-8CT4J-94KXY-H93TQ-JTYP6
slmgr -ato
echo 执行完成，按任意键返上一级菜单
pause>nul
goto activation

:office
cls
echo.
title Office激活
echo   =====================================================================
echo           Office Professional Plus 2016（专业版Plus）
echo   =====================================================================
echo.
echo  【1】激活                            【2】查看激活状态
echo  【R】返回                            【E】退出
echo.
set /p off= 请输入选择的对应的编号:
if %off%==1 goto office1
if %off%==2 goto office2
if /i %off%==R goto menu
if /i %off%==E exit
goto error

:office1
cls
setlocal EnableDelayedExpansion&color 3e & cd /d "%~dp0"

::修改下面的内容，定义选择想使用的KMS服务器。如果定义了多次，最后的有效

::set KMS_Sev=192.168.2.8
::set KMS_Sev=1.2.7.0
set KMS_Sev=kms.lotro.cc
::set KMS_Sev=54.223.212.31
::set KMS_Sev=kms.guowaifuli.com
::set KMS_Sev=mhd.kmdns.net
::set KMS_Sev=xykz.f3322.org
::set KMS_Sev=106.186.25.239
::set KMS_Sev=110.noip.me
::set KMS_Sev=3rss.vicp.net:20439
::set KMS_Sev=45.78.3.223
::set KMS_Sev=kms.chinancce.com
::set KMS_Sev=kms.didichuxing.com
::set KMS_Sev=skms.ddns.net
::set KMS_Sev=zh.us.to
::set KMS_Sev=franklv.ddns.net
::set KMS_Sev=k.zpale.com
::set KMS_Sev=m.zpale.com
::set KMS_Sev=mvg.zpale.com
::set KMS_Sev=122.226.152.230
::set KMS_Sev=222.76.251.188
::set KMS_Sev=annychen.pw
::set KMS_Sev=heu168.6655.la
::set KMS_Sev=kms.aglc.cc
::set KMS_Sev=kms.landiannews.com
::set KMS_Sev=kms.shuax.com
::set KMS_Sev=kms.xspace.in
::set KMS_Sev=winkms.tk
::set KMS_Sev=wrlong.com

if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

echo 正在安装 KMS 许可证...
for /f %%x in ('dir /b ..\root\Licenses16\project???vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\standardvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\visio???vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

echo 正在安装 MAK 许可证...
for /f %%x in ('dir /b ..\root\Licenses16\project???vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\standardvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\visio???vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul

echo 正在导入 KMS GVLK...
cscript ospp.vbs /inpkey:NYH39-6GMXT-T39D4-WVXY2-D69YY >nul
cscript ospp.vbs /inpkey:7WHWN-4T7MP-G96JF-G33KR-W8GF4 >nul
cscript ospp.vbs /inpkey:RBWW7-NTJD4-FFK2C-TDJ7V-4C2QP >nul
cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul
cscript ospp.vbs /inpkey:YG9NW-3K39V-2T3HJ-93F3Q-G83KT >nul
cscript ospp.vbs /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK >nul

echo 正在尝试 KMS 激活...
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul
cscript //nologo ospp.vbs /act | find /i "successful" && (
        echo.&echo ***** Office2016 激活成功 ***** & echo.) || (echo.&echo ***** Office2016 激活失败 ***** & echo.
        echo 请检查网络是否畅通，以及选择修改其他的 KMS 服务器后再试)
echo 执行完成，按任意键返上一级菜单
pause >nul
goto office

:office2
cls
echo.
cscript.exe "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /dstatus
echo 执行完成，按任意键返上一级菜单
pause>nul
goto office

:error
cls
echo.
title 输入错误
echo   =====================================================================
echo                        Windows 10 优化小工具                               
echo   =====================================================================
echo.
echo           输入错误，请重新输入。。。。。 
echo.
echo 按任意键返回主菜单
pause >nul
goto menu

:failed
cls
color 0a
mode con: cols=50 lines=15
title 操作失败
echo.
echo.
echo    操作失败。
echo    请右键“以管理员身份运行”
echo.
echo    按任意键退出...
echo.
echo.
pause >nul
exit