@echo off
:: // By Handsome Zombie or Aston or PyAston or �ֹ۵İ�˹�� or PositiveAston (All of them are me.)
:: Get system version
for /f "usebackq skip=2 tokens=2*" %%a in (`reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName`) do set System_Version=%%b
echo %System_Version%|findstr /I "XP">nul 2>nul&&set System_is_WinXP=1&&goto Do_Not_Get_UAC||goto Get_UAC
:Get_UAC
%1 start "" mshta vbscript:createobject("shell.application").shellexecute("""%~0""","::",,"runas",1)(window.close)&exit
:Do_Not_Get_UAC
cd %systemroot%\system32
setlocal enabledelayedexpansion
set lines=����������������������
:: Check language
reg query "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Nls\Language" /v InstallLanguage|find "0804" >nul 2>nul&&set Language=Chinese||set Language=English
if /I "%Language%"=="English" (
	set message1=Now I will help you to clear the KMS on your Windows.&set wait=Please wait...&set message2=Please reboot your Windows.
	set message3=AUTO installing the key to activate your Windows&set message4=Command error.&set message5=Getting your Installation ID.
	set message6=Installation ID&set message7=Trying to get Confirmation ID&set message8=Can not find Curl.exe in your Windows.
	set message9=If you want to use full functions,please download Curl.exe and copy it to %systemroot%\system32.&set message10=Some functions is not avaliable.
	set message11=Get confirmation ID successfully!&set message12=Confirmation ID&set message13=Installing confirmation ID&set Error_Msg1=Error:
	set message14=You are not connected to the network.
	)
if /I "%Language%"=="Chinese" (
	set message1=������ʼ��� Windows KMS��&set wait=���Ժ�...&set message2=�������������ϵͳ��
	set message3=�����Զ���Կ������� Windows&set message4=��������&set message5=���ڻ�ȡ��װ ID
	set message6=��װ ID&set message7=���ڳ��Ի�ȡȷ�� ID&set message8=���ϵͳ��ȱ�� Curl.exe��
	set message9=����ʹ���������������� Curl.exe �����Ƶ� %systemroot%\system32 Ŀ¼��&set message10=���ֹ��ܽ��޷�����ʹ�á�
	set message11=�ɹ���ȡȷ�� ID!&set message12=ȷ�� ID&set message13=���ڴ���ȷ�� ID&set Error_Msg1=����
	set message14=�ƺ�û�����ӵ����硣&set message15=�������Ӹ��·�������&set message16=����ʧ�ܣ����Ե� ^!Tried_Times^! ���ӡ�
	)
set Helper_Version=1.0&set Helper_Version_Number=1

:: Check other things.
if not exist "%systemroot%\system32\curl.exe" (
	echo %Error_Msg1%
	echo %message8%
	echo %message10%
	echo %message9%
	echo %lines%
	)
if /I "%Language%"=="chinese" (
	ping www.baidu.com /n 1 >nul&&set Internet=Online||set Internet=Offline
	) else (
	ping www.google.com /n 1 >nul&&set Internet=Online||set Internet=Offline
	)
if /I "%Internet%"=="Offline" (
	echo %Error_Msg1%
	echo %message14%
	echo %message10%
	echo %lines%
	)
title Activation_Helper on %System_Version% %PROCESSOR_ARCHITECTURE% %Internet%
:loop
if exist "%tmp%\cid.txt" del /q "%tmp%\cid.txt"
set input=
set /p input=^>
if "%input%"=="" goto loop
echo %input%|findstr /I /C:auto >nul 2>nul&&goto Auto_Activate
echo %input%|findstr /I /C:Backup >nul 2>nul&&goto Backup_Activation_Information
if /I "%input%"=="CKMSWin" goto Clean_Windows_KMS
if /I "%input%"=="CKMSOff" goto Clean_Office_KMS
if /I "%input%"=="ShowLic" goto Show_License
if /I "%input%"=="cls" cls&goto loop
if /I "%input%"=="help" start https://www.10001.cf/Batch_Activation_Helper/HowToUse.html
if /I "%input%"=="myiid" goto Get_MY_IID
if /I "%input%"=="Update" goto Check_Update
echo %message4%
goto loop

:Clean_Windows_KMS
echo %lines%
echo %message1%
echo %wait%
echo.
cscript slmgr.vbs -upk
cscript slmgr.vbs -ckms
cscript slmgr.vbs -rearm
echo %lines%
echo %message2%
goto loop

:Auto_Activate
echo %input%|findstr /I /C:"Windows 7" >nul 2>nul&&set Auto_Activate_Windows_7=1||set Auto_Activate_Windows_7=0
if "%Auto_Activate_Windows_7%"=="1" goto Auto_Activate_Windows_7
echo %message4%
goto loop

:Auto_Activate_Windows_7
set No_System_Version=0
if /I "%Language%"=="Chinese" (
	echo !input!|findstr "�콢" >nul 2>nul&&set System_Version=Ultimate&&set Shown_System_Version=�콢��&&set key1=BMB34-PMB4B-3DBVK-BVYTT-JXXQQ&&set key2=89V88-T9R6V-6TTQ3-RBQVG-T4G6G||set No_System_Version=1
	echo !input!|findstr "רҵ" >nul 2>nul&&set System_Version=Professional&&set Shown_System_Version=רҵ��&&set key1=VTJDT-4M6HH-JKMR3-87BC8-4DQF3&&set key2=72MV6-22PRR-6V73G-HBP4J-WGY87||set No_System_Version=1
	echo !input!|findstr "��ͥ��ͨ" >nul 2>nul&&set System_Version=HomeBasic&&set Shown_System_Version=��ͥ��ͨ��&&set key1=BJ89T-WCC72-JQXHF-CBT43-PH4TY&&set key2=6BJC3-BWX6M-K8XPW-VHVMM-D3X3D||set No_System_Version=1
	echo !input!|findstr "��ͨ" >nul 2>nul&&set System_Version=HomeBasic&&set Shown_System_Version=��ͥ��ͨ��&&set key1=BJ89T-WCC72-JQXHF-CBT43-PH4TY&&set key2=6BJC3-BWX6M-K8XPW-VHVMM-D3X3D||set No_System_Version=1
	echo !input!|findstr "��ͥ�߼�" >nul 2>nul&&set System_Version=HomePrimium&&set Shown_System_Version=��ͥ�߼���&&set key1=87H4R-M9XYV-HRFM6-WC3D8-Q4PP9&&set key2=HR8PM-RGKCG-YYPHP-RTPJR-HWXWV||set No_System_Version=1
	echo !input!|findstr "�߼�" >nul 2>nul&&set System_Version=HomePrimium&&set Shown_System_Version=��ͥ�߼���&&set key1=87H4R-M9XYV-HRFM6-WC3D8-Q4PP9&&set key2=HR8PM-RGKCG-YYPHP-RTPJR-HWXWV||set No_System_Version=1
	echo !input!|findstr /I "Ultimate" >nul 2>nul&&set System_Version=Ultimate&&set Shown_System_Version=�콢��&&set key1=BMB34-PMB4B-3DBVK-BVYTT-JXXQQ&&set key2=89V88-T9R6V-6TTQ3-RBQVG-T4G6G||set No_System_Version=1
	echo !input!|findstr /I "Professional" >nul 2>nul&&set System_Version=Professional&&set Shown_System_Version=רҵ��&&set key1=VTJDT-4M6HH-JKMR3-87BC8-4DQF3&&set key2=72MV6-22PRR-6V73G-HBP4J-WGY87||set No_System_Version=1
	echo !input!|findstr /I "Pro" >nul 2>nul&&set System_Version=Professional&&set Shown_System_Version=רҵ��&&set key1=VTJDT-4M6HH-JKMR3-87BC8-4DQF3&&set key2=72MV6-22PRR-6V73G-HBP4J-WGY87||set No_System_Version=1
	echo !input!|findstr /I "HomeBasic" >nul 2>nul&&set System_Version=HomeBasic&&set Shown_System_Version=��ͥ��ͨ��&&set key1=BJ89T-WCC72-JQXHF-CBT43-PH4TY&&set key2=6BJC3-BWX6M-K8XPW-VHVMM-D3X3D||set No_System_Version=1
	echo !input!|findstr /I "Basic" >nul 2>nul&&set System_Version=HomeBasic&&set Shown_System_Version=��ͥ��ͨ��&&set key1=BJ89T-WCC72-JQXHF-CBT43-PH4TY&&set key2=6BJC3-BWX6M-K8XPW-VHVMM-D3X3D||set No_System_Version=1
	echo !input!|findstr /I "HomePrimium" >nul 2>nul&&set System_Version=HomePrimium&&set Shown_System_Version=��ͥ�߼���&&set key1=87H4R-M9XYV-HRFM6-WC3D8-Q4PP9&&set key2=HR8PM-RGKCG-YYPHP-RTPJR-HWXWV||set No_System_Version=1
	echo !input!|findstr /I "Primium" >nul 2>nul&&set System_Version=HomePrimium&&set Shown_System_Version=��ͥ�߼���&&set key1=87H4R-M9XYV-HRFM6-WC3D8-Q4PP9&&set key2=HR8PM-RGKCG-YYPHP-RTPJR-HWXWV||set No_System_Version=1
	) else (
	echo !input!|findstr /I "Ultimate" >nul 2>nul&&set System_Version=Ultimate&&set Shown_System_Version=Ultimate&&set key1=BMB34-PMB4B-3DBVK-BVYTT-JXXQQ&&set key2=89V88-T9R6V-6TTQ3-RBQVG-T4G6G||set No_System_Version=1
	echo !input!|findstr /I "Professional" >nul 2>nul&&set System_Version=Professional&&set Shown_System_Version=Professional&&set key1=VTJDT-4M6HH-JKMR3-87BC8-4DQF3&&set key2=72MV6-22PRR-6V73G-HBP4J-WGY87||set No_System_Version=1
	echo !input!|findstr /I "Pro" >nul 2>nul&&set System_Version=Professional&&set Shown_System_Version=Professional&&set key1=VTJDT-4M6HH-JKMR3-87BC8-4DQF3&&set key2=72MV6-22PRR-6V73G-HBP4J-WGY87||set No_System_Version=1
	echo !input!|findstr /I "HomeBasic" >nul 2>nul&&set System_Version=HomeBasic&&set Shown_System_Version=HomeBasic&&set key1=BJ89T-WCC72-JQXHF-CBT43-PH4TY&&set key2=6BJC3-BWX6M-K8XPW-VHVMM-D3X3D||set No_System_Version=1
	echo !input!|findstr /I "Basic" >nul 2>nul&&set System_Version=HomeBasic&&set Shown_System_Version=HomeBasic&&set key1=BJ89T-WCC72-JQXHF-CBT43-PH4TY&&set key2=6BJC3-BWX6M-K8XPW-VHVMM-D3X3D||set No_System_Version=1
	echo !input!|findstr /I "HomePrimium" >nul 2>nul&&set System_Version=HomePrimium&&set Shown_System_Version=HomePrimium&&set key1=87H4R-M9XYV-HRFM6-WC3D8-Q4PP9&&set key2=HR8PM-RGKCG-YYPHP-RTPJR-HWXWV||set No_System_Version=1
	echo !input!|findstr /I "Primium" >nul 2>nul&&set System_Version=HomePrimium&&set Shown_System_Version=HomePrimium&&set key1=87H4R-M9XYV-HRFM6-WC3D8-Q4PP9&&set key2=HR8PM-RGKCG-YYPHP-RTPJR-HWXWV||set No_System_Version=1
	echo !input!|findstr "�콢" >nul 2>nul&&set System_Version=Ultimate&&set Shown_System_Version=Ultimate&&set key1=BMB34-PMB4B-3DBVK-BVYTT-JXXQQ&&set key2=89V88-T9R6V-6TTQ3-RBQVG-T4G6G||set No_System_Version=1
	echo !input!|findstr "רҵ" >nul 2>nul&&set System_Version=Professional&&set Shown_System_Version=Professional&&set key1=VTJDT-4M6HH-JKMR3-87BC8-4DQF3&&set key2=72MV6-22PRR-6V73G-HBP4J-WGY87||set No_System_Version=1
	echo !input!|findstr "��ͥ��ͨ" >nul 2>nul&&set System_Version=HomeBasic&&set Shown_System_Version=HomeBasic&&set key1=BJ89T-WCC72-JQXHF-CBT43-PH4TY&&set key2=6BJC3-BWX6M-K8XPW-VHVMM-D3X3D||set No_System_Version=1
	echo !input!|findstr "��ͨ" >nul 2>nul&&set System_Version=HomeBasic&&set Shown_System_Version=HomeBasic&&set key1=BJ89T-WCC72-JQXHF-CBT43-PH4TY&&set key2=6BJC3-BWX6M-K8XPW-VHVMM-D3X3D||set No_System_Version=1
	echo !input!|findstr "��ͥ�߼�" >nul 2>nul&&set System_Version=HomePrimium&&set Shown_System_Version=HomePrimium&&set key1=87H4R-M9XYV-HRFM6-WC3D8-Q4PP9&&set key2=HR8PM-RGKCG-YYPHP-RTPJR-HWXWV||set No_System_Version=1
	echo !input!|findstr "�߼�" >nul 2>nul&&set System_Version=HomePrimium&&set Shown_System_Version=HomePrimium&&set key1=87H4R-M9XYV-HRFM6-WC3D8-Q4PP9&&set key2=HR8PM-RGKCG-YYPHP-RTPJR-HWXWV||set No_System_Version=1
	)
	
:: // ������ʱ���߼������Ȱ�����ע�͵���No_System_Version ����������ڻ�û�ã��������Ҫ��errorlevel���ж���û�а汾�ˡ�
::if "%No_System_Version%"=="1" echo %message4%&&goto loop
set KeyNum=1
:Auto_Activate_Windows_7_Start
echo %lines%
echo %message3% 7 %Shown_System_Version%,%wait%
cscript slmgr.vbs -ipk !key%KeyNum%!>%tmp%\Activation_Message.txt
if /I "%Language%"=="chinese" (
	findstr "�ɹ�" "%tmp%\Activation_Message.txt" >nul 2>nul&&set Install_Key=success||set Install_Key=fail
	if /I "!Install_Key!"=="success" (
		echo ��Կ��װ�ɹ�!
		) else (
		echo ��Կ��װʧ�ܣ�������!
		goto loop
		)
	)
if /I "%Language%"=="english" (
	findstr /I "success" "%tmp%\Activation_Message.txt"&&set Install_Key=success||set Install_Key=fail
	if /I "!Install_Key!"=="success" (
		echo The key is installed successfully!
		) else (
		echo Failed to install the key!
		echo Please try it again!
		goto loop
		)
	)
	echo %lines%
echo %message5% ,%wait%
for /f "tokens=3" %%i in ('cscript slmgr.vbs /dti') do set InstallationID=%%i
echo %message6%^>!InstallationID!
echo %lines%
echo %message7%,%wait%

:Auto_Activate_Windows_7_Get_CID
curl -s -k -c "%tmp%\Aihao.cookie" -H "Content-Type: application/json" -H "Accept: application/json" "http://webact.185.hk" >nul
curl.exe -k -s -o "%tmp%\cid.txt" "http://webact.185.hk/getCID.php?IID=%InstallationID%" -H "Referer: http://webact.185.hk/" -b "%tmp%\Aihao.cookie"
if not exist "%tmp%\cid.txt" goto Auto_Activate_Windows_7_Get_CID
for /f tokens^=8^ delims^=^" %%i in ('findstr "CID" "%tmp%\cid.txt"') do set ConfirmationID=%%i
echo %message12%^>%ConfirmationID%
echo %lines%
echo %message13%,%wait%
set "ConfirmationID=%ConfirmationID:-=%"
cscript slmgr.vbs -atp %ConfirmationID%>"%tmp%\Activation_Message.txt"
if /I "%Language%"=="chinese" (
	findstr "�ɹ�" "%tmp%\Activation_Message.txt" >nul 2>nul&&set Install_CID=success||set Install_CID=fail
	if "%KeyNum%"=="1" (
		if /I "!Install_CID!"=="success" (
			echo ����ɹ���
			) else (
			echo ����ʧ�ܣ����ڳ��Ը�����Կ���� Windows 7 %Shown_System_Version%...
			set KeyNum=2
			goto Auto_Activate_Windows_7_Start
			)
		)
	if "%KeyNum%"=="2" (
		echo ���޿�����Կ��������Ѱ����Կ�ٳ��Լ���
		goto loop
		)
	)
if /I "%Language%"=="english" (
	findstr /I "success" "%tmp%\Activation_Message.txt" >nul 2>nul&&set Install_CID=success||set Install_CID=fail
	if "%KeyNum%"=="1" (
		if /I "!Install_CID!"=="success" (
			echo Your Windows is successfully activated��
			) else (
			echo Activation failed.Try to use another key to activate your Windows 7 %Shown_System_Version%...
			set KeyNum=2
			goto Auto_Activate_Windows_7_Start
			)
		)
	if "%KeyNum%"=="2" (
		echo There is no more key to activate your Windows 7 %Shown_System_Version%.
		echo Please find the key yourself and try to activate your Windows.
		)
	)
start slmgr.vbs -xpr
echo %lines%
goto loop

:Check_Update
echo %message15%,%wait%
set Tried_Times=1
:Get_Update_File
curl -k -s -o "%tmp%\Activation_Helper_Update.txt" --connect-timeout 5 "https://www.10001.cf/Batch_Activation_Helper/program/update.txt"
if not exist "%tmp%\Activation_Helper_Update.txt" (
	set /a Tried_Times+=1
	echo %message16%
	goto Get_Update_File
	)

for /f "usebackq tokens=2 delims==" %%a in (`findstr /I "Helper_Version_Number=" "%tmp%\Activation_Helper_Update.txt"`) do set New_Helper_Version_Number=%%a
for /f "usebackq tokens=2 delims==" %%a in (`findstr /I "Helper_Version=" "%tmp%\Activation_Helper_Update.txt"`) do set New_Helper_Version=%%a
for /f "usebackq tokens=2 delims==" %%a in (`findstr /I "Download_Link=" "%tmp%\Activation_Helper_Update.txt"`) do set Download_Link=%%a
if %New_Helper_Version_Number% GTR %Helper_Version_Number% (
	echo ��ǰ�汾��%Helper_Version%
	echo ���°汾��%New_Helper_Version%
	goto loop
	) else (
	echo �޸���
	goto loop
	)