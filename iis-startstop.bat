:: Created: 2016-05-03
:: Version: REV1.0
:: Author: Sam Myers

:: Lists IIS sites and allows stop and start



@ECHO OFF
Title -----------IIS Simple Manager-----------
color 72

goto check_Permissions

:check_permissions
    echo Administrative permissions required. Detecting permissions...
	echo          -

    net session >nul 2>&1
    if %errorLevel% == 0 (
        echo Success: Administrative permissions confirmed.
    ) else (
        echo Failure: Current permissions inadequate. Please run as Administrator.
	echo Press any key to exit.
		 pause >nul
		 exit
    )
echo          -

cd %SystemRoot%
cd System32
cd inetsrv

:list_sites

echo Here are the stati of your hosted sites.
echo ******************************************************************************
appcmd list sites
echo ******************************************************************************
echo Type the name of the site you wish to change exactly as it appears above.
echo          -
SET /P website=Site Name:
echo          ----------------


:start_or_stop
	echo You typed: "%website%"
    SET /P startStop= Start (1) or Stop (0):
   
    if %startStop% == 1 (
		appcmd start sites "%website%"
        echo          -
    ) else (
	    appcmd stop sites "%website%"
		echo          -
    )

goto list_sites
