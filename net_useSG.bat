REM @echo off

REM C:\ProgramData\Cisco\Cisco AnyConnect VPN Client\Script
Title FTFY

set cur_date=%date:~10,4%_%date:~4,2%_%date:~7,2% 
set cur_time=%time:~0,2%_%time:~3,2%_%time:~6,2%

echo %cur_date%_%cur_time%: S drive connection >> C:/logs/nav2Project.log

net use s: \\fs007\jobs /persistent:yes >> C:/logs/nav2Project.log

echo %cur_date%_%cur_time%: G drive connection >> C:/logs/nav2Project.log

net use g: \\fs002\data /persistent:yes  >> C:/logs/nav2Project.log


C:\Users\smyers\AppData\Local\Apps\2.0\HOCJZJO1.BTW\74EQYAQK.2EB\nav2..tion_20a030676f742022_0002.0000_cc2f3315f2618213\Nav2Project.exe
echo %cur_date%_%cur_time%: Nav2Project Launched >> C:/logs/nav2Project.log

