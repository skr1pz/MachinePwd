::Enable LAN/WLAN switching for UHS by Mark Mayer on 03-09-2015
::This script looks if patch has been applied by looking for a timestamped log file and will exit if it has been applied.  If has not been applied, it will run the vbs to enable LAN/WLAN switching
@echo off
if exist "%systemdrive%\Batchlogs\WLAN.txt" goto EOF
if not exist "%systemdrive%\BatchLogs" md "%systemdrive%\Batchlogs"

set now=%date%-%time%
echo %now% >> "%systemdrive%\Batchlogs\WLAN.txt"

%systemroot%\system32\cscript.exe \\mysqlserv\Scripts\bios-change.vbs

:EOF
exit