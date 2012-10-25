echo off
cls
copy "\\10.0.1.232\Attendance App\DLLS\*" "%SYSTEMROOT%\SYSTEM32\" /Y
regsvr32 /s "%SYSTEMROOT%\system32\SSubTmr6.dll"
regsvr32 /s "%SYSTEMROOT%\system32\vbalIml6.ocx"
regsvr32 /s "%SYSTEMROOT%\system32\vbalGrid6.ocx"
regsvr32 /s "%SYSTEMROOT%\system32\vbalSGrid6.ocx"