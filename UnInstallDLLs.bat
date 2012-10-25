echo off
cls
regsvr32 /s /u "%SYSTEMROOT%\system32\SSubTmr.dll"
regsvr32 /s /u "%SYSTEMROOT%\system32\SSubTmr6.dll"
regsvr32 /s /u "%SYSTEMROOT%\system32\vbalIml6.ocx"
regsvr32 /s /u "%SYSTEMROOT%\system32\vbalGrid6.ocx"
del /Q /S "%SYSTEMROOT%\system32\SSubTmr.dll"
del /Q /S "%SYSTEMROOT%\system32\SSubTmr6.dll"
del /Q /S %SYSTEMROOT%\system32\vbalIml6.ocx"
del /Q /S "%SYSTEMROOT%\system32\vbalGrid6.ocx"
@pause