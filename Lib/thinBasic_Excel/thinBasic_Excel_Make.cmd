C:
cd \ThinBASIC_On_GitHub\ThinBASIC_On_GitHub\Lib\thinBasic_Excel
c:\pb\pbwin1000\bin\pbwin thinBasic_Excel.bas /iC:\PB\PBWin1000\Jose\WINAPI_III\;C:\PB\PBWin1000\WinAPI /l /q

:UPX
if exist thinBasic_Excel.dll C:\thinbasic\upx\upx.exe --ultra-brute thinBasic_Excel.dll


:EndOFScript
pause
