@echo off  
if "%processor_architecture%"=="x86" goto REG32  
if "%processor_architecture%"=="AMD64" goto REG64  
:REG32  
if not exist %systemroot%\system32\COMDLG32.OCX COPY COMDLG32.OCX  %systemroot%\system32   
%systemroot%\system32\regsvr32.exe -u -s %systemroot%\system32\COMDLG32.OCX
%systemroot%\system32\regsvr32.exe -s %systemroot%\system32\COMDLG32.OCX  
if not exist %systemroot%\system32\MSCHRT20.OCX COPY MSCHRT20.OCX  %systemroot%\system32   
%systemroot%\system32\regsvr32.exe -u -s %systemroot%\system32\MSCHRT20.OCX
%systemroot%\system32\regsvr32.exe -s %systemroot%\system32\MSCHRT20.OCX
if not exist %systemroot%\system32\msscript.ocx COPY msscript.ocx  %systemroot%\system32   
%systemroot%\system32\regsvr32.exe -u -s %systemroot%\system32\msscript.ocx
%systemroot%\system32\regsvr32.exe -s %systemroot%\system32\msscript.ocx
if not exist %systemroot%\system32\MSWINSCK.OCX COPY MSWINSCK.OCX  %systemroot%\system32   
%systemroot%\system32\regsvr32.exe -u -s %systemroot%\system32\MSWINSCK.OCX
%systemroot%\system32\regsvr32.exe -s %systemroot%\system32\MSWINSCK.OCX
goto exit  
:REG64  
if not exist %systemroot%\syswow64\COMDLG32.OCX COPY COMDLG32.OCX  %systemroot%\syswow64   
%systemroot%\syswow64\regsvr32.exe -u -s %systemroot%\system32\COMDLG32.OCX  
%systemroot%\syswow64\regsvr32.exe -s %systemroot%\system32\COMDLG32.OCX  
if not exist %systemroot%\system32\MSCHRT20.OCX COPY MSCHRT20.OCX  %systemroot%\system32   
%systemroot%\syswow64\regsvr32.exe -u -s %systemroot%\system32\MSCHRT20.OCX
%systemroot%\syswow64\regsvr32.exe -s %systemroot%\system32\MSCHRT20.OCX
if not exist %systemroot%\system32\msscript.ocx COPY msscript.ocx  %systemroot%\system32   
%systemroot%\syswow64\regsvr32.exe -u -s %systemroot%\system32\msscript.ocx
%systemroot%\syswow64\regsvr32.exe -s %systemroot%\system32\msscript.ocx
if not exist %systemroot%\system32\MSWINSCK.OCX COPY MSWINSCK.OCX  %systemroot%\system32   
%systemroot%\syswow64\regsvr32.exe -u -s %systemroot%\system32\MSWINSCK.OCX
%systemroot%\syswow64\regsvr32.exe -s %systemroot%\system32\MSWINSCK.OCX
goto exit  
:exit  
copy ieframe.dll %windir%\system32\
copy wmp.dll %windir%\system32\
regsvr32 %windir%\system32\ieframe.dll /s
regsvr32 %windir%\system32\wmp.dll /s
exit

