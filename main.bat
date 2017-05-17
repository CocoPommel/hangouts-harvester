@echo off
setlocal enableextensions
cd /d %~dp0
Call :UnZipFile "%~dp0" "%~dp0hangouts-export.zip"
break > output.log
set "name=John Smith"
for %%f in (*.csv) do findstr /r /c:"%name%" /c:"^ORA-[0-9]*" "%%f" >> output.log
set "format=hangouts-conversation-"
for /r %%a in (*.csv) do for /f "delims=" %%i in ('echo("%%~na" ^| findstr /i "%format%"') do del "%%~fa"
exit /b

:UnZipFile <ExtractTo> <newzipfile>
set vbs="%temp%\_.vbs"
if exist %vbs% del /f /q %vbs%
>%vbs%  echo Set fso = CreateObject("Scripting.FileSystemObject")
>>%vbs% echo If NOT fso.FolderExists(%1) Then
>>%vbs% echo fso.CreateFolder(%1)
>>%vbs% echo End If
>>%vbs% echo set objShell = CreateObject("Shell.Application")
>>%vbs% echo set FilesInZip=objShell.NameSpace(%2).items
>>%vbs% echo objShell.NameSpace(%1).CopyHere(FilesInZip)
>>%vbs% echo Set fso = Nothing
>>%vbs% echo Set objShell = Nothing
cscript //nologo %vbs%
if exist %vbs% del /f /q %vbs%
