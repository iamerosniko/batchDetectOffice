@echo off
setlocal enabledelayedexpansion

set Office[7]=Office 97
set Office[8]=Office 98
set Office[9]=Office 2000
set Office[10]=Office XP
set Office[11]=Office 2003
set Office[12]=Office 2007
set Office[14]=Office 2010
set Office[15]=Office 2013
set Office[16]=Office 2016
set CurVer=
for /f "tokens=3 delims=." %%a in ('reg.exe query "HKCR\Word.Application\CurVer" /ve ^| find.exe /i "(Default)"') do set CurVer=%%a
if defined CurVer (
	set OfficeVersion=!Office[%CurVer%]!
	echo Found Office: Version %CurVer%, '!OfficeVersion!'
	@echo Found Office: Version %CurVer%, '!OfficeVersion!'> test.txt
) else (
	echo No Office installation found.
)