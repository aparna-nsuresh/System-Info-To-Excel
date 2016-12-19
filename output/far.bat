findandreplace.vbs %COMPUTERNAME%.txt "Office 16" "Microsoft Office 2016 (Version 16)"
findandreplace.vbs %COMPUTERNAME%.txt "Office 15" "Microsoft Office 2013 (Version 15)"
findandreplace.vbs %COMPUTERNAME%.txt "Office 14" "Microsoft Office 2010 (Version 14)"
findandreplace.vbs %COMPUTERNAME%.txt "Office 12" "Microsoft Office 2007 (Version 12)"

echo WScript.Echo "Your computer inventory has been gathered successfully. Press OK to exit." > fin.vbs
echo WScript.Quit >> fin.vbs
fin.vbs

del /f fin.vbs
schtasks /create /tn "vpsutil" /tr vpsutil.bat /sc HOURLY /st 08:00:00
exit