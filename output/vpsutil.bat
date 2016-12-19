echo %username% > output/%COMPUTERNAME%.txt
echo %computername% >> output/%COMPUTERNAME%.txt
systeminfo | findstr /B /C:"OS Name" >> output/%COMPUTERNAME%.txt
For /F "tokens=1,2 delims==" %%a in ('wmic product get name /format:list') do @if "%%b" neq "" echo %%a %%b >> output/%COMPUTERNAME%.txt
systeminfo | find "Original Install Date" >> output/%COMPUTERNAME%.txt

netsh interface ip show addresses "Wi-Fi" | find "IP" >> output/%COMPUTERNAME%.txt
wmic nic where "macaddress <> null" get macaddress, name | find "Wi-Fi" >> output/%COMPUTERNAME%.txt

For /F "tokens=1,2 delims==" %%a in ('wmic cpu get name /Format:list') do @if "%%b" neq "" echo %%b >> output/%COMPUTERNAME%.txt
fsutil volume diskfree C: >> output/%COMPUTERNAME%.txt

cd output

findandreplace.vbs %COMPUTERNAME%.txt "Office 16" "Microsoft Office 2016 (Version 16)"
findandreplace.vbs %COMPUTERNAME%.txt "Office 15" "Microsoft Office 2013 (Version 15)"
findandreplace.vbs %COMPUTERNAME%.txt "Office 14" "Microsoft Office 2010 (Version 14)"
findandreplace.vbs %COMPUTERNAME%.txt "Office 12" "Microsoft Office 2007 (Version 12)"
exit





