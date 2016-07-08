echo %username% > output/%COMPUTERNAME%.txt
echo %computername% >> output/%COMPUTERNAME%.txt
systeminfo | findstr /B /C:"OS Name" >> output/%COMPUTERNAME%.txt
wmic product get name /format:list | more | findstr . >> output/%COMPUTERNAME%.txt
systeminfo | find "Original Install Date" >> output/%COMPUTERNAME%.txt

netsh interface ip show addresses "Wi-Fi" | find "IP" >> output/%COMPUTERNAME%.txt
wmic nic where "macaddress <> null" get macaddress, name | find "Wi-Fi" >> output/%COMPUTERNAME%.txt
wmic cpu get name /Format:list | more | findstr . >> output/%COMPUTERNAME%.txt
fsutil volume diskfree C: >> output/%COMPUTERNAME%.txt






