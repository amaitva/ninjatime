$Path = "C:\Program Files (x86)\internalinfrastructuremainoffice-5.3.5097\"

cd $path

sleep -seconds 5

Start-Process -FilePath .\uninstall.exe -ArgumentList "--mode","unattended"

sleep -seconds 40

$Uninstall2 = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | Where-Object {$_.DisplayName -match "NinjaRMMAgent" } | Select-Object -Property UninstallString | foreach { $_.UninstallString }

sleep -seconds 220

$Uninstall3 = $Uninstall2 -Replace "MsiExec.exe " , ""

sleep -seconds 5

Start-Process -FilePath MSIExec.exe -ArgumentList $Uninstall3,"/quiet","/passive"
