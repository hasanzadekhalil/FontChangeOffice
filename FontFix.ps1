$WordFontUrl = “https://github.com/hasanzadekhalil/FontChangeOffice/raw/main/Normal.dotm”
$OutlookFontFix = “https://raw.githubusercontent.com/hasanzadekhalil/FontChangeOffice/main/Outlook.reg”

$WordFontPath=”C:\Temp\Normal.dotm”
$OutlookFontPath="C:\Temp\Outlook.reg"

Invoke-WebRequest -URI $WordFontUrl -OutFile $WordFontPath
Invoke-WebRequest -URI $OutlookFontFix -OutFile $OutlookFontPath

regedit.exe /s $OutlookFontPath

Copy-Item $WordFontPath -Destination "$env:APPDATA\Microsoft\Templates"