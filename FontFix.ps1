if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
        $Command = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
        Start-Process -FilePath PowerShell.exe -Verb RunAs -ArgumentList $Command
        Exit
 }
}

$WordFontUrl = “https://github.com/hasanzadekhalil/FontChangeOffice/raw/main/Normal.dotm”
$OutlookFontFix = “https://raw.githubusercontent.com/hasanzadekhalil/FontChangeOffice/main/Outlook.reg”

$WordFontPath=”C:\Temp\Normal.dotm”
$OutlookFontPath="C:\Temp\Outlook.reg"

Invoke-WebRequest -URI $WordFontUrl -OutFile $WordFontPath
Invoke-WebRequest -URI $OutlookFontFix -OutFile $OutlookFontPath

regedit.exe /s $OutlookFontPath

Copy-Item $WordFontPath -Destination "$env:APPDATA\Microsoft\Templates"

