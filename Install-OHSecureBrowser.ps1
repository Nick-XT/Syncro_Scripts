function Install-OHSecureBrowser {
    param(
        [String]$Uninstall
    )
    $Path = $env:TEMP
    $Installer = 'ohsecurebrowser_installer.msi'

    if ($Uninstall -eq 'True') {
        $UninstallGUID = (Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-Object { $_.DisplayName -eq "OHSecureBrowser" }).PSChildName
        Write-Host "Uninstalling OHSecureBrowser..."
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $UninstallGUID /quiet /norestart" -Wait
        Write-Host "OHSecureBrowser uninstalled successfully."
    } else {
        Write-Host "Installing OHSecureBrowser..."
        Invoke-WebRequest -uri 'https://sb.portal.cambiumast.com/geturls?clientName=ohio&operatingSystem=windows' -OutFile $Path\$Installer
        Start-Process -FilePath $Path\$Installer -Args '/quiet /norestart' -Wait
        Write-Host "OHSecureBrowser installed successfully."
    }
    If (Test-Path $Path\$Installer) {
        Remove-Item -Path $Path\$Installer
    }
}

Install-OHSecureBrowser -Uninstall:$Syncro_Uninstall