function Install-GoogleChrome {
    param(
        [String]$Uninstall
    )
    $Path = $env:TEMP
    $Installer = 'chrome_installer.exe'

    if ($Uninstall -eq 'False') {
        try {
        Invoke-WebRequest -Uri 'https://dl.google.com/chrome/install/latest/chrome_installer.exe' -OutFile $Path\$Installer
        Start-Process -FilePath $Path\$Installer -Args '/silent /install' -Wait
        Remove-Item -Path $Path\$Installer
        Write-Host 'Google Chrome uninstalled successfully.'
        } catch {
            Write-Error "Google Chrome installation failed.`nError: $_"
        }
    } else {
        try {
            $regPaths = @(
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            )
            $chromeReg = $regPaths | ForEach-Object {
                Get-ItemProperty "$_\*" | Where-Object { $_.DisplayName -match 'Google Chrome' }
            }
            Write-Host "Uninstalling Google Chrome..."
            Start-Process -FilePath $chromeReg.UninstallString.Split('"')[1] -ArgumentList '--uninstall --system-level --force-uninstall' -Wait
            Write-Host "Google Chrome uninstalled successfully."
        } catch {
            Write-Error "Google Chrome could not be uninstalled.`n(Error: $_)"
        }
    }
}

Install-GoogleChrome -Uninstall:$Syncro_Uninstall