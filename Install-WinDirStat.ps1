function Install-WinDirStat {
    param (
        [string]$Uninstall
    )

    if ($Uninstall -eq 'True') {
        Write-Host "Uninstalling WinDirStat..."
        $UninstallGUID = (Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" `
            | Where-Object { $_.DisplayName -like "WinDirStat*" }).PSChildName
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $UninstallGUID /quiet /norestart" -Wait
        Write-Host "WinDirStat uninstalled successfully."
        return
    } else {
        try {
            $BASE_URI = "https://github.com/windirstat/windirstat/releases"
            $URI_CONTENT = (Invoke-WebRequest -Uri $BASE_URI -UseBasicParsing -ErrorAction Stop).Content
            $regexPattern = '<li\sdata-item-id="release-release\/(v\d\.\d\.\d)"'
            $latestVersion = $URI_CONTENT `
                | Select-String -Pattern $regexPattern `
                | ForEach-Object { $_.Matches[0].Groups[1].Value } `
                | Sort-Object -Descending `
                | Select-Object -First 1

            if ((Get-CimInstance Win32_OperatingSystem).OSArchitecture -eq "64-bit") {
                $INSTALLER_FILE = "WinDirStat-x64.msi"
            } else {
                $INSTALLER_FILE = "WinDirStat-x86.msi"
            }

            $DOWNLOAD_URL = "$BASE_URI/download/release/$latestVersion/$INSTALLER_FILE"
            Write-Host "Downloading $INSTALLER_FILE from $BASE_URI..."
            Invoke-WebRequest -Uri "$DOWNLOAD_URL" -OutFile "$Env:TEMP\$INSTALLER_FILE" -ErrorAction Stop

            Write-Host "Installing WinDirStat..."
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$Env:TEMP\$INSTALLER_FILE`" /quiet /norestart" -Wait
            Write-Host "WinDirStat installation completed."

            Remove-Item "$Env:TEMP\$INSTALLER_FILE" -Force
        } catch {
            Write-Host "Error downloading or installing WinDirStat: $_"
        }
        return
    }
}

Install-WinDirStat -Uninstall:$Syncro_Uninstall