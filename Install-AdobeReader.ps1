function Install-AdobeReader {
    param(
        [String]$Uninstall
    )

    [System.Net.ServicePointManager]::MaxServicePointIdleTime = 900000 # 15 minutes in milliseconds

    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    $result = Invoke-RestMethod -Uri "https://rdc.adobe.io/reader/products?lang=mui&site=enterprise&os=Windows%2010&country=US&nativeOs=Windows%2010&api_key=dc-get-adobereader-cdn" -WebSession $session -Headers @{ "Accept" = "*/*"; "x-api-key" = "dc-get-adobereader-cdn" }

    $version = $result.products.reader[0].version
    $versionNoDots = $version.replace('.', '')

    if ((Get-CimInstance Win32_OperatingSystem).OSArchitecture -eq "64-bit") {
        $URI = "https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/$versionNoDots/AcroRdrDCx64${versionNoDots}_MUI.exe"
    } else {
        $URI = "https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/$versionNoDots/AcroRdrDCx86${versionNoDots}_MUI.exe"
    }
    $OutFile = Join-Path $env:TEMP "AcroRdrDCx64${versionNoDots}_MUI.exe"

    if ($Uninstall -eq 'True') {
        Write-Host "Uninstalling Adobe Reader..."
        $UninstallGUID = (Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-Object { $_.DisplayName -like "Adobe Acrobat*" }).PSChildName
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $UninstallGUID /quiet /norestart" -Wait
        Write-Host "Adobe Reader uninstalled successfully."
    } else {
        try {
            Write-Host "Downloading version $version..."
            Invoke-WebRequest -Uri $URI -OutFile $OutFile -ErrorAction Stop
            Write-Host "Installing..."
            Start-Process -FilePath $OutFile -ArgumentList "/sAll /rs /rps /msi /norestart /quiet EULA_ACCEPT=YES" -Wait
            Remove-Item $OutFile -Force
        } catch {
            Write-Host "Error downloading Adobe Reader: $_ - Falling back to local installer."
            $OutFile = Resolve-Path ".\AcroRdrDCx64*.exe"
            Write-Host "Installing..."
            Start-Process -FilePath $OutFile -ArgumentList "/sAll /rs /rps /msi /norestart /quiet EULA_ACCEPT=YES" -Wait
        }
        Write-Host "Installation complete."
    }
}

Install-AdobeReader -Uninstall:$Syncro_Uninstall