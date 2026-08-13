function SideloadDrivers-WindowsInstall {
    param(
        [string]$MountPath,
        [string]$DriverPath = "C:\DRIVERS",
        [switch]$Recurse
    )
        
    @("boot", "install") | ForEach-Object {
        $WimPath = Join-Path -Path $MountPath -ChildPath "sources\$_.wim"
        if (-not (Test-Path -Path $WimPath)) {
            Write-Error "$_.wim not found in $MountPath. Exiting."
            return
        } else {
            Write-Host "Sideloading drivers into $WimPath..."
            $ImageType = $_
            $WimData = @()
            switch -Regex (DISM.exe /Get-WimInfo /WimFile:$WimPath) {
                '^Index : (.+)$' {
                    $currentImage = @{ Index = $Matches.1 }
                }
                '^Name : (.+)$' {
                    $currentImage.Name = $Matches.1
                    $WimData += [PSCustomObject]$currentImage
                }
            }
            try {
                $WimData | ForEach-Object {
                    Write-Host "Processing $ImageType.wim - Index $($_.Index) | $($_.Name)..."
                    $MountPath = "C:\WimTemp\$ImageType$($_.Index)"
                    New-Item -Path $MountPath -ItemType Directory -Force | Out-Null

                    DISM.exe /Mount-Image /ImageFile:$WimPath /Index:$($_.Index) /MountDir:$MountPath
                    if ($LASTEXITCODE -ne 0) {
                        throw "Error mounting $ImageType.wim index $($_.Index): $LASTEXITCODE. Exiting."
                    }
                    if ($Recurse) {
                        DISM.exe /Image:$MountPath /Add-Driver /Driver:$DriverPath /Recurse
                    } else {
                        DISM.exe /Image:$MountPath /Add-Driver /Driver:$DriverPath
                    }
                    if ($LASTEXITCODE -ne 0) {
                        throw "Error adding drivers to $ImageType.wim index $($_.Index): $LASTEXITCODE. Exiting."
                    }
                    DISM.exe /Unmount-Image /MountDir:$MountPath /Commit
                    if ($LASTEXITCODE -ne 0) {
                        throw "Error unmounting $ImageType.wim index $($_.Index): $LASTEXITCODE. Exiting."
                    }
                    Write-Host "Drivers have been successfully sideloaded into the Windows ISO."
                }
            } catch {
                Write-Error "An error occurred while sideloading drivers into $ImageType.wim: $_"
                Write-Host "Cleaning up mounted images..."
                DISM /Get-MountedImageInfo | Select-String '^Mount Dir : (.+)$' | ForEach-Object { DISM /Unmount-Image /MountDir:$($_.Matches.Groups[1].Value) /Discard }
            }
        }
    }
    
    DISM.exe /Cleanup-Wim
    Remove-Item -Path $MountPath -Recurse -Force
}