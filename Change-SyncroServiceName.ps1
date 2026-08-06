$ServiceNames = @("Syncro", "SyncroLive")

foreach ($s in $ServiceNames) {
    try {
        if (Get-Service -Name $s -ErrorAction SilentlyContinue) {
            Set-Service -Name $s -Description "Monitoring Agent provided by Miller Network Innovations"
            Write-Host "Service $s description updated successfully."
        }
    } catch {
        Write-Error -Message "Service $s not found. Skipping description update."
    }
}