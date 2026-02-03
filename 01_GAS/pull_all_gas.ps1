$rootPath = $PSScriptRoot
Get-ChildItem -Path $rootPath -Directory | ForEach-Object {
    $targetDir = $_.FullName
    if (Test-Path "$targetDir\.clasp.json") {
        Write-Host "--------------------------------------------------"
        Write-Host "Processing: $($_.Name)"
        Write-Host "Path: $targetDir"
        
        Push-Location $targetDir
        try {
            # Execute clasp pull
            clasp pull
        } catch {
            Write-Error "Failed to pull in $targetDir"
        }
        Pop-Location
        Write-Host "--------------------------------------------------`n"
    } else {
        Write-Host "Skipping $($_.Name) (no .clasp.json found)`n"
    }
}

Write-Host "All operations completed."
Read-Host -Prompt "Press Enter to exit"
