# Start the dist executable so it creates its _MEI extraction folder.
$proc = Start-Process -FilePath .\dist\app.exe -PassThru
Start-Sleep -Seconds 2

$temp = $env:TEMP
$dir = Get-ChildItem -Path $temp -Filter '_MEI*' -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($null -eq $dir) {
    Write-Output "No _MEI* directory found"
    exit 1
}
Write-Output "Latest extraction dir: $($dir.FullName)"
Get-ChildItem -Path $dir.FullName -Filter 'logo.png' -Recurse -ErrorAction SilentlyContinue | Select-Object FullName
Get-ChildItem -Path $dir.FullName -Filter 'stamp.png' -Recurse -ErrorAction SilentlyContinue | Select-Object FullName

# Clean up: stop the app process so it doesn't lock the file.
Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
