$target = $PSScriptRoot
Set-Location $env:TEMP
foreach ($folder in [Environment]::GetFolderPath("Desktop"), [Environment]::GetFolderPath("Programs")) {
    Remove-Item -Force -ErrorAction SilentlyContinue (Join-Path $folder "Dockit V*.lnk")
}
Remove-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Dockit" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path $target -Recurse -Force -ErrorAction SilentlyContinue
