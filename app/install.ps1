$ErrorActionPreference = "Stop"
$appName = "Dockit V1.28"
$source = $PSScriptRoot
$target = Join-Path $env:LOCALAPPDATA "Programs\Dockit"
$oldTarget = Join-Path $env:LOCALAPPDATA "Programs\PDF Gen"

try {
    Write-Host "Memeriksa Python..."
    $python = (& python -c "import sys; print(sys.executable)").Trim()
    $pythonw = Join-Path (Split-Path $python) "pythonw.exe"
    if (-not (Test-Path $pythonw)) { throw "pythonw.exe tidak ditemukan di $python" }

    Write-Host "Memasang library..."
    & $python -m pip install --quiet pywin32 customtkinter pypdfium2 pillow lxml tkinterdnd2 openpyxl
    if ($LASTEXITCODE -ne 0) { throw "Library gagal dipasang." }
    & $python -c "import win32com.client, customtkinter, pypdfium2, PIL, lxml, tkinterdnd2, openpyxl"
    if ($LASTEXITCODE -ne 0) { throw "Library tidak bisa dimuat." }

    Write-Host "Menyalin aplikasi..."
    New-Item -ItemType Directory -Force -Path $target | Out-Null
    foreach ($file in "dockit.pyw", "di_data.py", "di_ui.py", "dockit_icons.py", "dockit_motion.py", "calendar_page.py", "di_template.xlsx", "po_template.xlsx", "dockit.ico", "dockit.png", "uninstall.ps1") {
        Copy-Item -Force (Join-Path $source $file) $target
    }
    $iconTarget = Join-Path $target "icons"
    if (Test-Path $iconTarget) { Remove-Item -Recurse -Force $iconTarget }
    Copy-Item -Recurse -Force (Join-Path $source "icons") $iconTarget
    $script = Join-Path $target "dockit.pyw"
    $icon = Join-Path $target "dockit.ico"

    Write-Host "Membersihkan PDF Gen versi lama..."
    $shell = New-Object -ComObject WScript.Shell
    $folders = [Environment]::GetFolderPath("Desktop"), [Environment]::GetFolderPath("Programs")
    foreach ($folder in $folders) {
        Remove-Item -Force -ErrorAction SilentlyContinue (Join-Path $folder "PDF Gen V*.lnk")
        Remove-Item -Force -ErrorAction SilentlyContinue (Join-Path $folder "Dockit V*.lnk")
    }
    Remove-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\PDFGen" -Recurse -Force -ErrorAction SilentlyContinue
    if (Test-Path $oldTarget) { Remove-Item -Recurse -Force -ErrorAction SilentlyContinue $oldTarget }

    Write-Host "Membuat shortcut..."
    foreach ($folder in $folders) {
        $link = $shell.CreateShortcut((Join-Path $folder "$appName.lnk"))
        $link.TargetPath = $pythonw
        $link.Arguments = "`"$script`""
        $link.WorkingDirectory = $target
        $link.IconLocation = $icon
        $link.Description = $appName
        $link.Save()
    }

    $key = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Dockit"
    New-Item -Path $key -Force | Out-Null
    $uninstall = "powershell -NoProfile -ExecutionPolicy Bypass -File `"$(Join-Path $target 'uninstall.ps1')`""
    $values = @{
        DisplayName = $appName
        DisplayVersion = "1.28"
        DisplayIcon = $icon
        InstallLocation = $target
        UninstallString = $uninstall
    }
    foreach ($name in $values.Keys) {
        Set-ItemProperty -Path $key -Name $name -Value $values[$name]
    }
    Set-ItemProperty -Path $key -Name NoModify -Value 1 -Type DWord
    Set-ItemProperty -Path $key -Name NoRepair -Value 1 -Type DWord

    Write-Host ""
    Write-Host "$appName berhasil diinstall. Buka dari shortcut di Desktop atau Start Menu."
    Write-Host "Pengaturan, password, dan backup dari PDF Gen tetap dipakai."
} catch {
    Write-Host ""
    Write-Host "Install gagal: $_" -ForegroundColor Red
    Write-Host "Pastikan Python sudah terpasang dan bisa dijalankan dengan perintah python."
}
