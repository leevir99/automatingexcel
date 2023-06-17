$foldersToKeep = @("1. ASIAKIRJALAADINTA", "MML lomakkeet", "Valmiit hankkeet")
$foldersToKeep += (Get-Item -Path 'shell:::{679F85CB-0220-4080-B29B-5540CC05AAB6}').GetFolder.Items() | Select-Object -Last 5 | ForEach-Object { $_.Path }

$shellApp = New-Object -ComObject shell.application
$quickAccess = $shellApp.Namespace(0x417)
$quickAccessItems = $quickAccess.Items()

foreach ($item in $quickAccessItems) {
    $folderPath = $quickAccess.GetFolder.Path
    $folderName = $folderPath.Split("\")[-1]

    if ($foldersToKeep -notcontains $folderPath) {
        $quickAccess.Delete($folderPath)
        Write-Host "Unpinned folder: $folderPath"
    }
}

$shellApp.Namespace('shell:::{679F85CB-0220-4080-B29B-5540CC05AAB6}').Self.InvokeVerb('Refresh')
Write-Host "Quick Access refreshed."



----
@echo off
powershell -ExecutionPolicy Bypass -File "C:\path\to\unpin_folders.ps1"
pause
