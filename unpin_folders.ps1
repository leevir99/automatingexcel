$foldersToKeep = @("1. ASIAKIRJALAADINTA", "MML lomakkeet", "Valmiit hankkeet")

$shellApp = New-Object -ComObject shell.application
$quickAccess = $shellApp.Namespace(0x417)
$quickAccessItems = $quickAccess.Items()

foreach ($item in $quickAccessItems) {
    $folderPath = $quickAccess.GetFolder.Path
    $folderName = $folderPath.Split("\")[-1]

    if ($foldersToKeep -notcontains $folderName) {
        $quickAccess.Delete($folderPath)
        Write-Host "Unpinned folder: $folderPath"
    }
}

$shellApp.Namespace('shell:::{679F85CB-0220-4080-B29B-5540CC05AAB6}').Self.InvokeVerb('Refresh')
Write-Host "Quick Access refreshed."
