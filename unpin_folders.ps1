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
powershell -ExecutionPolicy Bypass -File "%USERPROFILE%\Desktop\unpin_folders.ps1"
pause


----
@echo off
powershell -ExecutionPolicy Bypass -Command "& { $foldersToKeep = @('1. ASIAKIRJALAADINTA', 'MML lomakkeet', 'Valmiit hankkeet'); $foldersToKeep += (Get-Item -Path 'shell:::{679F85CB-0220-4080-B29B-5540CC05AAB6}').GetFolder.Items() | Select-Object -Last 5 | ForEach-Object { $_.Path }; $shellApp = New-Object -ComObject shell.application; $quickAccess = $shellApp.Namespace(0x417); $quickAccessItems = $quickAccess.Items(); foreach ($item in $quickAccessItems) { $folderPath = $quickAccess.GetFolder.Path; $folderName = $folderPath.Split('\')[-1]; if ($foldersToKeep -notcontains $folderPath) { $quickAccess.Delete($folderPath); Write-Host 'Unpinned folder: ' $folderPath; } } $shellApp.Namespace('shell:::{679F85CB-0220-4080-B29B-5540CC05AAB6}').Self.InvokeVerb('Refresh'); Write-Host 'Quick Access refreshed.' }"
pause




----
// Open the site
window.location.href = 'https://intra6.tuotanto.op.fi/eportti';

// Function to check if the page has finished loading
function checkPageLoaded() {
  // Check if the button is available
  var button = document.querySelector('a.btn.btn-default');
  
  if (button && button.textContent === 'Tietopalveluun') {
    // Click the button
    button.click();
  } else {
    // Retry after a short delay if the button is not yet available
    setTimeout(checkPageLoaded, 100);
  }
}

// Check if the page has finished loading
window.addEventListener('load', checkPageLoaded);

