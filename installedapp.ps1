# InstalledApps.ps1 – FINAL WORKING VERSION (PowerShell 5.1 compatible)
# Registry + Program Files + ALL AppData → Beautiful HTML report with detailed progress

$ErrorActionPreference = "SilentlyContinue"
$ReportPath = "$env:USERPROFILE\Desktop\Installed_Apps_$(Get-Date -Format 'yyyy-MM-dd_HH-mm').html"
$startTime = Get-Date
$AllApps   = @()

function Show-Progress {
    param($Phase, $Status, $Current = 0, $Total = 1)
    $percent = if ($Total -gt 0) { [int](($Current/$Total)*100) } else { 0 }
    $overall = [math]::Min(99, $Phase*20 + $percent/5)

    $elapsed = (Get-Date) - $startTime
    $eta = if ($overall -gt 0) {
        $sec = [int]($elapsed.TotalSeconds * (100-$overall)/$overall)
        "ETA ~$sec sec"
    } else { "" }

    Write-Progress -Activity "Scanning System – Phase $Phase/5" `
                   -Status "$Status | Found $($AllApps.Count) apps | $eta" `
                   -PercentComplete $overall
}

Write-Host "Starting full system scan (Registry + all folders + AppData)..." -ForegroundColor Cyan

# =================================== HTML HEADER ===================================
$Header = @'
<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>All Installed & Portable Apps</title>
<script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
<link rel="stylesheet" href="https://cdn.datatables.net/1.13.7/css/jquery.dataTables.min.css">
<script src="https://cdn.datatables.net/1.13.7/js/jquery.dataTables.min.js"></script>
<style>
  body{font-family:Segoe UI,sans-serif;margin:40px;background:#f8f9fa;}
  h1{text-align:center;color:#2c3e50;}
  table{width:100%;box-shadow:0 4px 20px rgba(0,0,0,0.1);}
  th{background:#3498db;color:white;}
  .reg{background:#d5f5e0;}
  .scan{background:#fff3cd;}
</style>
</head><body><div style="max-width:1600px;margin:auto">
<h1>Complete Application List – $(hostname)</h1>
<p style="text-align:center;color:#555"><b>User:</b> $env:USERNAME &nbsp; <b>Generated:</b> $(Get-Date)</p>
<table id="apps" class="display cell-border" style="width:100%">
<thead><tr>
  <th>Source</th><th>Program Name</th><th>Version</th><th>Publisher</th><th>Install Date</th><th>Size (MB)</th><th>Path</th>
</tr></thead><tbody>
'@

$Footer = @'
</tbody></table>
<script type="text/javascript">
  $(document).ready(function(){ $('#apps').DataTable({pageLength:100,order:[[1,"asc"]]}); });
</script>
</div></body></html>
'@

# =================================== PHASE 1 – Registry ===================================
Show-Progress -Phase 1 -Status "Reading registry uninstall keys..."

$RegPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

foreach ($Path in $RegPaths) {
    Get-ItemProperty $Path | ForEach-Object {
        if ($_.DisplayName -and $_.DisplayName -notmatch '^(KB\d{7,}|Update for|Security Intelligence|Hotfix|Microsoft Visual|Microsoft .NET)') {
            $date = ""
            if ($_.InstallDate -match '^\d{8}$') {
                try { $date = [DateTime]::ParseExact($_.InstallDate,"yyyyMMdd",$null).ToString("yyyy-MM-dd") } catch {}
            }
            $sizeMB = if ($_.EstimatedSize) { [math]::Round($_.EstimatedSize/1024,2) } else { $null }
            $loc = if ($_.InstallLocation) { $_.InstallLocation.Trim() } else { "—" }

            $AllApps += [pscustomobject]@{
                Source = "Registry"; Name = $_.DisplayName; Version = $_.DisplayVersion
                Publisher = $_.Publisher; InstallDate = $date; SizeMB = $sizeMB; Path = $loc
            }
        }
    }
}

# =================================== PHASE 2 – Build folder list ===================================
Show-Progress -Phase 2 -Status "Collecting folders to scan (including AppData)..."

$Folders = @($env:ProgramFiles, ${env:ProgramFiles(x86)}, "$env:LOCALAPPDATA", "$env:APPDATA")

if (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole("Administrator")) {
    Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $Folders += "$($_.FullName)\AppData\Local"
        $Folders += "$($_.FullName)\AppData\Roaming"
    }
}
$Folders = $Folders | Sort-Object -Unique | Where-Object { Test-Path $_ }

# =================================== PHASE 3 – Scan all folders ===================================
Show-Progress -Phase 3 -Status "Scanning $($Folders.Count) folders for .exe files..." -Current 0 -Total $Folders.Count

$Found = @()
$counter = 0
foreach ($Root in $Folders) {
    $counter++
    Show-Progress -Phase 3 -Status "Scanning: $Root" -Current $counter -Total $Folders.Count

    Get-ChildItem -Path $Root -Recurse -File -Include "*.exe" -ErrorAction SilentlyContinue | ForEach-Object {
        $dir = $_.DirectoryName.ToLower()
        if ($dir -like "*\windows*" -or $dir -like "*\temp*" -or $dir -like "*\cache*" -or
            $_.Name -match '^(setup|install|unins|uninstall|vc_redist)\.exe$') { return }

        $vi = $_.VersionInfo
        $name = if ($vi.ProductName -and $vi.ProductName.Trim()) { $vi.ProductName.Trim() } else { $_.Name -replace '\.exe$','' }
        $ver  = if ($vi.FileVersion) { $vi.FileVersion } elseif ($vi.ProductVersion) { $vi.ProductVersion } else { "" }
        $comp = if ($vi.CompanyName) { $vi.CompanyName } else { "" }

        $Found += [pscustomobject]@{
            Name = $name; Version = $ver; Company = $comp
            Path = $_.FullName; Folder = $_.DirectoryName
            SizeMB = [math]::Round($_.Length/1MB,2)
            Modified = $_.LastWriteTime.ToString("yyyy-MM-dd")
        }
    }
}

# Keep only the largest .exe per folder
$Main = $Found | Group-Object Folder | ForEach-Object { $_.Group | Sort-Object SizeMB -Desc | Select-Object -First 1 }

foreach ($exe in $Main) {
    $match = $AllApps | Where-Object { $_.Name -eq $exe.Name } | Select-Object -First 1
    $date = if ($match) { $match.InstallDate } else { $exe.Modified }

    $AllApps += [pscustomobject]@{
        Source = "File Scan"; Name = $exe.Name; Version = $exe.Version
        Publisher = $exe.Company; InstallDate = $date; SizeMB = $exe.SizeMB; Path = $exe.Path
    }
}

# =================================== PHASE 5 – Finalize & save ===================================
Show-Progress -Phase 5 -Status "Sorting $($AllApps.Count) apps and generating HTML..."

$AllApps = $AllApps | Sort-Object Name -Unique

$rows = ""
foreach ($app in $AllApps) {
    $badge = if ($app.Source -eq "Registry") {
        '<span style="background:#27ae60;color:white;padding:4px 10px;border-radius:5px;font-size:11px">REGISTRY</span>'
    } else {
        '<span style="background:#e67e22;color:white;padding:4px 10px;border-radius:5px;font-size:11px">PORTABLE</span>'
    }
    $size = if ($app.SizeMB) { "$($app.SizeMB) MB" } else { "—" }
    $name = [System.Web.HttpUtility]::HtmlEncode($app.Name)
    $path = [System.Web.HttpUtility]::HtmlEncode($app.Path)

    $rows += "<tr><td>$badge</td><td>$name</td><td>$($app.Version)</td><td>$($app.Publisher)</td><td>$($app.InstallDate)</td><td align=right>$size</td><td style='font-size:11px;word-break:break-all'>$path</td></tr>`n"
}

$Header + $rows + $Footer | Out-File -FilePath $ReportPath -Encoding UTF8

Write-Progress -Activity "DONE!" -Completed

$duration = ((Get-Date) - $startTime).ToString('mm\\:ss')

Write-Host @"


   SCAN COMPLETE!
   Found: $($AllApps.Count) applications
   Duration: $duration
   Report → $ReportPath

"@ -ForegroundColor Green

Start-Sleep 1
Invoke-Item $ReportPath
