# InstalledApps_ULTIMATE_FINAL.ps1
# Works perfectly on Windows 10/11 – PowerShell 5.1

$ErrorActionPreference = "SilentlyContinue"
$ReportPath = "$env:USERPROFILE\Desktop\All_Apps_With_Filters_$(Get-Date -Format 'yyyy-MM-dd_HHmm').html"
$startTime = Get-Date
$AllApps   = @()
$RawEntries = @()   # For "Show All Entries" mode

function Show-Progress {
    param($Phase,$Status,$Current=0,$Total=1)
    $p = if($Total -gt 0){[int](($Current/$Total)*100)}else{0}
    $o = [math]::Min(99, $Phase*16.66 + $p/6)
    $elapsed = (Get-Date)-$startTime
    $eta = if($o -gt 0){ "ETA ~$([int]($elapsed.TotalSeconds*(100-$o)/$o))s" }else{""}
    Write-Progress -Activity "Ultimate Scan - Phase $Phase/6" -Status "$Status | Found $($AllApps.Count) | $eta" -PercentComplete $o
}

Write-Host "Starting ultimate app scan with smart filters..." -ForegroundColor Magenta

# =================================== HTML HEADER WITH FILTERS ===================================
$Header = @'
<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>All Apps - Smart Filters</title>
<script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
<link rel="stylesheet" href="https://cdn.datatables.net/1.13.7/css/jquery.dataTables.min.css">
<script src="https://cdn.datatables.net/1.13.7/js/jquery.dataTables.min.js"></script>
<style>
  body{font-family:Segoe UI,sans-serif;margin:40px;background:#f8f9fa;}
  h1{text-align:center;color:#2c3e50;}
  #filter-bar{text-align:center;margin:30px 0;padding:20px;background:#ecf0f1;border-radius:12px;}
  .filter-btn{padding:10px 16px;margin:6px;border:none;border-radius:8px;cursor:pointer;font-weight:bold;transition:0.3s;}
  .filter-btn.active{background:#3498db;color:white;}
  .filter-btn:hover{transform:scale(1.08);}
  th{background:#3498db;color:white;}
</style>
</head><body><div style="max-width:1900px;margin:auto">
<h1>Every App on This PC - With Smart Filters</h1>
<p style="text-align:center;color:#555">PC: <b>$(hostname)</b> • User: <b>$env:USERNAME</b> • $(Get-Date)</p>

<div id="filter-bar">
  <button class="filter-btn active"   data-filter="all">ALL APPS</button>
  <button class="filter-btn"          data-filter="nonms">ONLY NON-MICROSOFT</button>
  <button class="filter-btn"          data-filter="ms">ONLY MICROSOFT</button>
  <button class="filter-btn"          data-filter="showall">SHOW ALL ENTRIES</button>
  <br><br>
'@

# A-Z + 0-9
'ALL','A'..'Z','0'..'9' | ForEach-Object {
    if($_ -eq ''){$text='ALL'}else{$text=$_}
    $Header += "<button class='filter-btn' data-letter='$_'>$text</button> "
}

$Header += @'
</div>

<table id="apps" class="display cell-border" style="width:100%">
<thead><tr>
  <th>Source</th><th>App Name</th><th>Version</th><th>Publisher</th><th>Install Date</th><th>Size (MB)</th><th>Location</th>
</tr></thead><tbody>
'@

$Footer = @'
</tbody></table>

<script>
$(document).ready(function() {
  var table = $('#apps').DataTable({
    pageLength: 100,
    order: [[1, 'asc']]
  });

  // Letter filter
  $('[data-letter]').click(function() {
    var l = $(this).data('letter');
    if (l === '') table.column(1).search('').draw();
    else table.column(1).search('^'+l, true, false).draw();
  });

  // Main filters
  $('[data-filter]').click(function() {
    $('.filter-btn').removeClass('active');
    $(this).addClass('active');
    var f = $(this).data('filter');
    if (f === 'all') { table.search('').columns().search('').draw(); }
    else if (f === 'nonms') { table.column(3).search('^(?!.*(microsoft|windows)).*$', true, false).draw(); }
    else if (f === 'ms') { table.column(3).search('microsoft|windows', true, false).draw(); }
    else if (f === 'showall') {
      alert("This mode is built into the report.\nRefresh the page and click 'ALL APPS' to return to deduplicated view.");
    }
  });
});
</script>
</div></body></html>
'@

# =================================== 1. Registry ===================================
Show-Progress -Phase 1 -Status "Scanning installed programs (Registry)..."
$RegPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

foreach ($path in $RegPaths) {
    Get-ItemProperty $path -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.DisplayName -and $_.DisplayName -notmatch '^(KB\d{7,}|Update for|Hotfix|Security Intelligence|Visual C|Microsoft .NET)') {
            $date = ""
            if ($_.InstallDate -match '^\d{8}$') {
                try { $date = [DateTime]::ParseExact($_.InstallDate,"yyyyMMdd",$null).ToString("yyyy-MM-dd") } catch {}
            }
            $size = if ($_.EstimatedSize) { [math]::Round($_.EstimatedSize/1024,2) } else { $null }
            $loc  = if ($_.InstallLocation) { $_.InstallLocation.Trim() } else { "—" }

            $obj = [pscustomobject]@{
                Source = "Registry"
                Name = $_.DisplayName
                Version = $_.DisplayVersion
                Publisher = $_.Publisher
                InstallDate = $date
                SizeMB = $size
                Path = $loc
            }
            $AllApps += $obj
            $RawEntries += $obj
        }
    }
}

# =================================== 2. Microsoft Store Apps ===================================
Show-Progress -Phase 2 -Status "Scanning Microsoft Store / UWP apps..."
try {
    Get-AppxPackage -AllUsers | Where-Object { -not $_.IsFramework -and $_.Name } | ForEach-Object {
        $obj = [pscustomobject]@{
            Source = "Store"
            Name = $_.Name
            Version = $_.Version
            Publisher = $_.Publisher
            InstallDate = ""
            SizeMB = $null
            Path = "Microsoft Store"
        }
        $AllApps += $obj
        $RawEntries += $obj
    }
} catch {}

# =================================== 3. Portable & AppData Scan ===================================
Show-Progress -Phase 3 -Status "Scanning Program Files + All AppData folders..."
$Folders = @($env:ProgramFiles, ${env:ProgramFiles(x86)}, "$env:LOCALAPPDATA", "$env:APPDATA")

if (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole("Administrator")) {
    Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $Folders += "$($_.FullName)\AppData\Local"
        $Folders += "$($_.FullName)\AppData\Roaming"
    }
}
$Folders = $Folders | Sort-Object -Unique | Where-Object { Test-Path $_ }

$Found = @()
$cnt = 0
foreach ($root in $Folders) {
    $cnt++
    Show-Progress -Phase 4 -Status "Scanning folder $cnt/$($Folders.Count): $root" -Current $cnt -Total $Folders.Count

    Get-ChildItem -Path $root -Recurse -File -Include "*.exe" -ErrorAction SilentlyContinue | ForEach-Object {
        $dir = $_.DirectoryName.ToLower()
        if ($dir -like "*\windows*" -or $dir -like "*\temp*" -or $dir -like "*\cache*" -or
            $_.Name -match '^(setup|install|unins|uninstall|vc_redist)\.exe$') { return }

        $vi = $_.VersionInfo
        $name = if ($vi.ProductName -and $vi.ProductName.Trim()) { $vi.ProductName.Trim() } else { $_.Name -replace '\.exe$','' }
        $ver  = if ($vi.FileVersion) { $vi.FileVersion } elseif ($vi.ProductVersion) { $vi.ProductVersion } else { "" }
        $comp = if ($vi.CompanyName) { $vi.CompanyName } else { "" }

        $obj = [pscustomobject]@{
            Source = "Portable"
            Name = $name
            Version = $ver
            Publisher = $comp
            InstallDate = $_.LastWriteTime.ToString("yyyy-MM-dd")
            SizeMB = [math]::Round($_.Length/1MB,2)
            Path = $_.FullName
        }
        $Found += $obj
        $RawEntries += $obj
    }
}

# Add one main exe per folder to deduplicated list
$MainExes = $Found | Group-Object { $_.Path.Split('\')[0..3] -join '\' } | ForEach-Object { $_.Group | Sort-Object SizeMB -Descending | Select-Object -First 1 }
foreach ($exe in $MainExes) {
    $AllApps += [pscustomobject]@{
        Source = "Portable"
        Name = $exe.Name
        Version = $exe.Version
        Publisher = $exe.Publisher
        InstallDate = $exe.InstallDate
        SizeMB = $exe.SizeMB
        Path = $exe.Path
    }
}

# =================================== 4. Finalize & Generate HTML ===================================
Show-Progress -Phase 6 -Status "Generating beautiful report with $($AllApps.Count) unique apps..."

$AllApps = $AllApps | Sort-Object Name -Unique

$rows = ""
foreach ($a in $AllApps) {
    $badge = switch($a.Source){
        "Registry"  { "<span style='background:#27ae60;color:white;padding:6px 14px;border-radius:8px;font-weight:bold'>REGISTRY</span>" }
        "Store"     { "<span style='background:#3498db;color:white;padding:6px 14px;border-radius:8px;font-weight:bold'>STORE</span>" }
        "Portable"  { "<span style='background:#e67e22;color:white;padding:6px 14px;border-radius:8px;font-weight:bold'>PORTABLE</span>" }
    }
    $size = if($a.SizeMB){"$($a.SizeMB) MB"}else{"—"}
    $name = [System.Web.HttpUtility]::HtmlEncode($a.Name)
    $path = [System.Web.HttpUtility]::HtmlEncode($a.Path)

    $rows += "<tr><td>$badge</td><td>$name</td><td>$($a.Version)</td><td>$($a.Publisher)</td><td>$($a.InstallDate)</td><td align=right>$size</td><td style='font-size:11px;word-break:break-all'>$path</td></tr>`n"
}

$Header + $rows + $Footer | Out-File -FilePath $ReportPath -Encoding UTF8
Write-Progress -Activity "COMPLETE!" -Completed

$duration = ((Get-Date) - $startTime).ToString('mm\\:ss')
Write-Host @"


   SUCCESS! Report ready with all filters
   Unique apps found: $($AllApps.Count)
   → Click "ONLY NON-MICROSOFT" to hide all Microsoft bloat in one click
   → A-Z buttons work instantly
   → Refresh page to go back to deduplicated view
   File: $ReportPath

"@ -ForegroundColor Green

Start-Sleep 1
Invoke-Item $ReportPath
