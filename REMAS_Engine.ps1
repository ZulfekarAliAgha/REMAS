# ==============================================================================
# Chrome Extensions - REMAS v0.79 (Beta) by Zulali
# Recover • Export • Merge • Audit • Sort
# ==============================================================================

# 1. SETUP PATHS
$user = "ADMIN" # <-- CHANGE THIS TO YOUR USERNAME. Example: If your Windows user folder is C:\Users\JohnDoe, change the "USERNAME" to "JohnDoe".
$source1 = if (Test-Path "C:\Users\$user\Desktop\Extensions_1") { "C:\Users\$user\Desktop\Extensions_1" } else { "C:\Users\$user\Desktop\Extensions" }
$source2 = "C:\Users\$user\Desktop\Extensions_2"
$dest = "C:\Users\$user\Desktop\Chrome_Extensions_REMAS"
$htmlFile = Join-Path $dest "Extensions.html"

# Master table and counters
$idTable = @{} 
$removedFromSource1 = 0
$removedFromSource2 = 0
$sameNameDiffIdCount = 0

if (!(Test-Path $dest)) { New-Item -ItemType Directory -Path $dest -Force }


Write-Host "--- REMAS: STARTING MERGE & AUDIT ---" -ForegroundColor Yellow

# 2. COLLECTION & COMPARISON FUNCTION
function Collect-Best-Version($sourcePath, $sourceLabel) {
    if (!(Test-Path $sourcePath)) { 
        Write-Host "Skipping: $sourcePath (Not found)" -ForegroundColor Red
        return 
    }
    $folders = Get-ChildItem -Path $sourcePath -Directory
    Write-Host "Scanning $sourceLabel ($($folders.Count) items)..." -ForegroundColor Cyan

    foreach ($f in $folders) {
        $id = $f.Name
        $vFolder = Get-ChildItem -Path $f.FullName -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($vFolder) {
            $vString = $vFolder.Name -replace '_.*',''
            try { $vObject = [System.Version]$vString } catch { $vObject = [System.Version]"0.0.0" }
            $displayName = $id
            $manifestPath = Join-Path $vFolder.FullName "manifest.json"
            if (Test-Path $manifestPath) {
                try {
                    $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
                    $displayName = $manifest.name
                    if ($displayName -like "__MSG_*") {
                        $localePath = Join-Path $vFolder.FullName "_locales\en\messages.json"
                        if (Test-Path $localePath) {
                            $msg = Get-Content $localePath -Raw | ConvertFrom-Json
                            $key = $displayName -replace "__MSG_","" -replace "__",""
                            if ($msg.$key.message) { $displayName = $msg.$key.message }
                        }
                    }
                } catch { }
            }
            if ($idTable.ContainsKey($id)) {
                $idTable[$id].SourceCount = "Both Sources"
                if ($vObject -gt $idTable[$id].Version) {
                    Write-Host "  UPDATE: $displayName -> v$vString (Newer in $sourceLabel)" -ForegroundColor Green
                    if ($idTable[$id].WinnerLabel -eq "Source 1") { $script:removedFromSource1++ } else { $script:removedFromSource2++ }
                    $idTable[$id].Version = $vObject
                    $idTable[$id].Path = $vFolder.FullName
                    $idTable[$id].WinnerLabel = $sourceLabel
                } else {
                    Write-Host "  IGNORE: $displayName (Already have v$($idTable[$id].Version))" -ForegroundColor Gray
                    if ($sourceLabel -eq "Source 1") { $script:removedFromSource1++ } else { $script:removedFromSource2++ }
                }
            } else {
                $idTable[$id] = [PSCustomObject]@{
                    Name         = $displayName
                    ID           = $id
                    Version      = $vObject
                    SourceCount  = "One Source"
                    WinnerLabel  = $sourceLabel
                    Path         = $vFolder.FullName
                }
            }
        }
    }
}

# 3. RUN COLLECTION
Collect-Best-Version $source1 "Source 1"
Collect-Best-Version $source2 "Source 2"

# 4. SORT AND GENERATE HTML
$sortedExtensions = $idTable.Values | Sort-Object Name
$totalExtensions = $sortedExtensions.Count
$counter = 1

$nameGroups = $sortedExtensions | Group-Object Name
foreach ($group in $nameGroups) {
    if ($group.Count -gt 1) { $script:sameNameDiffIdCount += ($group.Count - 1) }
}

$htmlHeader = @"
<html><head><title>Chrome Extensions - Recover, Export, Merge, Audit, Sort</title>
<style>
    body{font-family:'Segoe UI',sans-serif;padding:40px;background:#f8f9fa;} 
    .header-container{display:flex; justify-content:space-between; align-items:flex-end; border-bottom:2px solid #1a73e8; padding-bottom:10px; margin-bottom:10px;}
    h1{margin:0; color:#202124;}
    .stats{text-align:right; font-size:0.85em; color:#5f6368; line-height:1.4;}
    .guide-box {background:#fff3cd; border:1px solid #ffeeba; padding:15px; border-radius:8px; margin-bottom:20px; color:#856404; font-size:0.9em; display:flex; justify-content:space-between; align-items:center;}
    .copy-btn {background:#856404; color:white; border:none; padding:5px 10px; border-radius:4px; cursor:pointer; font-size:0.8em; margin-left:10px;}
    .item{background:white;margin:10px 0;padding:20px;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,0.12);display:flex;justify-content:space-between;align-items:center;} 
    .num{font-size:1.2em; font-weight:bold; color:#70757a; margin-right:20px; min-width:30px;}
    .name-link{font-size:1.1em;font-weight:bold;color:#1a73e8;text-decoration:none;}
    .tag-count {font-size:0.7em; background:#f1f3f4; color:#3c4043; padding:3px 10px; border-radius:50px; margin-left:10px; border:1px solid #dadce0;}
    .tag-origin {font-size:0.7em; background:#e8f0fe; color:#1a73e8; padding:3px 10px; border-radius:50px; margin-left:5px; border:1px solid #c2e7ff;}
    .tag-version {font-size:0.7em; background:#e6ffed; color:#22863a; padding:3px 10px; border-radius:50px; margin-left:5px; border:1px solid #bef5cb; font-weight:bold;}
    a.btn-search{text-decoration:none;background:#1a73e8;color:white;padding:10px 20px;border-radius:5px;font-weight:bold;}
</style>
<script>
    function copyLink() {
        navigator.clipboard.writeText('chrome://extensions');
        alert('Link copied! Paste it into your browser address bar.');
    }
</script></head>
<body>
<div class='header-container'>
    <h1>Chrome Extensions - Recover, Export, Merge, Audit, Sort ($totalExtensions)</h1>
    <div class='stats'>
        $removedFromSource1 duplicates removed from Source 1<br>
        $removedFromSource2 duplicates removed from Source 2<br>
        $sameNameDiffIdCount extensions with same name but different IDs
    </div>
</div>

<div class='guide-box'>
    <b>Pro-Tip for Sideloading:</b> Open <u>chrome://extensions</u> and enable <b>Developer Mode</b>. 
    Click an extension name below to open its folder, then <b>Drag & Drop</b> that folder directly onto the Chrome window to install it instantly!
</div>
"@

$htmlBody = ""
$usedFolderNames = @{}
Write-Host "`n--- FINALIZING FILES AND HTML ---" -ForegroundColor Yellow

foreach ($ext in $sortedExtensions) {
    $searchUrl = "https://chromewebstore.google.com/search/" + [uri]::EscapeDataString($ext.Name)
    $suffix = if ($ext.WinnerLabel -eq "Source 1") { "_1" } else { "_2" }
    $cleanFileName = ($ext.Name -replace '[\\\/\:\*\?\"\<\>\|]', '').Trim()
    $folderName = $cleanFileName + $suffix
    if ($usedFolderNames.ContainsKey($folderName)) { $folderName = $folderName + "_d" }
    $usedFolderNames[$folderName] = $true
    $localPath = "file:///" + (Join-Path $dest $folderName).Replace("\","/")

    $htmlBody += "<div class='item'><div style='display:flex; align-items:center;'><div class='num'>$counter.</div><div>" +
                 "<a href='$localPath' class='name-link'>$($ext.Name)</a>" +
                 "<span class='tag-count'>$($ext.SourceCount)</span>" +
                 "<span class='tag-origin'>$($ext.WinnerLabel)</span>" +
                 "<span class='tag-version'>v$($ext.Version)</span>" +
                 "<div style='font-size:0.85em;color:#5f6368;margin-top:4px;'>ID: $($ext.ID)</div></div></div>" +
                 "<a href='$searchUrl' target='_blank' class='btn-search'>Search in Store</a></div>"
    
    Write-Host "Saving ($counter/$totalExtensions): $folderName" -ForegroundColor Gray
    Copy-Item -Path $ext.Path -Destination (Join-Path $dest $folderName) -Recurse -Force
    $counter++
}

$finalHtml = $htmlHeader + $htmlBody + "</body></html>"
$finalHtml | Out-File -FilePath $htmlFile -Encoding utf8
Write-Host "`n--- REMAS COMPLETE ---" -ForegroundColor Cyan
Invoke-Item $dest
