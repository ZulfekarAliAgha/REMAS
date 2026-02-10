# ==============================================================================
# Chrome Extensions - REMAS v0.80 (Beta) by Zulali
# Recover • Export • Merge • Audit • Sort
# ==============================================================================

# 1. SETUP PATHS
$user = "ADMIN" # <-- CHANGE THIS TO YOUR USERNAME. Example: If your Windows user folder is C:\Users\JohnDoe, change the "USERNAME" to "JohnDoe".

if (Test-Path "C:\Users\$user\Desktop\Extensions_1") {
    $source1 = "C:\Users\$user\Desktop\Extensions_1"
    $source2 = "C:\Users\$user\Desktop\Extensions_2"
} else {
    # Fallback: Ignore Desktop, use Live Chrome Folder Only
    $source1 = "C:\Users\$user\AppData\Local\Google\Chrome\User Data\Default\Extensions"
    $source2 = "" # Empty to prevent errors in the loop
}

$dest = "C:\Users\$user\Desktop\Chrome Extensions - REMAS"
$htmlFile = Join-Path $dest "Extensions.html"

# Master table and counters
$idTable = @{} 
$removedFromSource1 = 0
$removedFromSource2 = 0
$sameNameDiffIdCount = 0

if (!(Test-Path $dest)) { New-Item -ItemType Directory -Path $dest -Force }

# 2. COLLECTION & COMPARISON FUNCTION
function Collect-Best-Version($sourcePath, $sourceLabel) {
    if ([string]::IsNullOrWhiteSpace($sourcePath)) { return }
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
<html><head><title>🚀 Chrome Extensions - REMAS v0.80</title>
<style>
    body{font-family:'Segoe UI',sans-serif;padding:40px;background:#f8f9fa;}
    .header-container{display:flex; justify-content:space-between; align-items:flex-end; border-bottom:2px solid #1a73e8; padding-bottom:10px; margin-bottom:10px;}
    h1{margin:0; color:#202124; display: flex; align-items: center;}
    .version-badge {background:#1a73e8; color:white; padding:2px 8px; border-radius:4px; font-size:0.45em; vertical-align:middle; margin-left:10px;}
    .github-icon-link {margin-left: 12px; display: flex; align-items: center;}
    .github-icon-link svg {fill: #24292e; transition: 0.2s;}
    .github-icon-link:hover svg {fill: #1a73e8;}
    .user-info{font-size:1em; color:#5f6368; margin-top:8px; font-weight:500;}
    .highlight{color:#d93025; font-weight:bold;}
    .stats{text-align:right; font-size:0.85em; color:#5f6368; line-height:1.4;}
    .guide-box {background:#fff3cd; border:1px solid #ffeeba; padding:15px; border-radius:8px; margin-bottom:20px; color:#856404; font-size:0.9em;}
    .item{background:white;margin:10px 0;padding:20px;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,0.12);display:flex;justify-content:space-between;align-items:center;}
    .num{font-size:1.2em; font-weight:bold; color:#70757a; margin-right:20px; min-width:30px;}
    .name-link{font-size:1.1em;font-weight:bold;color:#1a73e8;text-decoration:none;}
    .tag-count {font-size:0.7em; background:#f1f3f4; color:#3c4043; padding:3px 10px; border-radius:50px; margin-left:10px; border:1px solid #dadce0;}
    .tag-origin {font-size:0.7em; background:#e8f0fe; color:#1a73e8; padding:3px 10px; border-radius:50px; margin-left:5px; border:1px solid #c2e7ff;}
    .tag-version {font-size:0.7em; background:#e6ffed; color:#22863a; padding:3px 10px; border-radius:50px; margin-left:5px; border:1px solid #bef5cb; font-weight:bold;}
    .id{font-size:0.85em;color:#5f6368;margin-top:4px;}
    a.btn-search{text-decoration:none;background:#1a73e8;color:white;padding:10px 20px;border-radius:5px;font-weight:bold;}
</style></head>
<body>
<div class='header-container'>
    <div class='header-left'>
        <h1>
            🚀 Chrome Extensions - REMAS <span class='version-badge'>v0.80</span>
            <a href='https://github.com/ZulfekarAliAgha/REMAS' target='_blank' class='github-icon-link' title='View on GitHub'>
                <svg height="28" viewBox="0 0 16 16" width="28"><path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z"></path></svg>
            </a>
        </h1>
        <div class='user-info'>User: <span class='highlight'>$user</span> | <span class='highlight'>$totalExtensions</span> Extensions Recovered, Exported, Merged, Audited, Sorted</div>
    </div>
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

# 5. FINALIZING FILES AND HTML

foreach ($ext in $sortedExtensions) {
    $searchUrl = "https://chromewebstore.google.com/search/" + [uri]::EscapeDataString($ext.Name)
    $detailUrl = "https://chrome.google.com/webstore/detail/" + $ext.ID
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
                 "<div><a href='$detailUrl' target='_blank' class='btn-search' style='margin-right:10px;'>View in Store</a>" +
                 "<a href='$searchUrl' target='_blank' class='btn-search'>Search in Store</a></div></div>"
    
    Write-Host "Saving ($counter/$totalExtensions): $folderName" -ForegroundColor Gray
    Copy-Item -Path $ext.Path -Destination (Join-Path $dest $folderName) -Recurse -Force
    $counter++
}

$finalHtml = $htmlHeader + $htmlBody + "</body></html>"
$finalHtml | Out-File -FilePath $htmlFile -Encoding utf8
if (Test-Path $htmlFile) {
    Start-Process $htmlFile;
    Start-Sleep -Milliseconds 500;
}


Write-Host "---REMAS Engine v0.80 has finished successfully. Dashboard launched !---" -ForegroundColor Cyan
Read-Host "Press Enter to exit the console"
exit



