Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host " Las Villas de Bayamon WEBSITE BUILDER"
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host " Projecto web gratuito"
Write-Host " Credits and support: studio@andriani.it"
Write-Host "=============================================" -ForegroundColor white
Write-Host ""






# =========================================================
# EXCEL WEBSITE BUILDER WITH AUTO MENU + SOCIAL PREVIEWS
# =========================================================

# =========================================================
# PATHS
# =========================================================

$BasePath = $PSScriptRoot

$ExcelFile = Join-Path $BasePath "lasvillasweb.xlsx"
$RepoPath  = Join-Path $BasePath "gitrepo"
$BuildPath = Join-Path $RepoPath "build"


Write-Host "STEP 1"
Write-Host "Base path:"
Write-Host $BasePath

Write-Host ""
Write-Host "STEP 2"
Write-Host "Excel file path:"
Write-Host $ExcelFile

Write-Host ""
Write-Host "Exists:"
Write-Host (Test-Path $ExcelFile)

if (-not (Test-Path $ExcelFile)) {
    throw "Excel file not found: $ExcelFile"
}

Write-Host ""
Write-Host "STEP 3"




$AssetsPath = Join-Path $BuildPath "assets"

if (-not (Test-Path $AssetsPath)) {
    New-Item -ItemType Directory -Path $AssetsPath -Force | Out-Null
}

$DocsPath = Join-Path $BuildPath "documents"

if (-not (Test-Path $DocsPath)) {
    New-Item -ItemType Directory -Path $DocsPath -Force | Out-Null
}

# =========================================================
# CONFIG
# =========================================================

$BaseUrl = "https://las-villas-de-bayamon.onrender.com"
$LogoUrl = "$BaseUrl/assets/logo.png"

# =========================================================
# LOGO PIPELINE
# =========================================================

$LocalLogo  = Join-Path $BasePath "assets\logo.png"
$TargetLogo = Join-Path $BuildPath "assets\logo.png"

if (Test-Path $LocalLogo) {
    Copy-Item $LocalLogo $TargetLogo -Force
}

# =========================================================
# VALIDATION
# =========================================================

if (-not (Test-Path $ExcelFile)) { throw "Excel file not found" }
if (-not (Test-Path $BuildPath)) { New-Item $BuildPath -ItemType Directory -Force | Out-Null }




# =========================================================
# COPY DOCUMENTS (PDFs)
# =========================================================
# =========================================================

$LocalDocs = Join-Path $BasePath "documents"


if (Test-Path $LocalDocs) {

    Get-ChildItem $LocalDocs -Filter *.pdf | ForEach-Object {

        Copy-Item $_.FullName `
            -Destination $DocsPath `
            -Force

        Write-Host "PDF copied: $($_.Name)" -ForegroundColor Green
    }
}
else {
    Write-Host "No local documents folder found. Existing website PDFs preserved." -ForegroundColor Yellow
}


# Remove old generated html files

Get-ChildItem $BuildPath -Filter *.html -ErrorAction SilentlyContinue |
Remove-Item -Force

# Remove old Excel export folders

Get-ChildItem $BuildPath -Directory |
Where-Object {
    $_.Name -like "*_file"
} |
Remove-Item -Recurse -Force




# =========================================================
# EXCEL EXPORT
# =========================================================

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$pages = @()

try {

    $workbook = $excel.Workbooks.Open($ExcelFile)



# =========================================================
# LOAD OPTIONAL HEADER / FOOTER SHEETS
# =========================================================

$HeaderHtml = ""
$FooterHtml = ""

try {

    $headerSheet = $workbook.Worksheets.Item("#header")

    $HeaderFile = Join-Path $BuildPath "_header_temp.html"

    $excel.ActiveWorkbook.PublishObjects.Add(
        1,
        $HeaderFile,
        "#header",
        "",
        0,
        ""
    ).Publish($true)

  if (Test-Path $HeaderFile) {

    $HeaderHtml = Get-Content $HeaderFile -Raw

    # remove only html/head/body tags but keep styles

    $HeaderHtml = $HeaderHtml -replace '(?is)<!DOCTYPE.*?>',''
    $HeaderHtml = $HeaderHtml -replace '(?is)</?html[^>]*>',''
    $HeaderHtml = $HeaderHtml -replace '(?is)</?body[^>]*>',''

    Remove-Item $HeaderFile -Force
}

} catch {}

try {

    $footerSheet = $workbook.Worksheets.Item("#footer")

    $FooterFile = Join-Path $BuildPath "_footer_temp.html"

    $excel.ActiveWorkbook.PublishObjects.Add(
        1,
        $FooterFile,
        "#footer",
        "",
        0,
        ""
    ).Publish($true)

   if (Test-Path $FooterFile) {

    $FooterHtml = Get-Content $FooterFile -Raw

    # remove only html/body wrappers but keep styles

    $FooterHtml = $FooterHtml -replace '(?is)<!DOCTYPE.*?>',''
    $FooterHtml = $FooterHtml -replace '(?is)</?html[^>]*>',''
    $FooterHtml = $FooterHtml -replace '(?is)</?body[^>]*>',''

    Remove-Item $FooterFile -Force
}

} catch {}



foreach ($ws in $workbook.Worksheets) {

    # System sheets
    if ($ws.Name.StartsWith("#")) {
        continue
    }
	if ($ws.Name -in @("foglio1","sheet1")) { continue }



        Write-Host "Exporting: $($ws.Name)"

        $slug = ($ws.Name.ToLower() -replace '[^a-z0-9]+','-').Trim('-')
        $htmlFile = Join-Path $BuildPath ($slug + ".html")

        try {
            $ws.Activate()

            $excel.ActiveWorkbook.PublishObjects.Add(
                1,
                $htmlFile,
                $ws.Name,
                "",
                0,
                ""
            ).Publish($true)

            $pages += [PSCustomObject]@{
                Name = $ws.Name
                Slug = $slug
                File = $htmlFile
            }

        }
        catch {
            Write-Host "FAILED: $($ws.Name)" -ForegroundColor Red
        }
    }

    $workbook.Close($false)
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# =========================================================
# MENU
# =========================================================

$nav = ""
foreach ($p in $pages) {
    $nav += "<a href='$($p.Slug).html'>$($p.Name)</a>`n"
}
$menuCss = @"
<style>

/* =========================
   HAMBURGER NAVIGATION
========================= */

.site-nav {
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    background: #713B3B;
    z-index: 999999;
    font-family: Arial, sans-serif;
}

/* top bar */
.nav-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 12px 16px;
    color: white;
}

/* hamburger button */

.menu-text {
    font-size: 12px;
}

.nav-toggle {
    font-size: 26px;
    cursor: pointer;
    user-select: none;
}

/* menu container */
.nav-links {
    display: none;
    flex-direction: column;
    background: #5a2f2f;
    padding: 10px;
    gap: 6px;

    max-height: 70vh;
    overflow-y: auto;
}

/* open state */
.nav-links.open {
    display: flex;
}

/* links */
.nav-links a {
    color: white;
    text-decoration: none;
    padding: 10px;
    border-radius: 6px;
    background: rgba(255,255,255,0.12);
}

.nav-links a:hover {
    background: rgba(255,255,255,0.25);
}

/* body offset */
body {
    padding-top: 60px;
    margin: 0;
}

/* wrapper */
.excel-wrapper {
    width: 100%;
    overflow-x: auto;
}

</style>
"@



$menuScript = @"
<script>
function toggleMenu() {
    const el = document.getElementById('navLinks');
    if (!el) return;
    el.classList.toggle('open');
}

document.addEventListener('DOMContentLoaded', function () {
    document.querySelectorAll('.nav-links a').forEach(a => {
        a.addEventListener('click', function () {
            const el = document.getElementById('navLinks');
            if (el) el.classList.remove('open');
        });
    });
});
</script>
"@

# =========================================================
# SOCIAL META (FIXED SAFE VERSION)
# =========================================================
$ogMeta = @"
<meta charset="UTF-8">

<meta property="og:title" content="Las Villas de Bayamon" />
<meta property="og:description" content="Projecto web totalmente gratuito" />
<meta property="og:type" content="website" />
<meta property="og:url" content="$BaseUrl" />
<meta property="og:image" content="$LogoUrl" />

<meta name="twitter:card" content="summary_large_image" />
<meta name="twitter:title" content="Las Villas de Bayamon" />
<meta name="twitter:description" content="Projecto web totalmente gratuito" />
<meta name="twitter:image" content="$LogoUrl" />
"@

# =========================================================
# INJECT HTML (FIXED)
# =========================================================

$utf8 = New-Object System.Text.UTF8Encoding($false)

Get-ChildItem $BuildPath -Filter *.html |
Where-Object {
    $_.Name -notlike "_header_temp*" -and
    $_.Name -notlike "_footer_temp*"
} | ForEach-Object {

    Write-Host "Updating $($_.Name)"

    # =========================
    # SAFE READ
    # =========================
    $bytes = [System.IO.File]::ReadAllBytes($_.FullName)
    $text  = [System.Text.Encoding]::GetEncoding(1252).GetString($bytes)

    $content = [System.Text.Encoding]::UTF8.GetString(
        [System.Text.Encoding]::UTF8.GetBytes($text)
    )

    # =========================
    # FIX ENCODING ISSUES
    # =========================
    $content = $content -replace [char]0xA0,' '

    $content = $content.Replace("Ã¡","a")
    $content = $content.Replace("Ã©","e")
    $content = $content.Replace("Ã­","i")
    $content = $content.Replace("Ã³","o")
    $content = $content.Replace("Ãº","u")
    $content = $content.Replace("Ã±","n")

    # =========================
    # ENSURE HEAD EXISTS
    # =========================
    if ($content -notmatch "<head") {
        $content = $content -replace "<html[^>]*>", "$0<head></head>"
    }

    # =========================
    # VIEWPORT
    # =========================
    if ($content -notmatch "viewport") {
        $content = $content -replace "</head>",
        "<meta name='viewport' content='width=device-width, initial-scale=1'></head>"
    }

    # =========================
    # OG META (SAFE INSERT)
    # =========================




$AnalyticsCode = @"
<script data-goatcounter="https://lasvillasdebayamon.goatcounter.com/count"
        async
        src="https://gc.zgo.at/count.js"></script>
"@



$content = $content -replace "</head>",
"$ogMeta`n$menuCss`n$AnalyticsCode`n$menuScript</head>"



    # =========================
    # HEADER / FOOTER CLEAN INSERT
    # =========================

    # avoid breaking if Excel exported full HTML
    $bodyOpen = $content -match "<body[^>]*>"

    if ($bodyOpen) {

$content = $content -replace "<body[^>]*>",
"`$0
<div class='site-nav'>

    <div class='nav-bar'>

<div class='nav-toggle' onclick='toggleMenu()'>&#9776;  <span class='menu-text'>Menu</span></div>
    </div>

    <div class='nav-links' id='navLinks'>
        $nav
    </div>

</div>

$HeaderHtml

<div class='excel-wrapper'>
"

    }

    # close wrapper safely
    if ($content -match "</body>") {
 $content = $content -replace "</body>",
"</div>

$FooterHtml

</body>"
    }

    # =========================
    # WRITE FILE
    # =========================
    [System.IO.File]::WriteAllText($_.FullName, $content, $utf8)
}

# =========================================================
# INDEX
# =========================================================

if ($pages.Count -gt 0) {
    Copy-Item `
        (Join-Path $BuildPath ($pages[0].Slug + ".html")) `
        (Join-Path $BuildPath "index.html") `
        -Force
}



# =========================================================
# CURRENT SOURCE FILE
# =========================================================

$SourcePath = Join-Path $RepoPath "source"

if (-not (Test-Path $SourcePath)) {
    New-Item -ItemType Directory -Path $SourcePath -Force | Out-Null
}

Copy-Item `
    $ExcelFile `
    (Join-Path $SourcePath "lasvillasweb.xlsx") `
    -Force

Write-Host "Current source workbook updated" -ForegroundColor Green


# =========================================================
# BACKUP EXCEL SOURCE FILE
# =========================================================

$SourcePath   = Join-Path $RepoPath "source"
$BackupFolder = Join-Path $SourcePath "backups"

if (-not (Test-Path $SourcePath)) {
    New-Item -ItemType Directory -Path $SourcePath -Force | Out-Null
}

if (-not (Test-Path $BackupFolder)) {
    New-Item -ItemType Directory -Path $BackupFolder -Force | Out-Null
}

$DateStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

$BackupFile = Join-Path `
    $BackupFolder `
    "lasvillasweb_$DateStamp.xlsx"

Copy-Item $ExcelFile $BackupFile -Force

Write-Host "Excel backup created:" `
    (Split-Path $BackupFile -Leaf) `
    -ForegroundColor Green



# =========================================================
# PUBLISH LATEST EXCEL FILE TO WEBSITE
# =========================================================

$PublicSourcePath = Join-Path $BuildPath "source"

if (-not (Test-Path $PublicSourcePath)) {
    New-Item -ItemType Directory -Path $PublicSourcePath -Force | Out-Null
}

Copy-Item `
    $ExcelFile `
    (Join-Path $PublicSourcePath "lasvillasweb.xlsx") `
    -Force

Write-Host "Public Excel source published to website" -ForegroundColor Green

# =========================================================
# PREVIEW OR PUBLISH
# =========================================================

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host " BUILD COMPLETED"
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "1 = Preview locally"
Write-Host "2 = Publish online"
Write-Host ""

$choice = Read-Host "Select option"

if ($choice -eq "1") {

    $IndexFile = Join-Path $BuildPath "index.html"

    if (Test-Path $IndexFile) {
        Start-Process $IndexFile
    }

    Write-Host ""
    Write-Host "Preview mode only. Nothing uploaded." -ForegroundColor Yellow
    exit
}

if ($choice -ne "2") {
    Write-Host "Cancelled." -ForegroundColor Yellow
    exit
}

# =========================================================
# DEPLOY
# =========================================================

Set-Location $RepoPath

git add build
git add source

$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

git commit -m "Website update $Timestamp"

if ($LASTEXITCODE -ne 0) {
    Write-Host "No changes detected."
    exit
}

git pull --rebase origin main
git push origin main

Write-Host ""
Write-Host "PUBLISHED ONLINE" -ForegroundColor Green