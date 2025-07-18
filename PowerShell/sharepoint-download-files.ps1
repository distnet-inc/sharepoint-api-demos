# Install the unified Graph module if needed (run once)
# Install-Module Microsoft.Graph -Scope CurrentUser -Force

Import-Module Microsoft.Graph.Sites
Import-Module Microsoft.Graph.Files

# === Configuration ===
$config = @{
    TenantId           = "{TenantID}"
    ClientId           = "{ClientId}"
    CertificateThumb   = "{Thumbprint du certificat}"
    SiteHostname       = "{Adresse SharePoint}"
    SiteName           = "{Nom du site Sharepoint}" # Nom du site Sharepoint (celui qu'on retrouve dans l'URL)
    LibraryName        = "{Nom de la librairie}" # Le DisplayName de la librairie Sharepoint (Par défaut "Rapports")
    ReportDate         = "{Date du rapport sous format yyyy-MM-dd}" # Exemple: "2025-07-17"
    OutputFolder       = "C:\Temp"
}

# === Settings ===
$ErrorActionPreference = "Stop"
$ProgressPreference = 'SilentlyContinue'

# === Locate Certificate ===
Write-Host "`n🔐 Locating certificate..." -ForegroundColor Cyan
$certificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $config.CertificateThumb }

if (-not $certificate) {
    Write-Error "❌ Certificate with thumbprint $($config.CertificateThumb) not found in LocalMachine\My store."
    exit 1
}
Write-Host "✅ Certificate found." -ForegroundColor Green

# === Connect to Microsoft Graph ===
Write-Host "`n🌐 Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -ClientId $config.ClientId -TenantId $config.TenantId -Certificate $certificate
    Write-Host "✅ Connected to Microsoft Graph." -ForegroundColor Green
} catch {
    Write-Error "❌ Connection to Microsoft Graph failed: $_"
    exit 1
}

# === Resolve Site ===
Write-Host "`n📡 Resolving site '$($config.SiteName)'..." -ForegroundColor Cyan
try {
    $site = Get-MgSite -SiteId "$($config.SiteHostname):/sites/$($config.SiteName)"
    Write-Host "✅ Site resolved: $($site.Id)" -ForegroundColor Green
} catch {
    Write-Error "❌ Failed to resolve site: $_"
    exit 1
}

# === Get Document Library ===
Write-Host "`n📁 Locating document library '$($config.LibraryName)'..." -ForegroundColor Cyan
try {
    $list = Get-MgSiteList -SiteId $site.Id | Where-Object { $_.DisplayName -eq $config.LibraryName }
    if (-not $list) {
        Write-Error "❌ Document library '$($config.LibraryName)' not found."
        exit 1
    }
    Write-Host "✅ Library found: $($list.Id)" -ForegroundColor Green
} catch {
    Write-Error "❌ Failed to retrieve document library: $_"
    exit 1
}

# === Query Files by Metadata ===
Write-Host "`n🔍 Searching for files with ReportDate = '$($config.ReportDate)'..." -ForegroundColor Cyan
try {
    $files = Get-MgSiteListItem -SiteId $site.Id -ListId $list.Id `
        -Filter "fields/ReportDate eq '$($config.ReportDate)'" -ExpandProperty "driveItem"

    if (-not $files) {
        Write-Warning "⚠️ No files found for ReportDate = '$($config.ReportDate)'."
        exit 0
    }

    Write-Host "✅ Found $($files.Count) file(s)." -ForegroundColor Green
} catch {
    Write-Error "❌ Failed to search files by metadata: $_"
    exit 1
}

# === Prepare Output Directory ===
if (-not (Test-Path $config.OutputFolder)) {
    Write-Host "`n📂 Creating output folder: $($config.OutputFolder)" -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $config.OutputFolder | Out-Null
}

# === Download Files ===
Write-Host "`n⬇️ Downloading files..." -ForegroundColor Cyan
$counter = 0

foreach ($file in $files) {
    if ($file.DriveItem.File) {
        $name = $file.DriveItem.Name
        $reportType = $file.Fields.AdditionalProperties["ReportType"]
        $itemId = $file.DriveItem.Id
        $driveId = $file.DriveItem.ParentReference.DriveId
        $outPath = Join-Path $config.OutputFolder $name

        try {
            Get-MgDriveItemContent -DriveId $driveId -DriveItemId $itemId -OutFile $outPath
            Write-Host "✅ Downloaded: [$reportType] $name" -ForegroundColor Green
            $counter++
        } catch {
            Write-Warning "⚠️ Failed to download '$name': $_"
        }
    }
}

# === Summary ===
Write-Host "`n🎉 Completed. Downloaded $counter of $($files.Count) file(s) to: $($config.OutputFolder)" -ForegroundColor Green