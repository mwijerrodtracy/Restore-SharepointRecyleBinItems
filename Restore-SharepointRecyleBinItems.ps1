 <#
Restore-SharePointRecycleBinFolder.ps1
-------------------------------------
Purpose:
  1) Connect to a SharePoint Online site with PnP.PowerShell
  2) Collect recycle bin items (first + optional second stage)
  3) Filter items that belong to a target folder path fragment
  4) Export EVERYTHING that would be restored to CSV (full list + DirName summary)
  5) Optionally restore:
       - a single TEST folder first (recommended)
       - then the full scope

NEW FEATURE: -IgnoreDeletedByEmail "email@domain.com"
  Skips items deleted by a specific user (useful for migration cleanup).

NEW FEATURE: -RestoreListFile "path\to\file.txt"
  One path/fragment per line. Overrides folder filtering when provided.

IMPORTANT - Authentication (September 2024+):
  Register your own Entra ID app; use -ClientId.
  See: https://pnp.github.io/powershell/articles/registerapplication.html
  Permissions: Delegated AllSites.FullControl or similar + admin consent.

KNOWN LIMITATION - Large Recycle Bin:
  With >50k-100k items in the bin, Restore-PnPRecycleBinItem often fails with:
  "The attempted operation is prohibited because it exceeds the list view threshold"
  Updated to use REST API for bulk restore to attempt bypass.

Safety defaults:
  - Action defaults to Preview (no restore)
  - Exports always happen before restore
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$SiteUrl,

  [Parameter(Mandatory = $true)]
  [string]$ClientId,

  [Parameter(Mandatory = $false)]  # ← Changed to $false so list mode works without it
  [string]$RootFolderPathFragment,

  [Parameter(Mandatory = $false)]
  [string]$TestFolderPathFragment,

  [ValidateSet("Preview", "RestoreTestFolder", "RestoreAll")]
  [string]$Action = "Preview",

  [string]$OutputDir = ".\restore-output",

  [int]$RowLimit = 50000,

  [int]$BatchSize = 100,  # Lowered default for large bins / threshold safety

  [bool]$IncludeSecondStage = $true,

  [bool]$RequireTypedConfirmation = $true,

  [string]$IgnoreDeletedByEmail = $null,

  # NEW: Optional list file (overrides folder mode when provided)
  [string]$RestoreListFile = $null
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Ensure-OutputDir {
  param([string]$Path)
  if (-not (Test-Path $Path)) {
    New-Item -ItemType Directory -Path $Path | Out-Null
  }
}

function Timestamp {
  return (Get-Date -Format "yyyyMMdd-HHmmss")
}

function Require-Module {
  param([string]$Name)
  if (-not (Get-Module -ListAvailable -Name $Name)) {
    throw "Missing module '$Name'. Install it: Install-Module $Name -Scope CurrentUser -AllowClobber"
  }
  Write-Host "Tip: Keep PnP.PowerShell updated: Update-Module PnP.PowerShell -Force" -ForegroundColor Yellow
}

function Connect-SharePoint {
  param([string]$Url, [string]$ClientId)
  Write-Host "Connecting to SharePoint (interactive login)..." -ForegroundColor Cyan
  Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId
}

function Get-AllRecycleBinItems {
  param([int]$Limit, [bool]$IncludeSecond)
  Write-Host "Fetching recycle bin items (First stage) RowLimit=$Limit ..." -ForegroundColor Cyan
  $first = @(Get-PnPRecycleBinItem -FirstStage -RowLimit $Limit)

  $second = @()
  if ($IncludeSecond) {
    Write-Host "Fetching recycle bin items (Second stage) RowLimit=$Limit ..." -ForegroundColor Cyan
    $second = @(Get-PnPRecycleBinItem -SecondStage -RowLimit $Limit)
  }
  return @($first + $second)
}

function Split-FolderParts {
  param([string]$PathFragment)
  $clean = $PathFragment.Trim().TrimEnd("/")
  $parts = $clean -split "/"
  $leaf = $parts[-1]
  $parent = ""
  if ($parts.Count -gt 1) {
    $parent = ($parts[0..($parts.Count - 2)] -join "/")
  }
  return [PSCustomObject]@{
    CleanPath = $clean
    LeafName  = $leaf
    Parent    = $parent
  }
}

function Get-ScopedItemsForPath {
  param(
    [array]$AllItems,
    [string]$PathFragment
  )
  $p = Split-FolderParts -PathFragment $PathFragment
  $scoped = $AllItems | Where-Object {
    ($_.DirName -like "*$($p.CleanPath)*") -or
    (($_.ItemType -eq "Folder") -and ($_.LeafName -eq $p.LeafName) -and ($_.DirName -like "*$($p.Parent)*"))
  }
  return @($scoped)
}

function Export-ScopeCSVs {
  param(
    [array]$ScopeItems,
    [string]$BaseName,
    [string]$OutDir
  )
  $ts = Timestamp
  $fullPath    = Join-Path $OutDir "$BaseName-full-$ts.csv"
  $summaryPath = Join-Path $OutDir "$BaseName-dirname-summary-$ts.csv"

  $ScopeItems |
    Select-Object Id, ItemType, LeafName, DirName, DeletedDate, DeletedByName, DeletedByEmail |
    Export-Csv -Path $fullPath -NoTypeInformation -Encoding UTF8

  $ScopeItems |
    Group-Object DirName |
    Select-Object @{n="DirName";e={$_.Name}}, Count |
    Sort-Object Count -Descending |
    Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8

  Write-Host "Exported full list:      $fullPath" -ForegroundColor Green
  Write-Host "Exported DirName summary: $summaryPath" -ForegroundColor Green
}

function Confirm-Typed {
  param([string]$Phrase)
  $typed = Read-Host "Type exactly '$Phrase' to proceed"
  if ($typed -ne $Phrase) {
    throw "Confirmation phrase did not match. Aborting restore."
  }
}

function Restore-InBatches {
  param(
    [array]$Items,
    [int]$Batch,
    [string]$OutDir,
    [string]$LogBaseName
  )
  $ts = Timestamp
  $logPath = Join-Path $OutDir "$LogBaseName-restore-log-$ts.csv"
  $logRows = New-Object System.Collections.Generic.List[object]

  if ($null -eq $Items -or $Items.Count -eq 0) {
    Write-Host "Nothing to restore (0 items)." -ForegroundColor Yellow
    return
  }

  Write-Host "Starting restore: $($Items.Count) items, batchSize=$Batch" -ForegroundColor Cyan
  Write-Host "Using REST API for bulk restore to attempt bypassing threshold issues." -ForegroundColor Yellow

  $siteUrl = (Get-PnPContext).Url
  $apiUrl = "$siteUrl/_api/site/RecycleBin/RestoreByIds"

  for ($i = 0; $i -lt $Items.Count; $i += $Batch) {
    $end = [Math]::Min($i + $Batch - 1, $Items.Count - 1)
    $batchItems = $Items[$i..$end]
    Write-Host "Restoring batch $i..$end ($($batchItems.Count) items)..." -ForegroundColor Cyan

    $ids = $batchItems | ForEach-Object { "`"$($_.Id)`"" } | Join-String -Separator ","
    $body = "{`"ids`":[$ids]}"

    try {
      Invoke-PnPSPRestMethod -Method Post -Url $apiUrl -Content $body -ContentType "application/json;odata=verbose"
      foreach ($it in $batchItems) {
        $logRows.Add([PSCustomObject]@{
          Status      = "Restored"
          Id          = $it.Id
          ItemType    = $it.ItemType
          LeafName    = $it.LeafName
          DirName     = $it.DirName
          DeletedDate = $it.DeletedDate
          Error       = ""
        }) | Out-Null
      }
    }
    catch {
      $err = $_.Exception.Message
      Write-Warning "Batch restore failed. Falling back to per-item REST. Error: $err"
      foreach ($it in $batchItems) {
        try {
          $singleBody = "{`"ids`":[`"$($it.Id)`"]}"
          Invoke-PnPSPRestMethod -Method Post -Url $apiUrl -Content $singleBody -ContentType "application/json;odata=verbose"
          $logRows.Add([PSCustomObject]@{
            Status      = "Restored"
            Id          = $it.Id
            ItemType    = $it.ItemType
            LeafName    = $it.LeafName
            DirName     = $it.DirName
            DeletedDate = $it.DeletedDate
            Error       = ""
          }) | Out-Null
        }
        catch {
          $logRows.Add([PSCustomObject]@{
            Status      = "Failed"
            Id          = $it.Id
            ItemType    = $it.ItemType
            LeafName    = $it.LeafName
            DirName     = $it.DirName
            DeletedDate = $it.DeletedDate
            Error       = $_.Exception.Message
          }) | Out-Null
        }
      }
    }
  }

  $logRows | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8
  Write-Host "Restore log written: $logPath" -ForegroundColor Green
}

# -------------------- MAIN --------------------

Ensure-OutputDir -Path $OutputDir
Require-Module -Name "PnP.PowerShell"

$transcriptPath = Join-Path $OutputDir ("transcript-" + (Timestamp) + ".txt")
Start-Transcript -Path $transcriptPath -Append | Out-Null

try {
  Connect-SharePoint -Url $SiteUrl -ClientId $ClientId

  $all = Get-AllRecycleBinItems -Limit $RowLimit -IncludeSecond $IncludeSecondStage
  Write-Host "Total recycle bin items retrieved: $($all.Count)" -ForegroundColor Cyan

  # Ignore items deleted by specific user
  if ($IgnoreDeletedByEmail) {
    $originalCount = $all.Count
    $all = $all | Where-Object { $_.DeletedByEmail -ne $IgnoreDeletedByEmail }
    $ignoredCount = $originalCount - $all.Count
    Write-Host "Ignored $ignoredCount items deleted by '$IgnoreDeletedByEmail'" -ForegroundColor Yellow
  }

  [array]$targetItems = @()
  $scopeName = "scope"

  # ──────────────────────────────
  # Choose mode
  # ──────────────────────────────
  if ($RestoreListFile) {
    # List mode (overrides folder mode)
    if (-not (Test-Path $RestoreListFile -PathType Leaf)) {
      throw "File not found: $RestoreListFile"
    }

    Write-Host "Using list mode from file: $RestoreListFile" -ForegroundColor Cyan
    $patterns = Get-Content $RestoreListFile | ForEach-Object { $_.Trim() } | Where-Object { $_ }

    foreach ($pattern in $patterns) {
      $matches = $all | Where-Object {
        $_.DirName -like "*$pattern*" -or
        "$($_.DirName)/$($_.LeafName)" -like "*$pattern*"
      }
      $targetItems += $matches
    }

    $targetItems = $targetItems | Sort-Object Id -Unique
    $scopeName = "list-scope"

    if ($targetItems.Count -eq 0) {
      throw "No items matched any line in the file. Check paths in $RestoreListFile against your preview CSV (DirName + LeafName)."
    }

    Write-Host "Found $($targetItems.Count) matching items from file" -ForegroundColor Green
  }
  else {
    # Folder mode (original)
    if (-not $RootFolderPathFragment) {
      throw "Folder mode requires -RootFolderPathFragment (or use -RestoreListFile for list mode)."
    }

    $rootScope = Get-ScopedItemsForPath -AllItems $all -PathFragment $RootFolderPathFragment
    Write-Host "Items in ROOT scope ('$RootFolderPathFragment'): $(if ($rootScope) { $rootScope.Count } else { 0 })" -ForegroundColor Cyan

    if ($Action -eq "RestoreTestFolder") {
      if (-not $TestFolderPathFragment) {
        throw "Action=RestoreTestFolder requires -TestFolderPathFragment."
      }
      $targetItems = Get-ScopedItemsForPath -AllItems $rootScope -PathFragment $TestFolderPathFragment
      $scopeName = "test-scope"
    }
    else {
      $targetItems = $rootScope
      $scopeName = "root-scope"
    }
  }

  Export-ScopeCSVs -ScopeItems $targetItems -BaseName $scopeName -OutDir $OutputDir

  if ($Action -eq "Preview") {
    Write-Host "Action=Preview. No restore performed." -ForegroundColor Yellow
    Write-Host "Review the CSVs in: $OutputDir" -ForegroundColor Yellow
    return
  }

  if ($targetItems.Count -eq 0) {
    Write-Host "No items to restore after filtering. Aborting." -ForegroundColor Yellow
    return
  }

  # Confirmation
  $phrase = if ($Action -eq "RestoreTestFolder") { "RESTORE-TEST" } else { "RESTORE-ALL" }
  if ($RequireTypedConfirmation) {
    Confirm-Typed -Phrase $phrase
  }

  Restore-InBatches -Items $targetItems -Batch $BatchSize -OutDir $OutputDir -LogBaseName $scopeName

  Write-Host "Restore complete. Check log: $scopeName-restore-log-*.csv" -ForegroundColor Green

}
finally {
  Stop-Transcript | Out-Null
  Write-Host "Transcript written: $transcriptPath" -ForegroundColor Green
} 
