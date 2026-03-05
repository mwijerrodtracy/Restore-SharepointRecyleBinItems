# SharePoint Online Recycle Bin Restore Tool

A robust PowerShell utility designed to bulk-restore items from the SharePoint Online Recycle Bin. This script specifically addresses the "List View Threshold" error (items > 50k) by utilizing the SharePoint REST API for batch processing instead of standard PnP cmdlets.

## 🚀 Key Features

* **Threshold Bypass:** Uses RestoreByIds via the REST API to handle large recycle bins where standard PnP cmdlets typically fail.
* **Dual-Stage Support:** Optionally includes both First Stage (User) and Second Stage (Site Collection) recycle bins.
* **Flexible Filtering:** * Filter by folder path fragments.
    * Ignore items deleted by a specific user (useful for filtering migration cleanup).
    * **List Mode:** Provide a .txt file with specific paths to target only those items.
* **Safety First:** Defaults to Preview Mode. Generates CSV reports of exactly what will be restored before any action is taken.
* **Trial Runs:** Supports a RestoreTestFolder action to verify results on a small scale before a full recovery.

---

## 📋 Prerequisites

1. PnP PowerShell Module:
   `Install-Module PnP.PowerShell -Scope CurrentUser`

2. Entra ID App Registration:
   Due to SharePoint's security defaults (post-September 2024), you must use your own Client ID.
   * Permissions: Delegated 'AllSites.FullControl' or equivalent.
   * Guide: https://pnp.github.io/powershell/articles/registerapplication.html

---

## ⚙️ Parameters

| Parameter | Mandatory | Description |
| :--- | :---: | :--- |
| -SiteUrl | Yes | The full URL of the SharePoint Site Collection. |
| -ClientId | Yes | Your Entra ID (Azure AD) Application Client ID. |
| -RootFolderPathFragment | No* | The folder path to target. *Required if not using List Mode.* |
| -Action | No | Preview (Default), RestoreTestFolder, or RestoreAll. |
| -RestoreListFile | No | Path to a .txt file containing specific paths (one per line). |
| -IgnoreDeletedByEmail | No | Skip items deleted by a specific user. |
| -TestFolderPathFragment | No | A sub-folder used for a small-scale test restore. |
| -BatchSize | No | Items per REST call (Default: 100). |

---

## 📖 Usage Examples

### 1. Preview Mode (Recommended First Step)
.\Restore-SharePointRecycleBinFolder.ps1 -SiteUrl "https://site.sharepoint.com" -ClientId "GUID" -RootFolderPathFragment "Marketing/2023" -Action Preview

### 2. Restore Using a List File
.\Restore-SharePointRecycleBinFolder.ps1 -SiteUrl "https://site.sharepoint.com" -ClientId "GUID" -RestoreListFile "C:\Temp\Paths.txt" -Action RestoreAll

### 3. Full Restore Excluding a Migration Account
.\Restore-SharePointRecycleBinFolder.ps1 -SiteUrl "https://site.sharepoint.com" -ClientId "GUID" -RootFolderPathFragment "Shared Documents" -IgnoreDeletedByEmail "migrator@contoso.com" -Action RestoreAll

---

## 🛠 Logic & Workflow

1. Connection: Authenticates via Interactive Login using your provided Client ID.
2. Collection: Pulls items from the bin (up to the $RowLimit).
3. Filtering: Matches items against the path fragment or the provided text file.
4. Reporting: Creates a summary.csv and a full.csv before any restore begins.
5. Restoration:
   * Requires a typed confirmation (RESTORE-ALL or RESTORE-TEST).
   * Uses the REST endpoint /_api/site/RecycleBin/RestoreByIds.
   * Fallback: If a batch fails, the script attempts to restore items one-by-one.

---

## ⚠️ Important Notes

* Thresholds: While the REST API bypasses the 50k restore limit, the initial "Fetch" can still be slow if the bin is extremely large.
* Logs: Every execution generates a PowerShell Transcript and a detailed Restore Log CSV in the output directory.