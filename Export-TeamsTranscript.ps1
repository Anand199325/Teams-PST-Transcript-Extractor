# ==============================================================================
# Script  : Export-TeamsTranscript.ps1
# Purpose : Extract Microsoft Teams chat messages from an eDiscovery PST file
#           and generate a clean, readable transcript (.txt)
# Author  : S Anand Rao | M365 Nexus
# Version : 1.0
# Date    : April 2026
# ==============================================================================
# PRE-REQUISITES:
#   - Microsoft Outlook (desktop) must be installed and configured
#   - PST file must be accessible at the path defined in $pstPath
#   - Run as the logged-in user (Outlook COM requires interactive session)
# ==============================================================================

# ==== CONFIG ====
$pstPath    = "C:\Temp\test@test.001.pst"   # Path to the eDiscovery PST file
$outputFile = "C:\Temp\TeamsTranscript.txt"      # Output transcript file path

# ==============================================================================
# SECTION 1 — INITIALIZE OUTLOOK & MOUNT PST
# ==============================================================================

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  Teams Chat Transcript Generator - M365 Nexus" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "[INFO] Starting Outlook COM session..." -ForegroundColor Yellow

try {
    $outlook   = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
} catch {
    Write-Host "[ERROR] Failed to launch Outlook. Ensure Outlook is installed and not already running in a locked state." -ForegroundColor Red
    exit 1
}

# Mount the PST file
Write-Host "[INFO] Mounting PST: $pstPath" -ForegroundColor Yellow

try {
    $namespace.AddStore($pstPath)
} catch {
    Write-Host "[ERROR] Could not mount PST. Verify the file path and that the PST is not password-protected." -ForegroundColor Red
    exit 1
}

# Locate the mounted PST store
$pstStore = $namespace.Stores | Where-Object { $_.FilePath -eq $pstPath }

if (-not $pstStore) {
    Write-Host "[ERROR] PST store not found after mounting. Exiting." -ForegroundColor Red
    exit 1
}

$root = $pstStore.GetRootFolder()
Write-Host "[INFO] PST mounted successfully. Root folder: $($root.Name)" -ForegroundColor Green

# ==============================================================================
# SECTION 2 — RECURSIVE FOLDER SEARCH
# ==============================================================================

function Get-FolderByName {
    param (
        [object]$Folder,
        [string]$TargetName
    )
    if ($Folder.Name -like "*$TargetName*") {
        return $Folder
    }
    foreach ($subFolder in $Folder.Folders) {
        $found = Get-FolderByName -Folder $subFolder -TargetName $TargetName
        if ($found) { return $found }
    }
    return $null
}

Write-Host "[INFO] Searching for TeamsMessagesData folder..." -ForegroundColor Yellow
$teamsFolder = Get-FolderByName -Folder $root -TargetName "TeamsMessagesData"

if (-not $teamsFolder) {
    Write-Host "[ERROR] TeamsMessagesData folder not found in PST. The PST may not contain Teams export data." -ForegroundColor Red
    exit 1
}

Write-Host "[INFO] Found folder: $($teamsFolder.Name) | Items: $($teamsFolder.Items.Count)" -ForegroundColor Green

# ==============================================================================
# SECTION 3 — PROCESS & EXTRACT MESSAGES
# ==============================================================================

Write-Host "[INFO] Processing messages..." -ForegroundColor Yellow

$result       = @()
$successCount = 0
$skipCount    = 0
$errorCount   = 0

# Sort messages by CreationTime for chronological order
$messages = $teamsFolder.Items | Sort-Object CreationTime

foreach ($msg in $messages) {
    try {
        # --- SENDER ---
        $sender = if ($msg.SenderName) { $msg.SenderName } else { "Unknown Sender" }

        # --- BEST TIMESTAMP LOGIC ---
        # Priority: CreationTime > SentOn > ReceivedTime
        $time = $null

        if ($msg.PSObject.Properties["CreationTime"] -and $msg.CreationTime -and
            $msg.CreationTime -ne [DateTime]::MinValue) {
            $time = $msg.CreationTime
        } elseif ($msg.PSObject.Properties["SentOn"] -and $msg.SentOn -and
                  $msg.SentOn -ne [DateTime]::MinValue) {
            $time = $msg.SentOn
        } elseif ($msg.PSObject.Properties["ReceivedTime"] -and $msg.ReceivedTime -and
                  $msg.ReceivedTime -ne [DateTime]::MinValue) {
            $time = $msg.ReceivedTime
        }

        $timeStr = if ($time) {
            $time.ToString("dd MMM yyyy, hh:mm tt")
        } else {
            "No Date"
        }

        # --- CLEAN MESSAGE BODY ---
        $body = $msg.Body

        $cleanBody = $body `
            -replace "(?m)^From:.*$",    "" `
            -replace "(?m)^Sent:.*$",    "" `
            -replace "(?m)^To:.*$",      "" `
            -replace "(?m)^Subject:.*$", "" `
            -replace "(?m)^CC:.*$",      "" `
            -replace "`r`n",             " " `
            -replace "`n",               " " `
            -replace "\s{2,}",           " "

        $cleanBody = $cleanBody.Trim()

        # Skip empty messages
        if ([string]::IsNullOrWhiteSpace($cleanBody)) {
            $skipCount++
            continue
        }

        # --- FORMAT LINE ---
        $line      = "[{0}] {1}: {2}" -f $timeStr, $sender, $cleanBody
        $result   += $line
        $successCount++

    } catch {
        $errorCount++
        continue
    }
}

# ==============================================================================
# SECTION 4 — WRITE OUTPUT FILE
# ==============================================================================

Write-Host "[INFO] Writing transcript to: $outputFile" -ForegroundColor Yellow

$header = @(
    "============================================================",
    "  TEAMS CHAT TRANSCRIPT - eDiscovery Export",
    "  Generated by : Export-TeamsTranscript.ps1 | M365 Nexus",
    "  Generated on : $(Get-Date -Format 'dd MMM yyyy, hh:mm tt')",
    "  PST Source   : $pstPath",
    "  Total Messages Extracted: $successCount",
    "============================================================",
    ""
)

$fullOutput = $header + $result

try {
    $fullOutput | Out-File -FilePath $outputFile -Encoding utf8 -Force
} catch {
    Write-Host "[ERROR] Failed to write output file: $_" -ForegroundColor Red
    exit 1
}

# ==============================================================================
# SECTION 5 — SUMMARY
# ==============================================================================

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  TRANSCRIPT GENERATION COMPLETE" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  Messages Extracted : $successCount" -ForegroundColor White
Write-Host "  Messages Skipped   : $skipCount (empty body)" -ForegroundColor Yellow
Write-Host "  Errors Encountered : $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "White" })
Write-Host "  Output File        : $outputFile" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""
