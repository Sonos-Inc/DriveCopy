# Google Drive Backup & Shared Drive Rotation (PowerShell + GAM)

PowerShell automation for backing up users’ **Google Drive (My Drive)** into a **designated Shared Drive folder**, plus **automatic Shared Drive rotation** when utilization thresholds are exceeded.

This repo is designed for operational use in Google Workspace environments where:
- **GAM** performs the Drive and Sheets operations
- backups are written into a controlled **Shared Drive** structure
- copy operations may require **temporary ACL changes**
- batch runs are driven by a **Google Sheet** (exported to CSV) and produce consistent results files for reporting

---

## What’s Included

### 1) `Copy-SingleDriveToShare.ps1`
Copies a single user’s My Drive into a target folder on a Shared Drive.

**Modes**
- **Single-user mode**: provide `-Email`
- **Batch mode**: omit `-Email` and read `CopyRunEligible.csv` from the current directory

**High-level workflow**
- Validates GAM availability and required modules
- Resolves destination Shared Drive folder
- Creates a per-user backup folder
- Adjusts the user’s state (archive/OU) via `Set-UserCopyState`
- Applies temporary ACLs required for copy
- Performs a full recursive copy with GAM
- Removes temporary ACLs and finalizes user state
- Optional manager access + notification
- Emits a per-user result object
- Single-user mode can optionally update a tracking sheet and send a summary email

---

### 2) `Run-BatchDriveBackups.ps1`
Orchestrates multi-user backups driven by the `"CopyRunEligible"` Google Sheet.

**What it does**
- Downloads `"CopyRunEligible"` from Google Sheets as `CopyRunEligible.csv` (via GAM)
- If eligible users exist, invokes `Copy-SingleDriveToShare.ps1` once
- The child script processes all rows and writes a local results file
- Imports results and updates the sheet + reporting
- Uses **file-based coordination** (no JSON IPC)

---

### 3) `RotateBackup.ps1`
Automates Shared Drive rotation when usage thresholds are exceeded.

**What it does**
- Reads Shared Drive utilization from a tracking Google Sheet
- If usage exceeds configured thresholds:
  1. Creates a new Shared Drive
  2. Updates the tracking sheet with the new entry
  3. Grants organizer ACLs to designated admins
- Sends email alerts for key events and errors via `Send-Alert.psm1`

**Design notes**
- GAM operations are validated via `$LASTEXITCODE`
- Intended to be **idempotent** (safe to re-run)

---

## Requirements

### Runtime
- Windows PowerShell 5.1 or PowerShell 7+
- **GAM** installed and available in `PATH`
- Google Workspace permissions sufficient for:
  - Drive copy operations (including Shared Drive access)
  - Sheets export/import/update
  - Shared Drive creation + ACL management (for rotation)

### Repo Modules / Dependencies
- `Send-Alert.psm1` (required for rotation and alerting)
- `Set-UserCopyState` (used during copy workflows)
- Any supporting modules/scripts referenced by your environment (kept intentionally modular)

> Note: This repo assumes GAM commands are executed and validated using `$LASTEXITCODE`.
