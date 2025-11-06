# Import/Export Feature - Quick Start Guide

## Summary

This feature allows you to transfer cables and endpoints between different versions of the cable register spreadsheets.

**Use Cases**:
- Upgrading to new spreadsheet version
- Backing up cable data
- Migrating between plants
- Data recovery

---

## How It Works

### Export Process

```
Old Spreadsheet
    ↓
[Export Button]
    ↓
CSV/JSON Files Created
    ↓
Save to Disk
```

### Import Process

```
New Spreadsheet
    ↓
[Import Button]
    ↓
Select CSV/JSON Files
    ↓
Auto-Fix Compatibility Issues
    ↓
Import Complete + Report
```

---

## User Workflow

### Scenario: Upgrading to New Version

**You have**: `Power Cable Register Rev0.xlsm` (old, with your data)
**You get**: `Power Cable Register Rev1.xlsm` (new, empty template)

**Steps**:

1. **Open OLD version** (`Rev0.xlsm`)
2. Click **Dashboard** → **Data Management** button
3. Select export scope:
   - "All Plants" (exports everything)
4. Check both formats:
   - ☑ CSV
   - ☑ JSON
5. Click **Export to File...**
6. Save to safe location (e.g., Desktop, Documents)
7. **Close OLD version**

8. **Open NEW version** (`Rev1.xlsm`)
9. Click **Dashboard** → **Data Management** button
10. Select import mode:
    - "Replace" (for fresh install)
    - "Append" (to add to existing)
    - "Merge" (to update existing by Cable ID)
11. Ensure "Auto-Fix" is checked ☑
12. Click **Import from File...**
13. Select the export files you saved
14. Review compatibility report
15. Check log file for details
16. **Verify your data**

**Done!** Your cables are now in the new version.

---

## Export Options Explained

### Scope

**All Plants**
- Exports cables from all 3 plants
- Exports all endpoints
- Best for: full backups, version upgrades

**Single Plant** (Wet/Ore/Retreatment)
- Exports only that plant's data
- Best for: plant-specific transfers

### Formats

**CSV** (Comma-Separated Values)
- ✓ Opens in Excel, Google Sheets
- ✓ Human-readable
- ✓ Simple format
- ✗ Multiple files (cables + endpoints)

**JSON** (JavaScript Object Notation)
- ✓ Single file with everything
- ✓ Includes metadata (version, date)
- ✓ Better for automation
- ✗ Not human-friendly

**Recommendation**: Export BOTH for safety

---

## Import Modes Explained

### Append

**What it does**: Adds imported cables to existing ones

**Use when**:
- Merging data from two sources
- Adding plant data to empty spreadsheet
- Combining backups

**Watch out**:
- Can create duplicates if Cable IDs match
- No overwriting of existing data

### Replace

**What it does**: Deletes all existing cables, then imports

**Use when**:
- Fresh installation
- Restoring from backup
- Clean migration

**Watch out**:
- **DESTRUCTIVE!** All current data is deleted
- Backup is created automatically

### Merge

**What it does**: Updates cables by Cable ID, adds new ones

**Use when**:
- Updating specific cables
- Syncing between versions
- Selective restore

**How it works**:
- Matches by Cable ID
- If ID exists: update that cable
- If ID doesn't exist: add as new

---

## Auto-Fix Compatibility

### What Gets Auto-Fixed

✅ **Missing Endpoints**
- Creates endpoints that don't exist in new version
- Marked for review in report

✅ **Column Name Changes**
- Maps old column names to new ones
- E.g., "Insulation" → "InsulationType"

✅ **Data Format Differences**
- Converts "45m" → "45"
- Trims whitespace
- Normalizes Boolean values

✅ **Plant ID Variations**
- "WETPLANT" → "WET_PLANT"
- "Ore-Sorter" → "ORE_SORTER"

### What Doesn't Get Auto-Fixed

⚠️ **Corrupt data** - Skipped with error

⚠️ **Invalid Cable IDs** - Imported as-is, flagged

⚠️ **Completely incompatible formats** - Import fails

### Compatibility Report

After import with auto-fixes, you'll see:

```
45 cables imported successfully
3 auto-fixes applied
2 warnings

See log file for details:
C:\...\Logs\Import_Log_20241221_103000.txt
```

**Always review** the log file after import!

---

## Backup & Recovery

### Automatic Backups

Before EVERY import, a backup is created:
- Location: `Backups/Import_Backup_YYYYMMDD_HHMMSS.json`
- Contains: Full state before import
- Use for: Rollback if needed

### Manual Backups

**Best practice**:
1. Before major changes, export to JSON
2. Name meaningfully: `Backup_BeforeUpgrade_20241221.json`
3. Store outside Excel folder (cloud, USB, etc.)

### Rollback

If import goes wrong:
1. Note the backup file name from import dialog
2. Use Import function
3. Select "Replace" mode
4. Choose the backup file
5. All data restored!

---

## File Naming Convention

Exports are automatically named with timestamp:

```
Cables_Export_20241221_103045.csv
Endpoints_Export_20241221_103045.csv
CableRegister_Export_20241221_103045.json
```

**Format**: `{Type}_Export_{YYYYMMDD}_{HHMMSS}.{ext}`

This prevents accidental overwrites!

---

## Troubleshooting

### Problem: "Import failed - file not found"

**Solution**: Ensure file path has no special characters or spaces. Move file to simple location like `C:\Temp\`.

### Problem: "Endpoint not found" warnings

**Solution**:
- Auto-fix created the endpoint in Data sheet
- Review Data sheet endpoints table
- Verify descriptions match expectations
- Edit if needed

### Problem: "Version incompatible"

**Solution**:
- Check auto-fix is enabled
- If still fails, export from old version again with latest template
- Contact support with log file

### Problem: Duplicate cables after append

**Solution**:
- Use "Merge" mode instead
- Or manually delete duplicates
- Or use "Replace" for clean start

### Problem: Slow import (large datasets)

**Solution**:
- Normal for 500+ cables
- Progress bar shows status
- Don't interrupt process
- Consider importing single plants

---

## Best Practices

### ✅ DO

- **Export both CSV and JSON** for safety
- **Test import on copy** before production
- **Review compatibility report** after import
- **Keep backups** for at least 30 days
- **Verify data** after import
- **Document custom changes** in log

### ❌ DON'T

- **Don't skip backups** - always create before import
- **Don't ignore warnings** - review compatibility report
- **Don't rush** - verify data after import
- **Don't delete export files** immediately
- **Don't modify CSV files in Excel** (can corrupt format)

---

## Log Files

### Location

`Logs/Import_Log_YYYYMMDD_HHMMSS.txt`

### Contents

```
=== IMPORT LOG ===
Date: 2024-12-21 10:30:45
Source File: CableRegister_Export_20241220_150000.json
Target: Power Cable Register Rev1.xlsm
Mode: REPLACE
Auto-Fix: ENABLED

PROCESSING:
[10:30:46] Reading import file...
[10:30:46] Detected version: 1.0 (current: 1.1)
[10:30:46] Compatibility check: NEEDS_FIX
[10:30:47] Processing WET_PLANT endpoints... 31 found
[10:30:47] Processing WET_PLANT cables... 45 found
[10:30:48] Auto-fix: Created endpoint MCC102 (was missing)
[10:30:48] Auto-fix: Mapped column "Insulation" → "InsulationType"
[10:30:49] Import complete

RESULTS:
✓ Cables imported: 45
✓ Endpoints imported: 31
⚠ Auto-fixes applied: 2
✗ Errors: 0

See compatibility report for details.
```

**Keep logs** for troubleshooting!

---

## FAQ

**Q: Can I edit CSV files in Excel?**

A: Not recommended - Excel can corrupt CSV format. Use Notepad or proper CSV editor.

**Q: What if I import the same file twice?**

A: Depends on mode:
- Append: Creates duplicates
- Replace: Same result
- Merge: Updates existing, no duplicates

**Q: Can I import from older version (e.g., v0.5)?**

A: Yes! Auto-fix handles version differences.

**Q: What if auto-fix creates wrong endpoint?**

A: Edit in Data sheet, then update cable references.

**Q: Can I export just 5 specific cables?**

A: Not in v1. Use export all, then edit CSV to remove unwanted.

**Q: Will import overwrite endpoints?**

A: Only in Replace mode. Append/Merge preserve existing endpoints.

---

## Support

If you encounter issues:

1. Check log file in `Logs/` folder
2. Check compatibility report
3. Review this guide's Troubleshooting section
4. Contact developer with:
   - Log file
   - Export file (if possible)
   - Steps to reproduce
   - Error message screenshot

---

## Version History

- **v1.0** (2024-12): Initial implementation
  - CSV export/import
  - JSON export/import
  - Basic auto-fix
  - Compatibility reporting

---

Ready to start? Open your spreadsheet and look for the **Data Management** button on the Dashboard!
