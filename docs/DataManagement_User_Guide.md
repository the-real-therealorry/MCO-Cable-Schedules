# Data Management - User Guide

## Overview

The Data Management system allows you to safely export and import cable and endpoint data between different versions of the Cable Register spreadsheets.

**Use this when**:
- Upgrading to a new spreadsheet version
- Creating backups of your cable data
- Transferring cables between spreadsheets
- Recovering from data loss

---

## Accessing Data Management

1. Open your Cable Register spreadsheet
2. Go to the **Dashboard** sheet
3. Click the **"Data Management"** button

The Data Management form will open.

---

## Export Data

### What is Export?

Export creates files on your computer containing all your cable and endpoint data. These files can be:
- Imported into a new version of the spreadsheet
- Used as backups
- Shared with others
- Archived for records

### Step-by-Step: Export Your Data

1. **Click "Export to CSV"** or **"Export to JSON"** button

2. **Choose what to export**:
   - If prompted, select scope:
     - "All Plants" - Exports everything (recommended for backups)
     - "Wet Plant Only" - Just Wet Screen Crushing Plant
     - "Ore Sorter Only" - Just Ore Sorting Plant
     - "Retreatment Only" - Just Retreatment Gravity Plant

3. **Choose where to save**:
   - A file browser will open
   - Navigate to where you want to save (e.g., Desktop, Documents)
   - The filename is pre-filled with date/time (don't change it)
   - Click **Save**

4. **Wait for confirmation**:
   - You'll see a message: "Export complete!"
   - The message shows where files were saved

5. **Done!** Your data is now safely exported

### Export File Types

#### CSV Files (Comma-Separated Values)

**When to use**:
- You want to view/edit data in Excel or Google Sheets
- Simple backup
- Sharing with non-technical users

**What you get**:
- `Cables_Export_YYYYMMDD_HHMMSS.csv` - All cable records
- `Endpoints_Export_YYYYMMDD_HHMMSS.csv` - All endpoint definitions

**Format**: Simple text file that opens in Excel
- Each row is a cable or endpoint
- Each column is a field (Cable ID, Source, etc.)
- Human-readable and editable

#### JSON Files (JavaScript Object Notation)

**When to use**:
- Complete backup including metadata
- Professional data management
- Automated processes

**What you get**:
- `CableRegister_Export_YYYYMMDD_HHMMSS.json` - Everything in one file
- Includes: cables, endpoints, version info, export date

**Format**: Structured text file
- Contains all data in organized format
- Includes version information for compatibility checking
- Not meant for manual editing (use CSV for that)

---

## Import Data

### What is Import?

Import brings cable and endpoint data FROM a file INTO your spreadsheet. This is how you:
- Restore from backups
- Upgrade to new spreadsheet versions
- Merge data from different sources

### ⚠️ IMPORTANT: Before You Import

**IMPORTING CAN CHANGE OR DELETE YOUR DATA!**

**Best Practices**:
1. ✅ **Always export your current data first** (as a safety backup)
2. ✅ **Test on a copy** of your spreadsheet before doing it for real
3. ✅ **Understand which import mode you need** (see below)
4. ✅ **Check the compatibility report** after import
5. ✅ **Verify your data** looks correct after import

### Import Modes Explained

You must choose HOW to import. Think carefully about which mode fits your situation:

#### APPEND Mode ➕
**What it does**: Adds imported cables to what you already have

**Example**:
- You have: 10 cables in spreadsheet
- You import: 5 cables from file
- Result: 15 cables total (10 old + 5 new)

**When to use**:
- Adding new cables from another source
- Merging data from two different spreadsheets
- Adding cables to empty spreadsheet

**Watch out for**:
- ⚠️ Can create duplicates if Cable IDs already exist
- ⚠️ Doesn't update existing cables
- ⚠️ Everything gets added, even if it's already there

**Best for**: Empty spreadsheets or when you KNOW there are no duplicates

---

#### REPLACE Mode 🗑️➕
**What it does**: DELETES everything, then imports from file

**Example**:
- You have: 10 cables in spreadsheet
- You import: 5 cables from file
- Result: 5 cables total (10 deleted, 5 imported)

**When to use**:
- Fresh installation of new spreadsheet version
- Complete restore from backup
- Starting over with clean data

**⚠️ DANGER**:
- **DESTRUCTIVE**: All existing data is permanently deleted
- Only the cables in the import file will remain
- Cannot undo (except from backup)

**Safety Features**:
- ✅ Automatic backup created before deletion
- ✅ Confirmation prompt before proceeding

**Best for**: New spreadsheet or complete data restoration

---

#### MERGE Mode 🔄
**What it does**: Updates existing cables, adds new ones (smart mode)

**Example**:
- You have: Cable A, Cable B, Cable C
- You import: Cable B (updated), Cable D (new)
- Result: Cable A (unchanged), Cable B (updated), Cable C (unchanged), Cable D (added)

**How it works**:
- Matches cables by Cable ID
- If Cable ID exists: Updates that cable with new data
- If Cable ID doesn't exist: Adds as new cable
- Cables NOT in import file: Left unchanged

**When to use**:
- Syncing changes between spreadsheets
- Updating specific cables
- Selective restore from backup

**Best for**: Most import situations - it's the smartest mode

---

### Step-by-Step: Import Data

1. **Click "Import from File"** button

2. **Choose import mode**:
   - A dialog asks: "Import Mode?"
   - Click button for your choice:
     - **YES** = Append (add to existing)
     - **NO** = Replace (delete all, then import)
     - **CANCEL** = Merge (smart update)

   *Read the modes above if unsure!*

3. **Select the file to import**:
   - File browser opens
   - Navigate to your export file
   - Select the CSV or JSON file
   - Click **Open**

4. **Automatic backup**:
   - Message: "Backup created: [filename]"
   - This backup can restore your data if something goes wrong
   - Click **OK**

5. **Import happens**:
   - Screen may appear frozen - this is NORMAL
   - Wait for it to finish (usually a few seconds)
   - Don't click anything or close Excel

6. **Review the results**:
   - A summary appears:
     ```
     Import Complete!

     Cables Imported: 45
     Cables Skipped: 3

     AUTO-FIXES APPLIED:
     1. Created missing endpoint: MCC102 - MCC 2 (Imported - Review)
     2. Created missing endpoint: CV105 - Conveyor belt 5 (Imported - Review)
     ```

7. **Check your data**:
   - Go to the plant sheets (Wet Plant, Ore Sorter, Retreatment)
   - Scroll through and verify cables look correct
   - **Check the Data sheet** for any endpoints marked "(Imported - Review)"

8. **Review imported endpoints** (if any):
   - Go to **Data** sheet
   - Look for endpoints with "(Imported - Review)" in description
   - These were auto-created during import
   - Verify they're correct:
     - Is the short name right? (e.g., MCC102)
     - Is the description right? (e.g., MCC 2)
   - Edit if needed, then remove "(Imported - Review)" marker

---

## Auto-Fix System

### What is Auto-Fix?

When you import cables, sometimes the endpoint descriptions don't exactly match what's in your Endpoints table. Instead of failing, the system **automatically fixes** these issues.

### What Gets Auto-Fixed?

#### Missing Endpoints

**Problem**: Import file has cable with source "Water Pump 5" but this endpoint doesn't exist in your Endpoints table

**Auto-Fix**:
1. Searches for exact match - not found
2. Tries fuzzy match (ignoring spaces/caps) - not found
3. **Creates new endpoint**:
   - Short Name: WP105 (auto-generated)
   - Description: "Water Pump 5 (Imported - Review)"
4. Logs this fix in the import report

**What you need to do**:
- Go to Data sheet
- Find "Water Pump 5 (Imported - Review)"
- Verify it's correct
- If correct: Remove "(Imported - Review)" marker
- If wrong: Edit or delete, then fix cables that use it

#### Fuzzy Matching

**Problem**: Import has "MCC2" but your table has "MCC 2" (different spacing)

**Auto-Fix**:
1. Exact match fails
2. Removes spaces and compares: "mcc2" = "mcc2" ✓
3. Finds existing endpoint: MCC102 - "MCC 2"
4. Uses that instead of creating new one
5. Logs: "Fuzzy matched 'MCC2' → MCC102"

**What you need to do**: Nothing! It found the right match automatically.

### Reading the Auto-Fix Report

After import, the report shows all fixes:

```
AUTO-FIXES APPLIED:
1. Created missing endpoint: WP105 - Water Pump 5 (Imported - Review)
2. Fuzzy matched endpoint: 'MCC2' → MCC102
3. Created missing endpoint: CV110 - Conveyor belt 10 (Imported - Review)
```

**Action required**: Review any created endpoints (items with "Created missing endpoint")

---

## Backups

### Automatic Backups

**Every time you import**, a backup is automatically created BEFORE any changes are made.

**Backup files**:
- Location: `Backups/` folder (same location as your Excel file)
- Format: `Import_Backup_YYYYMMDD_HHMMSS.json`
- Contains: Complete snapshot of data before import
- Auto-cleanup: Only last 10 backups kept, older ones deleted

**You can restore from backup anytime!**

### How to Restore from Backup

If an import goes wrong:

1. Click **"Import from File"**
2. Choose **REPLACE** mode (you want to restore completely)
3. Navigate to `Backups/` folder
4. Select the backup file (check date/time)
5. Import it
6. Your data is restored!

### Manual Backups

**Best practice**: Before major changes, export to JSON as a manual backup:

1. Click **"Export to JSON"**
2. Save with meaningful name: `Backup_Before_Upgrade_Dec2024.json`
3. Store somewhere safe (USB drive, cloud, etc.)

**Keep these backups**:
- Before spreadsheet upgrades
- Before major data changes
- Monthly/quarterly (for records)

---

## Troubleshooting

### "Import failed - file not found"

**Cause**: File path has special characters or file was moved

**Fix**:
- Move file to simple location (e.g., `C:\Temp\`)
- Use simple filename (no special characters)
- Try again

### "Export creates empty files"

**Cause**: No cables in selected plant or table names don't match

**Fix**:
1. Check cables exist in the plant you're exporting
2. If just created fresh spreadsheet, add some test cables first
3. If problem persists, check with developer

### "Endpoint not found" warnings after import

**Cause**: Imported cables reference endpoints that don't exist (auto-fix created them)

**Fix**:
1. Go to **Data** sheet
2. Look for "(Imported - Review)" in endpoint descriptions
3. Review each one:
   - Correct? Remove "(Imported - Review)" marker
   - Wrong? Edit or delete it

### Import appears frozen

**Cause**: Processing large dataset (this is normal!)

**Fix**:
- Wait patiently (up to 1 minute for 1000+ cables)
- Don't click anything
- Don't close Excel
- It will complete

### Duplicate cables after import

**Cause**: Used APPEND mode when cables already existed

**Fix**:
- Future: Use MERGE mode instead of APPEND
- Now: Manually delete duplicates or restore from backup and re-import with MERGE

### Some cables missing after import

**Cause**: Used REPLACE mode (deletes everything not in import file)

**Fix**:
- Restore from backup
- Use MERGE mode instead

---

## Tips & Best Practices

### ✅ DO

- **Export regularly** as backups (weekly/monthly)
- **Test imports on a copy** of spreadsheet first
- **Review auto-fix reports** after import
- **Check "(Imported - Review)" endpoints** after import
- **Use MERGE mode** for most imports (it's smartest)
- **Keep export files** for at least 30 days
- **Name files meaningfully** for manual backups

### ❌ DON'T

- **Don't use REPLACE mode** unless you really want to delete everything
- **Don't ignore "(Imported - Review)"** markers - review them!
- **Don't edit CSV files in Excel** (can corrupt format - use Notepad)
- **Don't delete backup files** immediately after import
- **Don't skip the summary report** - read it!
- **Don't import without testing** on a copy first

---

## Common Scenarios

### Upgrading to New Spreadsheet Version

1. **In OLD version**:
   - Export to both CSV and JSON
   - Save files somewhere safe

2. **In NEW version**:
   - Click Import
   - Choose **REPLACE** mode (fresh start)
   - Select JSON file (includes everything)
   - Review auto-fix report
   - Check "(Imported - Review)" endpoints
   - Verify data

3. **Done!** Old data now in new version

### Backing Up Your Data

1. Click **"Export to JSON"**
2. Save as: `Backup_[Description]_[Date].json`
   - Example: `Backup_BeforeUpgrade_Dec2024.json`
3. Store on USB drive or cloud
4. Keep for records

### Recovering from Mistake

1. Check `Backups/` folder for recent backup
2. Click **"Import from File"**
3. Choose **REPLACE** mode
4. Select backup file
5. Import
6. Data restored!

### Merging Cables from Another Spreadsheet

1. **In other spreadsheet**: Export to CSV
2. **In your spreadsheet**: Import
3. Choose **MERGE** mode (updates existing, adds new)
4. Review results
5. Check for duplicates

---

## File Formats Reference

### CSV Structure

**Cables CSV**:
```
Version,Plant,Scheduled,IDAttached,CableID,Source,Destination,...
2024.12.1,WET_PLANT,FALSE,TRUE,CV102-C1001-CV103,MCC 2,Conveyor belt 3,...
```

**Endpoints CSV**:
```
Version,Plant,ShortName,Description
2024.12.1,WET_PLANT,MCC102,MCC 2
```

### JSON Structure

One file with everything:
- Version information
- Export date and source
- All plants data (endpoints + cables)
- Metadata

---

## Getting Help

If you encounter problems:

1. Check this guide's Troubleshooting section
2. Look in `Backups/` folder for auto-backups
3. Check VBA Immediate Window (Alt+F11, Ctrl+G) for error messages
4. Contact your spreadsheet administrator with:
   - What you were trying to do
   - Which file you were importing
   - Error message (screenshot if possible)
   - Backup file name (from message)

---

## Version History

- **v1.0** (Dec 2024): Initial release
  - CSV export/import
  - JSON export
  - Auto-fix system
  - Automatic backups

---

**Remember**: When in doubt, test on a copy first and always review the import report!
