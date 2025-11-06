# VBA Import/Export System - Code Documentation

## Overview

This folder contains the complete VBA code for the cable import/export system.

## Files in This Folder

### Core Modules

1. **modImportExport.bas** (Main Module)
   - `ExportCablesToCSV()` - Export cables to CSV format
   - `ExportEndpointsToCSV()` - Export endpoints to CSV format
   - `ExportToJSON()` - Export combined data to JSON
   - Helper functions for CSV/JSON formatting

2. **modImportExport_Import.bas** (Import Functions - Merge into modImportExport)
   - `ImportCablesFromCSV()` - Import cables from CSV
   - `ImportEndpointsFromCSV()` - Import endpoints from CSV
   - `ClearPlantCables()` - Clear cables for REPLACE mode
   - `ClearPlantEndpoints()` - Clear endpoints for REPLACE mode
   - CSV parsing and row population functions

3. **modCompatibilityFix.bas** (Auto-Fix System)
   - `FixMissingEndpoint()` - Auto-create missing endpoints
   - `NormalizePlantID()` - Handle plant ID variations
   - `ConvertDataType()` - Auto-convert data types
   - Fuzzy matching for endpoint names
   - Auto-fix reporting

4. **modBackup.bas** (Backup Management)
   - `CreateBackup()` - Create pre-import backup
   - `CleanupOldBackups()` - Keep last 10 backups
   - `ListBackups()` - Get available backups

## Quick Start

### Installation

1. Import all `.bas` files into your Excel VBA project
2. Merge `modImportExport_Import.bas` into `modImportExport` module
3. See `docs/IMPORT_EXPORT_Installation_Guide.md` for detailed steps

### Basic Usage

```vba
' Export all cables
modImportExport.ExportCablesToCSV "ALL", "C:\Temp\cables.csv"
modImportExport.ExportEndpointsToCSV "ALL", "C:\Temp\endpoints.csv"

' Export to JSON (includes cables and endpoints)
modImportExport.ExportToJSON "ALL", "C:\Temp\export.json"

' Import with backup and auto-fix
modBackup.CreateBackup "ALL"
modCompatibilityFix.InitializeAutoFix
Set results = modImportExport.ImportCablesFromCSV("C:\Temp\cables.csv", "MERGE")
```

## Function Reference

### Export Functions

#### ExportCablesToCSV(strPlantID, strFilePath)
**Purpose**: Export cables to CSV file
**Parameters**:
- `strPlantID`: "WET_PLANT", "ORE_SORTER", "RETREATMENT", or "ALL"
- `strFilePath`: Full path where CSV should be saved
**Returns**: Boolean (True if successful)

#### ExportEndpointsToCSV(strPlantID, strFilePath)
**Purpose**: Export endpoints to CSV file
**Parameters**: Same as ExportCablesToCSV
**Returns**: Boolean

#### ExportToJSON(strPlantID, strFilePath)
**Purpose**: Export cables and endpoints to JSON file
**Parameters**: Same as above
**Returns**: Boolean

### Import Functions

#### ImportCablesFromCSV(strFilePath, importMode)
**Purpose**: Import cables from CSV file
**Parameters**:
- `strFilePath`: Full path to CSV file
- `importMode`: "APPEND", "REPLACE", or "MERGE"
**Returns**: Dictionary with results:
- `Success`: Boolean
- `CablesImported`: Long
- `CablesSkipped`: Long
- `Errors`: Collection

#### ImportEndpointsFromCSV(strFilePath, importMode)
**Purpose**: Import endpoints from CSV file
**Parameters**: Same as ImportCablesFromCSV
**Returns**: Dictionary with results

### Auto-Fix Functions

#### FixMissingEndpoint(strPlantID, endpointDesc, endpointType)
**Purpose**: Auto-creates missing endpoint or finds fuzzy match
**Parameters**:
- `strPlantID`: "WET_PLANT", "ORE_SORTER", or "RETREATMENT"
- `endpointDesc`: Description of missing endpoint
- `endpointType`: "SOURCE" or "DESTINATION" (for logging)
**Returns**: String (short name created or found)

**Auto-Fix Strategies**:
1. Try exact description match
2. Try fuzzy match (removes spaces, special chars)
3. Create new endpoint with "(Imported - Review)" marker
4. Generate short name from description

#### GetAutoFixReport()
**Purpose**: Gets report of all auto-fixes applied
**Returns**: String (formatted report)

#### NormalizePlantID(plantID)
**Purpose**: Normalizes plant ID variations
**Examples**:
- "WETPLANT" → "WET_PLANT"
- "Ore-Sorter" → "ORE_SORTER"
- "1" → "WET_PLANT"

### Backup Functions

#### CreateBackup(plantID)
**Purpose**: Creates JSON backup before import
**Parameters**: `plantID` - "ALL" or specific plant
**Returns**: String (path to backup file)
**Side Effect**: Auto-deletes backups older than the 10 most recent

#### ListBackups()
**Purpose**: Gets list of available backup files
**Returns**: Collection of backup file paths

## Import Modes Explained

### APPEND Mode
- Adds imported cables to existing ones
- Does NOT check for duplicates
- Safe for adding new data
- Use when: Merging from different sources

### REPLACE Mode
- **DANGEROUS**: Deletes ALL existing cables first
- Then imports from file
- Use when: Fresh installation, complete restore
- Always creates backup first

### MERGE Mode
- Updates existing cables by Cable ID
- Adds new cables that don't exist
- Does NOT delete cables not in import file
- Use when: Syncing changes, selective updates

## CSV File Format

### Cables CSV

```csv
Version,Plant,Scheduled,IDAttached,CableID,Source,Destination,CoreSize,EarthSize,CoreConfig,InsulationType,CableType,CableLength
2024.12.1,WET_PLANT,FALSE,TRUE,CV102-C1001-CV103,MCC 2,Conveyor belt 3,2.5mm²,1.5mm²,4C + E,HFI-90-TP,Black Circular SWA,45
```

### Endpoints CSV

```csv
Version,Plant,ShortName,Description
2024.12.1,WET_PLANT,MCC102,MCC 2
2024.12.1,WET_PLANT,CV103,Conveyor belt 3
```

## JSON File Format

See `docs/IMPORT_EXPORT_Implementation_Plan.md` for complete JSON spec.

## Error Handling

All functions include comprehensive error handling:

- Errors are logged to Debug.Print (Immediate Window)
- User-friendly error messages via MsgBox
- Functions return False or error dictionaries on failure
- No crashes - always returns gracefully

## Auto-Fix Examples

### Missing Endpoint

**Scenario**: Import cable with source "Water Pump 5" but endpoint doesn't exist

**Auto-Fix Actions**:
1. Search for exact match: "Water Pump 5" - NOT FOUND
2. Try fuzzy match: "waterpump5" vs existing - NOT FOUND
3. Generate short name: "WP" + plant digit + next number = "WP105"
4. Create endpoint: WP105 - "Water Pump 5 (Imported - Review)"
5. Log: "Created missing endpoint: WP105 - Water Pump 5 (Imported - Review)"
6. Return: "WP105"

### Fuzzy Match

**Scenario**: Import has "MCC2" but table has "MCC 2"

**Auto-Fix Actions**:
1. Exact match fails
2. Clean both: "mcc2" = "mcc2" - MATCH!
3. Find existing short name: "MCC102"
4. Log: "Fuzzy matched endpoint: 'MCC2' → MCC102"
5. Return: "MCC102"

## Performance Notes

- CSV export: ~100 cables/second
- JSON export: ~50 cables/second (more complex formatting)
- CSV import: ~50 cables/second (includes validation)
- Large datasets (1000+ cables): Consider progress bar (not yet implemented)

## Dependencies

These modules require:
- Excel tables with specific names (tbl_WetPlantCables, etc.)
- Worksheet codenames (sht_WetPlant, sht_Data, etc.)
- Scripting.FileSystemObject (standard in Office)
- Dictionary objects (standard in Office)

## Version Compatibility

- Built for version: 2024.12.1
- Compatible with: Excel 2010 and later
- VBA 7.0 and 7.1 (32-bit and 64-bit)

## Testing

Run these tests before production use:

```vba
Sub TestImportExport()
    ' Test 1: Export
    Debug.Print "Test 1: Export"
    If modImportExport.ExportCablesToCSV("ALL", "C:\Temp\test.csv") Then
        Debug.Print "✓ Export successful"
    Else
        Debug.Print "✗ Export failed"
    End If

    ' Test 2: Backup
    Debug.Print "Test 2: Backup"
    Dim backup As String
    backup = modBackup.CreateBackup("ALL")
    If backup <> "" Then
        Debug.Print "✓ Backup created: " & backup
    Else
        Debug.Print "✗ Backup failed"
    End If

    ' Test 3: Auto-Fix
    Debug.Print "Test 3: Auto-Fix"
    modCompatibilityFix.InitializeAutoFix
    Dim shortName As String
    shortName = modCompatibilityFix.FixMissingEndpoint("WET_PLANT", "Test Endpoint", "SOURCE")
    If shortName <> "" Then
        Debug.Print "✓ Auto-fix created: " & shortName
    Else
        Debug.Print "✗ Auto-fix failed"
    End If

    Debug.Print "Tests complete!"
End Sub
```

## Known Limitations

1. **JSON Import Not Implemented**: Currently only CSV import works
2. **No Progress Bar**: Large imports may appear frozen (but aren't)
3. **No Undo**: Once imported with REPLACE, only way back is from backup
4. **No Validation Rules Export**: Data validation dropdowns not exported
5. **Limited Fuzzy Matching**: Only basic string cleanup, not advanced algorithms

## Future Enhancements

- JSON import functionality
- Progress bar for large operations
- Export/import validation rules
- Scheduled auto-exports
- Cloud integration (SharePoint, OneDrive)
- Diff viewer (compare two versions)
- Advanced fuzzy matching (Levenshtein distance)
- Custom filter exports (by date, cable type, etc.)

## License

Copyright 2024 - Internal use only
Contact: jorr@mtcarbine.com.au

---

**Ready to use!** See `docs/IMPORT_EXPORT_Installation_Guide.md` for installation instructions.
