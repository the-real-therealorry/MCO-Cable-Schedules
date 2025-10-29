# MCO Cable Schedules - Power Cable Register

> **Part of the MCO Cable Schedules System**  
> Comprehensive cable management solution for Mt Carbine Operations  
> This repository contains three separate Excel workbooks for managing different types of industrial cabling across multiple processing plants.

A comprehensive VBA-based Excel application for managing industrial power cable infrastructure across multiple processing plants. This system provides a centralized database for tracking cable installations, endpoints, and technical specifications with built-in validation and data integrity features.

![Excel Version](https://img.shields.io/badge/Excel-2016%2B-green)
![VBA](https://img.shields.io/badge/VBA-7.0%2B-blue)
![License](https://img.shields.io/badge/license-MIT-blue)
![Status](https://img.shields.io/badge/status-active-success)

## 📦 System Components

This repository manages three separate but architecturally identical workbooks:

| Workbook | Status | Purpose |
|----------|--------|---------|
| **Power Cable Register Rev0.xlsm** | ✅ Active Development | High-voltage power distribution cables |
| **Control & Instrument Cable Register Rev0.xlsm** | 🔄 Planned | Control system and instrumentation cabling |
| **Structured Cable Register Rev0.xlsm** | 🔄 Planned | Communications and data network cabling |

> **Note:** This README documents the **Power Cable Register** workbook. The Control & Instrument and Structured Cable workbooks will be clones with minor label/naming changes to suit their specific cable types.

## 📋 Table of Contents

- [Overview](#overview)
- [Repository Structure](#repository-structure)
- [Features](#features)
- [System Architecture](#system-architecture)
- [Installation](#installation)
- [Usage Guide](#usage-guide)
- [Data Structure](#data-structure)
- [Development](#development)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)
- [Contact](#contact)

## 📁 Repository Structure

```
MCO-Cable-Schedules/
│
├── Power_Cable_Register_Rev0.xlsm           # High-voltage power cables ✅ Active
│   ├── 11 worksheets
│   ├── 13 class modules
│   ├── 2 user forms
│   └── 8 standard modules
│
├── Control_&_Instrument_Cable_Register_Rev0.xlsm  # C&I signals 🔄 Planned
│   └── (Clone of Power Cable Register with C&I-specific customizations)
│
├── Structured_Cable_Register_Rev0.xlsm      # Data/comms cables 🔄 Planned
│   └── (Clone of Power Cable Register with structured cabling customizations)
│
├── docs/                                    # Additional documentation
│   ├── ModDiagnostics.bas                   # Diagnostic tools
│   ├── TROUBLESHOOTING_GUIDE.txt            # Detailed troubleshooting
│   └── QUICK_REFERENCE.txt                  # Quick fix guide
│
├── README.md                                # This file
└── LICENSE                                  # MIT License
```

### Workbook Status

| Workbook | Status | Completion | Next Milestone |
|----------|--------|------------|----------------|
| Power Cable Register | 🟢 Active Development | 85% | User acceptance testing |
| Control & Instrument Cable Register | 🟡 Planned | 0% | Awaiting Power Cable completion |
| Structured Cable Register | 🟡 Planned | 0% | Awaiting C&I completion |

## 🎯 Overview

The MCO Cable Schedules system consists of three separate Excel workbooks, each managing a different category of industrial cabling. This modular approach allows specialized teams to work independently while maintaining consistent data structure and functionality across all cable types.

### Power Cable Register (This Workbook)

The Power Cable Register is designed for industrial facilities managing high-voltage electrical infrastructure across multiple processing plants. It provides:

- **Centralized cable management** for three plant types:
  - Wet Screen Crushing Plant
  - Ore Sorting Plant
  - Retreatment Gravity Plant
  
- **Endpoint tracking** with hierarchical naming conventions
- **Cable specifications** including core size, insulation type, and configuration
- **Visual interface** with plant-specific color coding
- **Data validation** with regex-based format checking
- **Edit/delete functionality** with data loss prevention
- **Automatic circuit numbering** for each plant type

### Planned Workbooks

#### Control & Instrument Cable Register
- **Purpose:** Manage control system and instrumentation cabling
- **Cable Types:** 4-20mA signals, digital I/O, PLCs, field devices
- **Status:** Planned - will be created after Power Cable Register is fully tested
- **Architecture:** Clone of Power Cable Register with adapted labels and specifications

#### Structured Cable Register  
- **Purpose:** Manage communications and data network cabling
- **Cable Types:** Ethernet, fiber optic, telephone, CCTV, access control
- **Status:** Planned - will be created after Power Cable Register is fully tested
- **Architecture:** Clone of Power Cable Register with adapted labels and specifications

### Why Separate Workbooks?

1. **Separation of Concerns:** Different teams manage different cable types
2. **Reduced Complexity:** Smaller, focused datasets per workbook
3. **Performance:** Better Excel performance with smaller files
4. **Security:** Different access levels for different cable types
5. **Specialization:** Cable-type-specific validations and lookup tables

### Shared Architecture

All three workbooks share the same:
- VBA module structure and class hierarchy
- Form designs and user interface
- Validation logic and error handling
- Edit/delete functionality
- Dashboard navigation pattern
- Data integrity mechanisms

Only minor changes needed between workbooks:
- Form labels and titles
- Lookup table values (cable types, sizes, etc.)
- Validation patterns (if cable ID formats differ)
- Color themes (optional differentiation)

## ✨ Features

### Core Functionality

- **📝 Cable Registration**
  - Register new power cables with comprehensive specifications
  - Automatic cable ID generation (Source-Circuit-Destination format)
  - Auto-incrementing circuit numbers per plant
  - Source and destination endpoint selection
  - Technical specifications (core size, earth size, configuration, insulation, type)
  - Cable length tracking

- **🔌 Endpoint Management**
  - Create and manage cable endpoints (connection points)
  - Plant-specific naming conventions (1xx for Wet Plant, 2xx for Ore Sorter, 3xx for Retreatment)
  - Short name + detailed description
  - Duplicate prevention
  - Format validation

- **✏️ Edit & Update**
  - Edit existing cable records
  - Data integrity validation
  - Endpoint lookup verification
  - Data loss prevention mechanisms

- **🗑️ Delete Functionality**
  - Row-level delete with confirmation
  - Floating action buttons (Edit/Delete) on row selection
  - Table structure preservation

### User Interface

- **🎨 Plant-Specific Themes**
  - Wet Plant: Dark blue theme
  - Ore Sorter: Teal theme
  - Retreatment: Gray/red theme

- **📊 Dashboard**
  - Central navigation hub
  - Quick access to all plant forms
  - Visual plant identification

- **🔘 Dynamic Buttons**
  - Edit/Delete buttons appear on row selection
  - Context-aware button positioning
  - Color-coded actions (red for delete, blue for edit)

### Data Integrity

- **✅ Validation**
  - Regex-based format validation for cable IDs and endpoint names
  - Required field checking
  - Numeric validation for cable length
  - Duplicate prevention for endpoints

- **🛡️ Error Handling**
  - Centralized error handling system
  - Comprehensive error logging
  - User-friendly error messages
  - Debug output to Immediate Window

- **💾 Data Loss Prevention**
  - Endpoint lookup validation before editing
  - Automatic backup prompts
  - Confirmation dialogs for destructive operations

## 🏗️ System Architecture

### Application Structure

```
Power_Cable_Register_Rev0.xlsm
│
├── Worksheets
│   ├── Dashboard (Main navigation)
│   ├── Crushing Wet Screen Plant (Cable data)
│   ├── Ore Sorting Plant (Cable data)
│   ├── Retreatment Gravity Plant (Cable data)
│   └── Data (Endpoints and lookup tables)
│
├── Class Modules
│   ├── Dashboard.cls (Dashboard initialization)
│   ├── ThisWorkbook.cls (Workbook events)
│   ├── sht_WetPlant.cls (Wet plant operations)
│   ├── sht_OreSorter.cls (Ore sorter operations)
│   ├── sht_Retreatment.cls (Retreatment operations)
│   ├── sht_Data.cls (Endpoint management)
│   ├── clCable.cls (Cable object)
│   └── clEndpoint.cls (Endpoint object)
│
├── UserForms
│   ├── frm_RegisterCable (Cable entry/edit form)
│   └── frm_RegisterEndpoint (Endpoint entry form)
│
└── Standard Modules
    ├── modDatabase.bas (Database abstraction layer)
    ├── modDashCode.bas (Form launching)
    ├── ModRowActions.bas (Edit/Delete buttons)
    ├── ModError.bas (Error handling)
    ├── ModAppState.bas (Performance optimization)
    ├── modUtils.bas (Utility functions)
    ├── modUDT.bas (Type definitions)
    └── ModDiagnostics.bas (Troubleshooting tools)
```

### Data Model

#### Cable Record
```
Cable {
  IsScheduled: Boolean
  IDAttached: Boolean
  CableID: String (format: SOURCE-CIRCUIT-DESTINATION)
  Source: String (Endpoint description)
  Destination: String (Endpoint description)
  CoreSize: String
  EarthSize: String
  CoreConfig: String
  InsulationType: String
  CableType: String
  CableLength: Integer
}
```

#### Endpoint Record
```
Endpoint {
  ShortName: String (format: [A-Z]{2,3}[PlantID][0-9]{2})
  Description: String
}
```

#### Plant Identifiers
- **Wet Plant:** `1` (e.g., WM101, CV102)
- **Ore Sorter:** `2` (e.g., OS201, CV202)
- **Retreatment:** `3` (e.g., RP301, CV302)

### Excel Tables

| Table Name | Purpose | Columns |
|------------|---------|---------|
| `tbl_WetPlantCables` | Wet plant cable records | 11 columns |
| `tbl_OreSorterCables` | Ore sorter cable records | 11 columns |
| `tbl_RetreatmentCables` | Retreatment cable records | 11 columns |
| `tbl_WetPlantEndpoints` | Wet plant endpoints | 2 columns (ShortName, Description) |
| `tbl_OreSorterEndpoints` | Ore sorter endpoints | 2 columns |
| `tbl_RetreatmentEndpoints` | Retreatment endpoints | 2 columns |
| `tbl_CableSizes` | Lookup: Available cable sizes | 1 column |
| `tbl_CoreConfigs` | Lookup: Core configurations | 1 column |
| `tbl_InsulationTypes` | Lookup: Insulation types | 1 column |
| `tbl_CableTypes` | Lookup: Cable types | 1 column |

## 📥 Installation

### Prerequisites

- **Microsoft Excel** 2016 or later
- **Windows** 10 or later (for VBA 7.0)
- **Macro-enabled workbook** support

### Setup Steps

1. **Download the repository**
   ```bash
   git clone https://github.com/yourusername/MCO-Cable-Schedules.git
   cd MCO-Cable-Schedules
   ```

2. **Open the appropriate workbook**
   - For power cables: `Power_Cable_Register_Rev0.xlsm`
   - For control/instrument: `Control_&_Instrument_Cable_Register_Rev0.xlsm` (when available)
   - For structured cabling: `Structured_Cable_Register_Rev0.xlsm` (when available)

3. **Enable macros**
   - Open the workbook
   - Click "Enable Content" when prompted
   - If macros are blocked, go to File → Options → Trust Center → Trust Center Settings → Macro Settings
   - Select "Enable all macros" (not recommended for production) or add folder to Trusted Locations

4. **Initialize the system**
   - The Dashboard will automatically load on workbook open
   - If not, press `Alt+F8`, select `Dashboard.initialize`, click Run

5. **Verify installation**
   - Check that Dashboard buttons respond to clicks
   - Try registering a test endpoint
   - Try registering a test cable

### Optional: Diagnostic Tools

For troubleshooting, import the diagnostic module:

1. Press `Alt+F11` to open VBA Editor
2. File → Import File
3. Select `ModDiagnostics.bas` (if provided separately)
4. Use commands in Immediate Window (`Ctrl+G`):
   ```vba
   DiagnoseEndpointLookup "CableID", "PlantType"
   ListAllEndpoints "WET_PLANT"
   ```

## 📖 Usage Guide

### Registering a New Endpoint

1. **Navigate to Dashboard**
2. **Click** "Register New Endpoint" for the appropriate plant
3. **Enter details:**
   - **Short Name:** Follow format (e.g., `WM101` for Wet Plant)
     - 2-3 uppercase letters
     - Plant identifier digit (1, 2, or 3)
     - 2-digit number (00-19)
   - **Description:** Clear, descriptive name (e.g., "Wet Mill Motor 1")
4. **Click "Save"** or **"Save & Continue"** for multiple entries

### Registering a New Cable

1. **Navigate to Dashboard**
2. **Click** "Register New Cable" for the appropriate plant
3. **Fill in details:**
   - **Source:** Select from dropdown (create endpoint first if needed)
   - **Circuit Name:** Auto-generated (e.g., `C1001`)
   - **Destination:** Select from dropdown
   - **Cable ID:** Auto-populated (`SOURCE-CIRCUIT-DESTINATION`)
   - **Core Size:** Select from dropdown
   - **Earth Size:** Select from dropdown
   - **Configuration:** Select from dropdown (e.g., `3C` for 3-core)
   - **Insulation Type:** Select from dropdown
   - **Cable Type:** Select from dropdown
   - **Length (m):** Enter numeric value
4. **Check boxes** (optional):
   - Schedule Completed
   - ID Securely Attached
5. **Click "Save"** or **"Save & Continue"**

### Editing a Cable

1. **Navigate to** the plant worksheet (not Dashboard)
2. **Click on a row** to select a cable
3. **Edit/Delete buttons** appear to the right
4. **Click "Edit"** button
5. **Modify** fields (Source/Destination/Circuit Name are locked)
6. **Click "Save"**

> **Note:** If you see an error about missing endpoints, use the diagnostic tool to identify the issue.

### Deleting a Cable

1. **Navigate to** the plant worksheet
2. **Click on a row** to select a cable
3. **Click "Del"** button
4. **Confirm** deletion when prompted

> **Warning:** Deletion is permanent and cannot be undone!

## 📊 Data Structure

### Cable ID Format

```
SOURCE-CIRCUIT-DESTINATION

Example: WM101-C1001-CV102
         │    │    └─ Destination endpoint short name
         │    └────── Circuit identifier
         └─────────── Source endpoint short name
```

### Endpoint Naming Convention

```
[PREFIX][PLANT_ID][NUMBER]

Examples:
- WM101  → Wet Mill, Plant 1, Number 01
- CV102  → Conveyor, Plant 1, Number 02
- OS201  → Ore Sorter, Plant 2, Number 01
- RP301  → something, Plant 3, Number 01
```

### Circuit Numbering

- **Wet Plant:** C1001, C1002, C1003...
- **Ore Sorter:** C2001, C2002, C2003...
- **Retreatment:** C3001, C3002, C3003...

### Validation Patterns (Regex)

#### Wet Plant
- **Endpoint:** `^[A-Z]{2,3}[1][0-9]{2}$`
- **Circuit:** `^[C][1][0-9]{3}$`
- **Cable ID:** `^[A-Z]{2,3}[1][01][0-9][-][C][1][0-9]{3}[-][A-Z]{2,3}[1][01][0-9]$`

#### Ore Sorter
- **Endpoint:** `^[A-Z]{2,3}[2][0-9]{2}$`
- **Circuit:** `^[C][2][0-9]{3}$`
- **Cable ID:** `^[A-Z]{2,3}[2][01][0-9][-][C][2][0-9]{3}[-][A-Z]{2,3}[2][01][0-9]$`

#### Retreatment
- **Endpoint:** `^[A-Z]{2,3}[3][0-9]{2}$`
- **Circuit:** `^[C][3][0-9]{3}$`
- **Cable ID:** `^[A-Z]{2,3}[3][01][0-9][-][C][3][0-9]{3}[-][A-Z]{2,3}[3][01][0-9]$`

## 🔧 Development

### Architecture Patterns

- **Facade Pattern:** `modDatabase.bas` provides unified interface to plant-specific modules
- **Object-Oriented Design:** `clCable` and `clEndpoint` classes encapsulate data
- **Separation of Concerns:** UI (forms) separated from business logic (modules) separated from data (worksheets)
- **Error Handling:** Centralized `HandleError` procedure for consistent error management
- **State Management:** `ModAppState` for performance optimization during bulk operations

### Key Design Decisions

1. **Excel Tables over Ranges:** Ensures data structure integrity and simplifies references
2. **Class Modules for Worksheets:** Encapsulates plant-specific logic, enables polymorphism
3. **Regex Validation:** Ensures consistent data format across all records
4. **Disabled Fields During Edit:** Prevents corruption of primary key (Cable ID)
5. **Floating Action Buttons:** Improves UX without cluttering worksheet

### Module Responsibilities

| Module | Responsibility |
|--------|----------------|
| `modDatabase` | Abstraction layer, routing to plant-specific methods |
| `modDashCode` | Form initialization and launching |
| `ModRowActions` | Dynamic button creation and positioning |
| `ModError` | Centralized error handling and logging |
| `ModAppState` | Excel application state management |
| `modUtils` | String manipulation, regex, validation utilities |
| `sht_*` classes | Plant-specific CRUD operations |

### Performance Considerations

- **Application State Management:** Disable screen updating/events during bulk operations
- **Array Operations:** Convert ranges to arrays for processing to minimize Excel interaction
- **Table References:** Use ListObject references instead of Range.Find for reliability

### Adding a New Plant Type

1. **Create new worksheet** (e.g., "New Plant Name")
2. **Add worksheet class module** (e.g., `sht_NewPlant.cls`)
3. **Copy structure** from existing plant module (e.g., `sht_WetPlant.cls`)
4. **Update constants:**
   - Plant identifier (e.g., `4`)
   - Regex patterns for validation
   - Circuit number starting point (e.g., `C4001`)
5. **Create endpoints table** (e.g., `tbl_NewPlantEndpoints`)
6. **Create cables table** (e.g., `tbl_NewPlantCables`)
7. **Add form ID** (e.g., `"NEW_PLANT"`) to:
   - `modDatabase.bas` (all functions)
   - `frm_RegisterCable.UserForm_Activate`
   - `frm_RegisterEndpoint.UserForm_Activate`
   - `sht_Data.GetEndpointsArray`
8. **Add Dashboard button** and initialize in `Dashboard.initialize`
9. **Test thoroughly** before production use

### Creating Clone Workbooks (Control & Instrument, Structured)

When creating the Control & Instrument and Structured Cable workbooks:

1. **Make a copy** of `Power_Cable_Register_Rev0.xlsm`
   ```bash
   cp Power_Cable_Register_Rev0.xlsm Control_&_Instrument_Cable_Register_Rev0.xlsm
   ```

2. **Open the new workbook** in Excel

3. **Update workbook-level items:**
   - Workbook title in `ThisWorkbook` comments
   - Dashboard sheet title/header
   - Form captions and labels

4. **Update lookup tables** (on Data sheet):
   - `tbl_CableSizes` - relevant sizes for cable type
   - `tbl_CoreConfigs` - relevant configurations
   - `tbl_InsulationTypes` - relevant insulation types
   - `tbl_CableTypes` - relevant cable types for the discipline

5. **Update validation patterns** (if needed):
   - Cable ID format in form activation code
   - Endpoint naming if different convention needed

6. **Update color themes** (optional):
   - Form background colors
   - Plant-specific theming
   - Table styles

7. **Clear existing data:**
   - Delete all rows from cable tables
   - Delete all rows from endpoint tables (or add discipline-specific endpoints)

8. **Test thoroughly:**
   - Register endpoints
   - Register cables
   - Edit cables
   - Delete cables
   - Verify all validations work

9. **Update documentation:**
   - Update form help text
   - Update sheet names if desired
   - Update Dashboard instructions

### Recommended Customizations per Workbook

#### Control & Instrument Cable Register
- **Cable Types:** Multi-pair signal, screened, armored signal cables
- **Core Configs:** 2-pair, 4-pair, 8-pair, 16-pair, etc.
- **Sizes:** Signal cable sizes (e.g., 0.5mm², 0.75mm², 1.0mm²)
- **Additional Fields (optional):** Shield type, intrinsic safety rating
- **Color Theme:** Orange/amber to distinguish from power

#### Structured Cable Register
- **Cable Types:** Cat5e, Cat6, Cat6a, Fiber (SM/MM), Coax
- **Core Configs:** 4-pair UTP, 4-pair FTP, 12-core fiber, 24-core fiber
- **Sizes:** N/A for data cables (use cable category instead)
- **Additional Fields (optional):** Bandwidth rating, fiber type
- **Color Theme:** Green to distinguish from electrical

## 🔍 Troubleshooting

### Common Issues

#### Issue: Edit form shows empty Source/Destination fields

**Cause:** Endpoint lookup failed - stored description doesn't match any endpoint in the endpoints table.

**Solution:**
1. Press `Ctrl+G` to open Immediate Window
2. Run: `DiagnoseEndpointLookup "CableID", "PlantType"`
3. Follow diagnostic output to identify missing/mismatched endpoints
4. Add missing endpoints or fix spelling differences

**Prevention:** Improved `ShowForUpdate` method validates lookup before opening form.

---

#### Issue: "Invalid property value" error when editing

**Cause:** Attempting to set combo box value that doesn't exist in the list.

**Solution:**
1. Verify endpoint exists in endpoints table
2. Check for exact spelling match (spaces, capitalization)
3. Use diagnostic tool to identify the issue

---

#### Issue: Duplicate cable ID error

**Cause:** Cable ID already exists in the table.

**Solution:**
1. Check if cable was accidentally created twice
2. Delete duplicate if appropriate
3. Or modify source/circuit/destination to create unique ID

---

#### Issue: Buttons not appearing when row is selected

**Cause:** Worksheet event handler may not be firing or shapes are hidden.

**Solution:**
1. Check that worksheet is not protected
2. Verify `Worksheet_SelectionChange` event exists in worksheet class module
3. Check shape names: `AI_RowDel` and `AI_RowEdit`
4. Manually run: `ModRowActions.PositionRowButtons Worksheets("SheetName"), RowNumber`

---

#### Issue: Form doesn't open when Dashboard button clicked

**Cause:** Button OnAction property not set or procedure doesn't exist.

**Solution:**
1. Press `Alt+F11` to open VBA Editor
2. Run `Dashboard.ValidateAllButtons` in Immediate Window
3. Check output for missing shapes or procedures
4. Re-run `Dashboard.initialize`

### Diagnostic Commands

Available in Immediate Window (`Ctrl+G`):

```vba
' Diagnose endpoint lookup for specific cable
DiagnoseEndpointLookup "CV103-C1001-CV103", "WET_PLANT"

' List all endpoints for a plant
ListAllEndpoints "WET_PLANT"
ListAllEndpoints "ORE_SORTER"
ListAllEndpoints "RETREATMENT"

' Validate dashboard button setup
Dashboard.ValidateAllButtons

' Reset all dashboard buttons
Dashboard.ResetAllButtons

' Re-initialize dashboard
Dashboard.initialize
```

### Debug Mode

To enable detailed logging:

1. Open `ModError.bas`
2. Set constants:
   ```vba
   Private Const SHOW_MSG As Boolean = True
   Private Const LOG_TO_IMMEDIATE As Boolean = True
   ```
3. All errors will be logged to Immediate Window

## 🤝 Contributing

Contributions are welcome! Please follow these guidelines:

### Reporting Issues

1. **Check existing issues** to avoid duplicates
2. **Use issue template** (if provided)
3. **Include:**
   - Excel version
   - Windows version
   - Steps to reproduce
   - Expected vs actual behavior
   - Screenshots if applicable
   - Error messages from Immediate Window

### Submitting Changes

1. **Fork the repository**
2. **Create a feature branch:** `git checkout -b feature/your-feature-name`
3. **Follow coding standards:**
   - Use descriptive variable names
   - Add comments for complex logic
   - Include error handling
   - Update documentation
4. **Test thoroughly:**
   - All three plant types
   - Create, edit, delete operations
   - Edge cases and error conditions
5. **Commit with clear messages:** `git commit -m "Add feature: description"`
6. **Push to your fork:** `git push origin feature/your-feature-name`
7. **Open a Pull Request** with description of changes

### Coding Standards

- **Naming Conventions:**
  - Variables: camelCase (`strCableID`, `lngRowNumber`)
  - Constants: UPPER_SNAKE_CASE (`BTN_WIDTH`, `RE_WPEP_FORMAT`)
  - Functions/Subs: PascalCase (`GetNextCircuitNumber`)
  - Private members: prefix with `m` (`mCableID`, `mDescription`)

- **Error Handling:**
  - Use `HandleError` procedure for all error handling
  - Include module name, procedure name, error number, description, and line number
  - Provide user-friendly error messages

- **Documentation:**
  - Comment block at top of each module
  - Function header comments with purpose, parameters, returns, notes
  - Inline comments for complex logic

- **Code Organization:**
  - One responsibility per module
  - Keep functions focused and under 50 lines when possible
  - Extract complex logic into helper functions

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

```
MIT License

Copyright (c) 2024 [Your Name/Organization]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## 📞 Contact

**Project Maintainer:** jorr@mtcarbine.com.au

**Repository:** [MCO-Cable-Schedules](https://github.com/yourusername/MCO-Cable-Schedules)

**Issues & Discussions:** [GitHub Issues](https://github.com/yourusername/MCO-Cable-Schedules/issues)

## 🙏 Acknowledgments

- Built with Microsoft Excel VBA
- Designed for industrial cable management
- Enhanced with AI assistance (Claude by Anthropic)

## 📈 Project Status

**System Name:** MCO Cable Schedules  
**Current Phase:** Power Cable Register - Rev 0 (Initial Release)  
**Overall Status:** Active Development (Phase 1 of 3)

### Workbook Status Details

#### Power Cable Register Rev0.xlsm
- **Status:** 🟢 Active Development  
- **Completion:** ~85%
- **Last Updated:** October 2025
- **Current Focus:** Testing, bug fixes, user acceptance
- **Next Milestone:** Production-ready release

#### Control & Instrument Cable Register Rev0.xlsm
- **Status:** 🟡 Planned
- **Completion:** 0%
- **Expected Start:** Q1 2026 (after Power Cable testing complete)
- **Estimated Duration:** 2-3 weeks (clone + customization)

#### Structured Cable Register Rev0.xlsm
- **Status:** 🟡 Planned
- **Completion:** 0%
- **Expected Start:** Q1 2026 (after C&I Register complete)
- **Estimated Duration:** 2-3 weeks (clone + customization)

---

## 🗺️ Roadmap

### Current Phase: Power Cable Register (Rev 0)

**Status:** Active Development  
**Target Completion:** Testing and stabilization

**Immediate Priorities:**
- [ ] Complete testing of all CRUD operations
- [ ] Validate data integrity mechanisms
- [ ] Finalize edit/delete functionality
- [ ] User acceptance testing
- [ ] Bug fixes and refinements

### Phase 2: Control & Instrument Cable Register

**Status:** Planned  
**Start Date:** After Power Cable Register is fully tested

**Tasks:**
- [ ] Clone Power Cable Register workbook
- [ ] Update cable types and specifications for C&I
- [ ] Customize lookup tables (multi-pair cables, signal types)
- [ ] Update form labels and titles
- [ ] Test with C&I-specific data
- [ ] User acceptance testing

### Phase 3: Structured Cable Register

**Status:** Planned  
**Start Date:** After Control & Instrument Cable Register is complete

**Tasks:**
- [ ] Clone Power Cable Register workbook
- [ ] Update cable types (Cat5e/6/6a, fiber, coax)
- [ ] Customize lookup tables (data cable specifications)
- [ ] Update form labels and titles
- [ ] Test with structured cabling data
- [ ] User acceptance testing

### Future Enhancements (All Workbooks)

- [ ] **Export/Import functionality** - CSV/Excel export for reporting
- [ ] **Search and filter** - Advanced search across all plants
- [ ] **Cable scheduling integration** - SharePoint integration for installation scheduling
- [ ] **Bulk operations** - Import multiple cables/endpoints from CSV
- [ ] **Reporting module** - Generate cable installation reports
- [ ] **Audit trail** - Track who created/modified/deleted records
- [ ] **Data validation rules** - Additional business logic validation
- [ ] **Cross-workbook integration** - Link related cables across disciplines
- [ ] **Mobile-friendly version** - Web-based interface using Office Scripts

### Known Limitations

- **Single-user:** Excel file locking prevents simultaneous editing
- **No version control:** Manual backup required for change tracking
- **Limited scalability:** Excel table performance degrades above ~10,000 rows
- **Windows-only:** VBA requires Windows Excel (Mac Excel has limited VBA support)

### Future Considerations

- Migration to web-based system (React + Node.js + PostgreSQL)
- Multi-user support with proper database backend
- REST API for integration with other systems
- Mobile app for field technicians

---

## 📚 Additional Resources

### Documentation

- [Excel VBA Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
- [Regular Expressions in VBA](https://www.regular-expressions.info/vba.html)
- [Excel Table (ListObject) Programming](https://docs.microsoft.com/en-us/office/vba/api/excel.listobject)

### Related Projects

- [Excel VBA Best Practices](https://github.com/topics/vba-best-practices)
- [Excel Database Templates](https://github.com/topics/excel-database)

---

**⭐ If you find this project useful, please consider giving it a star!**

---

*Last updated: October 29, 2025*
