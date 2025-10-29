# Power Cable Register - User Guide

**Application Version:** Rev0  
**Author:** jorr@mtcarbine.com.au  
**Last Updated:** December 2024  
**Document Purpose:** End-user guide for the Power Cable Register application

---

## Table of Contents

1. [Introduction](#introduction)
2. [Getting Started](#getting-started)
3. [Dashboard Overview](#dashboard-overview)
4. [Managing Endpoints](#managing-endpoints)
5. [Managing Cables](#managing-cables)
6. [Viewing and Editing Records](#viewing-and-editing-records)
7. [Plant Types](#plant-types)
8. [Data Validation Rules](#data-validation-rules)
9. [Troubleshooting](#troubleshooting)
10. [Glossary](#glossary)

---

## 1. Introduction

### Purpose
The Power Cable Register is a Microsoft Excel-based application designed to track and manage power cable installations across three distinct plant types:
- **Wet Plant (Crushing Wet Screen Plant)**
- **Ore Sorter Plant**
- **Retreatment Plant (Gravity Plant)**

### Key Features
- Centralized dashboard for quick access to all plant types
- Endpoint registration and management
- Cable registration with detailed specifications
- Automatic circuit numbering
- Data validation and duplicate prevention
- Edit and delete capabilities with floating action buttons

### System Requirements
- Microsoft Excel with macro support enabled (.xlsm file format)
- Excel 2013 or later recommended for best compatibility
- Macros must be enabled for the application to function

---

## 2. Getting Started

### Opening the Application

1. **Locate the file**: Find `Power_Cable_Register_Rev0.xlsm` on your computer
2. **Enable Macros**: When you open the file, you may see a security warning. Click **Enable Content** or **Enable Macros** to allow the application to run
3. **Dashboard Launch**: The application will automatically:
   - Initialize all dashboard buttons
   - Activate the Dashboard worksheet
   - Display the main interface

### First Time Setup
No special setup is required. The application creates and manages its own data tables internally.

---

## 3. Dashboard Overview

The Dashboard is the main control center of the application. It provides quick access to all registration functions for each plant type.

### Dashboard Layout

The dashboard contains **six main buttons** organized by function:

#### Cable Registration Buttons
- **Wet Plant Cable** - Register new cables for the Crushing Wet Screen Plant
- **Ore Sorter Cable** - Register new cables for the Ore Sorting Plant
- **Retreatment Cable** - Register new cables for the Retreatment Gravity Plant

#### Endpoint Registration Buttons
- **Wet Plant Endpoint** - Register new endpoints for the Crushing Wet Screen Plant
- **Ore Sorter Endpoint** - Register new endpoints for the Ore Sorting Plant
- **Retreatment Endpoint** - Register new endpoints for the Retreatment Gravity Plant

### Navigation
Simply click the appropriate button to open the registration form for your desired plant type and function (cable or endpoint).

---

## 4. Managing Endpoints

### What is an Endpoint?
An endpoint represents a connection point in the electrical system where a power cable begins (source) or ends (destination). Examples include:
- Motor control centers (MCC)
- Distribution boards
- Equipment connection points
- Switchgear

### Registering a New Endpoint

1. **Access the Form**:
   - From the Dashboard, click the appropriate endpoint registration button for your plant type
   - The "Register Endpoint" form will open

2. **Fill in Required Information**:
   - **Short Name**: A unique identifier for the endpoint
   - **Description**: A detailed description of the endpoint's purpose or location

3. **Submit**:
   - Click the **Submit** button to save the endpoint
   - The form will validate your input and add the endpoint to the database

### Endpoint Naming Conventions

Each plant type has a specific format for endpoint short names:

#### Wet Plant Endpoints
- Format: **2-3 letters + "1" + 2 digits**
- Examples: `AB101`, `ABC105`, `MCC110`
- The "1" in the third or fourth position identifies this as a Wet Plant endpoint

#### Ore Sorter Endpoints
- Format: **2-3 letters + "2" + 2 digits**
- Examples: `AB201`, `XYZ205`, `OSP220`
- The "2" in the third or fourth position identifies this as an Ore Sorter endpoint

#### Retreatment Plant Endpoints
- Format: **2-3 letters + "3" + 2 digits**
- Examples: `AB301`, `RGP305`, `GRV320`
- The "3" in the third or fourth position identifies this as a Retreatment Plant endpoint

### Validation Rules

The application will reject an endpoint if:
- The short name format doesn't match the required pattern for that plant type
- The short name already exists in the database (duplicates not allowed)
- The description already exists in the database (duplicates not allowed)
- Any required field is empty

---

## 5. Managing Cables

### Cable Registration Overview

Cable registration captures detailed information about power cables installed between endpoints.

### Registering a New Cable

1. **Access the Form**:
   - From the Dashboard, click the appropriate cable registration button for your plant type
   - The "Register Cable" form will open

2. **Fill in Cable Information**:

   **Status Fields:**
   - **Scheduled**: Check if the cable installation is scheduled but not yet complete
   - **ID Attached**: Check if cable identification tags have been physically attached

   **Connection Information:**
   - **Cable ID**: The circuit number (auto-generated by the system)
   - **Source**: The starting endpoint (select from registered endpoints)
   - **Destination**: The ending endpoint (select from registered endpoints)

   **Cable Specifications:**
   - **Core Size**: Wire gauge of the cable cores (e.g., "4mm²", "10mm²")
   - **Earth Size**: Wire gauge of the earth/ground conductor
   - **Core Configuration**: Number and arrangement of cores (e.g., "3C+E" for 3 cores plus earth)
   - **Cable Type**: Type designation (e.g., "NYFGBY", "H07RN-F")
   - **Insulation Type**: Type of cable insulation (e.g., "PVC", "XLPE")
   - **Cable Length**: Physical length in meters

3. **Circuit Number**:
   - The system automatically assigns the next available circuit number for the selected plant
   - Circuit numbers are sequential and cannot be manually edited

4. **Submit**:
   - Click the **Submit** button to save the cable record
   - The form will validate your input and add the cable to the appropriate plant worksheet

### Understanding Circuit Numbers

Each plant maintains its own sequential circuit numbering:
- **Wet Plant**: Starts at 1 and increments (1, 2, 3, ...)
- **Ore Sorter**: Starts at 1 and increments (1, 2, 3, ...)
- **Retreatment**: Starts at 1 and increments (1, 2, 3, ...)

Circuit numbers are automatically assigned and cannot be skipped or edited.

---

## 6. Viewing and Editing Records

### Viewing Cable Data

Cable data for each plant type is stored in separate worksheets:
- **Crushing Wet Screen Plant** - Contains all Wet Plant cables
- **Ore Sorting Plant** - Contains all Ore Sorter cables
- **Retreatment Gravity Plant** - Contains all Retreatment cables

### Floating Action Buttons

When you select a cable row in any plant worksheet, two action buttons automatically appear next to the row:

#### Delete Button (Red)
- Click to delete the selected cable record
- You will be asked to confirm before deletion
- **Warning**: Deletion cannot be undone

#### Edit Button (Blue)
- Click to edit the selected cable record
- **Note**: Edit functionality is currently under development

### Using Action Buttons

1. **Select a Row**: Click on any cell in the cable record you want to work with
2. **Buttons Appear**: The Delete and Edit buttons will appear to the right of the row
3. **Click the Button**: Click the appropriate button for your desired action
4. **Multiple Selection**: If you select multiple rows, the buttons will hide (select only one row at a time)

### Deleting a Cable

1. Click on the cable row you want to delete
2. Click the **Del** button that appears
3. Confirm the deletion when prompted
4. The cable record will be permanently removed from the table

**Important**: Deleting a cable does NOT delete the endpoints it was connected to. Endpoints remain available for future cable assignments.

---

## 7. Plant Types

### Overview of the Three Plant Types

The application manages cables for three distinct processing facilities:

### 1. Wet Plant (Crushing Wet Screen Plant)
- **Full Name**: Crushing Wet Screen Plant
- **Abbreviation**: CWSP or WP
- **Endpoint Format**: `XX1XX` (letters-1-numbers)
- **Table Name**: `tbl_WetPlantCables`
- **Worksheet**: "Crushing Wet Screen Plant"

### 2. Ore Sorter
- **Full Name**: Ore Sorting Plant
- **Abbreviation**: OSP or OS
- **Endpoint Format**: `XX2XX` (letters-2-numbers)
- **Table Name**: `tbl_OreSorterCables`
- **Worksheet**: "Ore Sorting Plant"

### 3. Retreatment Plant
- **Full Name**: Retreatment Gravity Plant
- **Abbreviation**: RGP or GP
- **Endpoint Format**: `XX3XX` (letters-3-numbers)
- **Table Name**: `tbl_RetreatmentCables`
- **Worksheet**: "Retreatment Gravity Plant"

### Data Separation

- Each plant type maintains completely separate cable and endpoint data
- Circuit numbers are independent for each plant
- Endpoints registered in one plant cannot be used in another plant
- This separation prevents confusion and maintains data integrity

---

## 8. Data Validation Rules

### Endpoint Validation

**Short Name Requirements:**
- Must match the exact format for the plant type (see Section 4)
- Must be unique (no duplicates allowed)
- Must be uppercase letters and numbers only
- Length: 5 characters total (e.g., `AB101` or `ABC201`)

**Description Requirements:**
- Must be unique (no duplicates allowed)
- Cannot be empty
- Should be descriptive and meaningful

**Rejection Scenarios:**
- Invalid format for the plant type
- Duplicate short name exists
- Duplicate description exists
- Required fields are empty

### Cable Validation

**Required Fields:**
All cable fields are required and cannot be left empty:
- Cable ID (auto-generated)
- Source endpoint
- Destination endpoint
- Core Size
- Earth Size
- Core Configuration
- Cable Type
- Insulation Type
- Cable Length

**Source and Destination:**
- Must select from registered endpoints only
- Cannot use unregistered endpoint names
- Source and destination can be the same (for loops)

**Cable ID:**
- Automatically generated by the system
- Cannot be manually edited
- Sequential numbering per plant type

---

## 9. Troubleshooting

### Common Issues and Solutions

#### Issue: Dashboard buttons don't work
**Solution:**
- Ensure macros are enabled
- Close and reopen the workbook
- Check that you haven't renamed any worksheets

#### Issue: "Endpoint name already in use" error
**Solution:**
- Choose a different short name
- Check the endpoint list for that plant type
- Verify you're using the correct format for the plant type

#### Issue: Action buttons (Edit/Delete) don't appear
**Solution:**
- Make sure you've selected only ONE row
- Click on a cell within the cable data area (not the header)
- The worksheet must be one of the three plant worksheets

#### Issue: Cannot submit cable - validation error
**Solution:**
- Check that all required fields are filled in
- Verify source and destination endpoints exist
- Ensure you're selecting registered endpoints from the dropdown

#### Issue: Edit button shows "Not yet implemented"
**Solution:**
- Edit functionality is under development
- For now, you can manually edit cells in the table
- Delete and re-create the record if needed

#### Issue: Endpoint format rejected
**Solution:**
- Double-check the format requirements for your plant type:
  - Wet Plant: `XX1XX` (2-3 letters, then "1", then 2 numbers)
  - Ore Sorter: `XX2XX` (2-3 letters, then "2", then 2 numbers)
  - Retreatment: `XX3XX` (2-3 letters, then "3", then 2 numbers)
- Ensure all letters are UPPERCASE
- Use numbers only in the numeric positions

### Getting Help

If you encounter issues not covered in this guide:
1. Check the Immediate Window in VBA (Ctrl+G) for detailed error messages
2. Contact the application developer: jorr@mtcarbine.com.au
3. Note the exact error message and steps that caused the issue

---

## 10. Glossary

**Cable**: A physical electrical conductor connecting two endpoints

**Circuit Number**: A sequential identifier automatically assigned to each cable within a plant type

**Core Configuration**: The arrangement of conductors within a cable (e.g., "3C+E" means 3 cores plus earth)

**Core Size**: The cross-sectional area or gauge of the main conductors in a cable

**Dashboard**: The main control interface of the application with buttons to access all functions

**Earth Size**: The cross-sectional area or gauge of the protective earth/ground conductor

**Endpoint**: A connection point in the electrical system where cables terminate (source or destination)

**Form ID**: An internal identifier used by the application to distinguish between plant types ("WET_PLANT", "ORE_SORTER", "RETREATMENT")

**Insulation Type**: The material used to insulate the cable conductors (e.g., PVC, XLPE)

**Plant Type**: One of three processing facilities (Wet Plant, Ore Sorter, or Retreatment Plant)

**Scheduled**: A status indicator showing that a cable installation is planned but not yet completed

**Short Name**: A unique abbreviated identifier for an endpoint, following plant-specific formatting rules

**Source**: The starting endpoint where a cable originates

**Destination**: The ending endpoint where a cable terminates

**Table**: An Excel ListObject that stores structured data (cables or endpoints) with automatic filtering and formatting

---

## Best Practices

### Endpoint Management
1. **Use descriptive names**: Make descriptions clear and specific (e.g., "Motor Control Center - Crushing Area 1" rather than just "MCC1")
2. **Consistent naming**: Develop a naming convention and stick to it
3. **Pre-register endpoints**: Register all endpoints before starting cable registration for smoother workflow

### Cable Management
1. **Complete information**: Fill in all fields accurately at the time of registration
2. **Verify endpoints**: Double-check that you're selecting the correct source and destination
3. **Update status fields**: Remember to check "ID Attached" once physical tags are applied
4. **Use "Scheduled" properly**: Mark cables as scheduled during planning phase, uncheck when installed

### Data Quality
1. **Regular reviews**: Periodically review your data for accuracy
2. **Backup frequently**: Save copies of the workbook regularly
3. **Single user editing**: Avoid having multiple users editing the same workbook simultaneously
4. **Don't edit tables directly**: Use the provided forms when possible to maintain data integrity

### Workflow Efficiency
1. **Batch endpoint registration**: Register all endpoints for an area before adding cables
2. **Sequential cable entry**: Enter cables in a logical order (by area or circuit)
3. **Use keyboard shortcuts**: Tab to move between form fields
4. **Keep the dashboard visible**: Use it as your home base for navigation

---

## Document Control

**Version History:**

| Version | Date | Author | Changes |
|---------|------|--------|---------|
| 1.0 | Dec 2024 | System Documentation | Initial user guide creation |

**Feedback:**
For suggestions or corrections to this documentation, please contact jorr@mtcarbine.com.au

---

*End of User Guide*
