# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

AIDD XVBA Agents is a VBA-focused AI-Driven Development project that uses Claude Code and custom agents to automate Excel VBA application development from requirements to implementation. The core feature is the XVBA Mock Creator agent that generates complete VBA implementations from specifications.

## Core Architecture

### Dual Encoding System
The project uses a dual encoding approach for VBA development:
- **Development files**: `customize/vba-files/` - UTF-8 encoding (version control friendly)
- **Production files**: `vba-files/` - Shift-JIS encoding (required by Excel VBA)

The conversion between these is handled by `xvba_pre_export.ps1` and must be run before Excel import.

### Directory Structure
```
customize/vba-files/    # UTF-8 source files (edit these)
├── Module/             # Standard modules (mod*.bas)
└── Class/              # Class modules (*.cls, Sheet*.cls, ThisWorkbook.cls)

vba-files/              # Shift-JIS production files (auto-generated, do not edit)
├── Module/
└── Class/

xvba_modules/           # Installed XVBA packages (Xdebug, excel-types)
```

### Configuration Files
- **config.json**: Project configuration (app_name, excel_file, vba_folder, xvba_packages)
- **package.json**: NPM-style package management for XVBA dependencies
- **basefile.xlsm**: Template workbook that gets copied to the target Excel file

## Essential Commands

### Build and Deploy
```powershell
.\xvba_pre_export.ps1
```
This script performs:
1. Converts UTF-8 files from `customize/vba-files/` to Shift-JIS in `vba-files/`
2. Copies `basefile.xlsm` to the target Excel filename specified in config.json
3. Prepares all VBA files for Excel import

### VBA Module Export to Excel
```bash
xvba-macro list
```
Batch imports all VBA modules from `vba-files/` into the Excel workbook.

## VBA Development Guidelines

### Module Architecture
Standard module naming and organization pattern:
- **modConstants.bas**: System constants (sheet names, colors, fonts, indices)
- **modCmn.bas**: Common utilities (pre-existing, use these functions)
- **modData.bas**: Data access layer
- **modBusiness.bas**: Business logic
- **modUI.bas**: UI operations and formatting
- Feature-specific modules as needed

### Sheet Management Strategy
**Critical**: Never create new sheets. Instead:
1. Rename existing Sheet1-9 from basefile.xlsm
2. Define sheet name constants in modConstants.bas
3. Define sheet index constants (e.g., `SHEET_INDEX_DASHBOARD = 1`)
4. **Always access sheets by index** using `modCmn.GetWorksheetByIndex(index)`

Example pattern:
```vba
' In modConstants.bas
Public Const SHEET_DASHBOARD As String = "Dashboard"
Public Const SHEET_INDEX_DASHBOARD As Integer = 1

' In other modules
Set ws = modCmn.GetWorksheetByIndex(SHEET_INDEX_DASHBOARD)
```

### modCmn.bas Common Functions
The project includes a comprehensive common utilities module (`customize/vba-files/Module/modCmn.bas`). Key functions to use:
- `GetWorksheet(sheetName)`: Safe worksheet retrieval by name
- `GetWorksheetByIndex(sheetIndex)`: Safe worksheet retrieval by index (preferred)
- `GetTable(ws, tableName)`: Safe ListObject retrieval
- `TableExists(ws, tableName)`: Check table existence
- `LogError(funcName, message)`: External error logging
- Font/formatting functions for UI consistency
- String utilities, validation functions

Always leverage these existing utilities rather than reimplementing.

### UI Implementation Rules
- **All user interactions via buttons**: Registration, search, update, delete, import, export
- **No cell change events (Worksheet_Change)**: Prohibited
- **No sheet selection events**: Prohibited
- All functionality must be triggered by button clicks

### Workbook Event Restrictions
- **Only Workbook_Open() is allowed**
- **Workbook_BeforeClose() is PROHIBITED**
- **Workbook_BeforeSave() is PROHIBITED**
These restrictions prevent infinite loops and save/close event conflicts.

### Mandatory VBA Standards
Every VBA function/subroutine must include:
1. Error handling with `On Error GoTo ErrHandler`
2. Error logging using `LogError(functionName, errorMessage)`
3. Proper cleanup in error handler
4. Exit Function/Sub before ErrHandler label

## XVBA Mock Creator Agent

The primary workflow uses the custom agent defined in `.claude/agents/xvba-mock-creator.md`.

### Invocation
```
@xvba-mock-creator <specification description>
```

### Agent Workflow
1. Analyzes specification documents (design.md, specification.md, etc.)
2. Creates project structure (config.json, package.json)
3. Implements all required VBA modules in `customize/vba-files/`
4. Ensures all features from specification are fully implemented
5. Provides implementation guide

### Implementation Completion Criteria
The agent must complete ALL of these before finishing:
- [ ] All specification features implemented in VBA code
- [ ] No syntax errors, code is runnable
- [ ] Sheet structure, buttons, and formatting complete
- [ ] Error handling and logging implemented
- [ ] Specification requirements verification table created
- [ ] Test cases executed and bugs fixed

**Critical**: The agent must not stop at design phase - it must produce working VBA code.

## File Editing Workflow

When modifying VBA code:
1. **Edit only files in `customize/vba-files/`** (UTF-8 encoded)
2. Run `.\xvba_pre_export.ps1` to convert to Shift-JIS
3. Use `xvba-macro list` to import into Excel
4. Test in Excel
5. Never manually edit files in `vba-files/` directory

## Quality Requirements

### Security Checklist
- [ ] Only Workbook_Open() event used
- [ ] All user actions via buttons (no cell/sheet events)
- [ ] No BeforeClose/BeforeSave events
- [ ] No recursive calls
- [ ] Defensive programming practices

### Implementation Checklist
- [ ] LogError() for external logging
- [ ] Error handlers in all functions
- [ ] Sheet access via index constants
- [ ] modCmn.bas utilities utilized
- [ ] Unified font and color schemes

## Button Macro References

When assigning button OnAction properties:
```vba
' For sheet class procedures
btn.OnAction = "Sheet1.ProcedureName"

' For standard module procedures
btn.OnAction = "ModuleName.ProcedureName"

' Called procedure must be Public
Public Sub ProcedureName()
    ' implementation
End Sub
```

## System Initialization Pattern

Typical Workbook_Open implementation:
```vba
Private Sub Workbook_Open()
    Call InitializeSystem
    Call CreateInitialSampleData  ' Only on first run
    Call ShowSplashScreen
End Sub

Private Sub CreateInitialSampleData()
    If IsDataEmpty() Then
        Call CreateSampleData  ' From standard module
    End If
End Sub
```
