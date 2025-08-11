# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

XVBA Mock Creator is an Excel VBA development framework that provides a modern development environment for Excel VBA applications. It uses a dual-encoding system with UTF-8 for development and Shift-JIS for Excel compatibility, along with NPM-style package management for VBA modules.

## Key Commands

### Build and Export Process
```powershell
.\xvba_pre_export.ps1
```
This is the main build command that:
- Converts UTF-8 source files from `customize/vba-files/` to Shift-JIS in `vba-files/`
- Copies `basefile.xlsm` to the target Excel file specified in `config.json`
- Prepares all VBA files for Excel import

### VBA Module Management
```bash
xvba-macro list
```
Imports all VBA modules into Excel workbook (requires @localsmart/xvba-cli)

## Architecture Overview

### Dual-Source File System
The project maintains two parallel VBA file structures:
- **`customize/vba-files/`**: UTF-8 encoded source files for development
- **`vba-files/`**: Shift-JIS encoded files for Excel import (auto-generated)

### Package Management
- **`package.json`**: NPM-style dependencies for XVBA packages
- **`xvba_modules/`**: Installed XVBA packages (similar to node_modules)
- **Built-in packages**:
  - `Xdebug`: VS Code debugging utilities (`Xdebug.printx`, `Xdebug.printError`)
  - `excel-types`: TypeScript-style type definitions for Excel VBA objects

### Configuration System
- **`config.json`**: Main project configuration
  - `excel_file`: Target Excel workbook filename
  - `vba_folder`: VBA source directory (default: "vba-files")
  - `xvba_packages`: Package dependencies
- **`basefile.xlsm`**: Template Excel workbook copied to target file

## Development Workflow

1. **Development**: Edit VBA source files in `customize/vba-files/` (UTF-8 encoding)
2. **Build**: Run `.\xvba_pre_export.ps1` to convert and prepare files
3. **Import**: Use `xvba-macro list` to import VBA modules into Excel
4. **Test**: Test and debug in Excel environment

## File Structure Patterns

### VBA Module Organization
```
customize/vba-files/
├── Class/          # VBA Class Modules (.cls files)
│   ├── Sheet1.cls
│   ├── Sheet2.cls
│   └── ThisWorkbook.cls
└── Module/         # VBA Standard Modules (.bas files)
    └── [module files]
```

### Package Structure
```
xvba_modules/
├── Xdebug/
│   ├── Xdebug.cls
│   ├── xvba.package.json
│   └── README.md
└── excel-types/
    ├── *.d.vb (type definition files)
    └── xvba.package.json
```

## VBA Development Patterns

### Button Event Handling
```vba
' Reference sheet class procedures
btn.OnAction = "Sheet1.ProcedureName"

' Reference standard module procedures  
btn.OnAction = "ModuleName.ProcedureName"

' Procedures must be Public
Public Sub ProcedureName()
    ' Implementation
End Sub
```

### System Initialization
```vba
Private Sub Workbook_Open()
    Call InitializeSystem
    Call CreateInitialSampleData  ' First run only
    Call ShowSplashScreen
End Sub
```

## Encoding Requirements

Always work with UTF-8 files in `customize/vba-files/` during development. The build process handles Shift-JIS conversion automatically. Never manually edit files in the `vba-files/` directory as they will be overwritten.

## Dependencies

- Windows PowerShell (for build script)
- Microsoft Excel (.xlsm support)
- `@localsmart/xvba-cli` (for VBA import functionality)