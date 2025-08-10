# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an XVBA (Extended VBA) mock creator project designed to build Excel VBA applications with modern development tooling. XVBA extends traditional VBA development with package management, type definitions, and enhanced debugging capabilities.

## Key Commands

### Pre-Export Preparation
```bash
.\xvba_pre_export.ps1
```
This PowerShell script is the primary build command that:
- Converts UTF-8 source files from `customize/vba-files/` to Shift-JIS encoding in `vba-files/`
- Copies `basefile.xlsm` to the configured Excel filename (from config.json)
- Prepares VBA files for Excel import

### Package Management
```bash
npx xvba install <package-name>
```
Install XVBA packages from the xvba.dev repository.

## Project Architecture

### Dual Directory Structure
The project maintains two parallel VBA directory structures:

1. **`customize/vba-files/`** - UTF-8 encoded source files for development
   - Edit VBA code here in UTF-8 for better version control
   - Contains Class/ and Module/ subdirectories

2. **`vba-files/`** - Shift-JIS encoded files for Excel export
   - Generated automatically by xvba_pre_export.ps1
   - Used for importing into Excel workbooks
   - Should not be edited directly

### Configuration Files

- **`config.json`** - Main project configuration
  - `excel_file`: Target Excel workbook filename
  - `vba_folder`: Points to the vba-files directory
  - `xvba_packages`: Installed XVBA package dependencies
  - Application metadata (name, description, author)

- **`package.json`** - NPM-style dependency management for XVBA packages

### XVBA Modules System

Located in `xvba_modules/`, this contains installed packages:

- **`Xdebug/`** - VBA debugging utilities that output to VSCode
  - Provides `Xdebug.printx` for variable debugging
  - Provides `Xdebug.printError` for error handling

- **`excel-types/`** - TypeScript-style definitions for Excel VBA objects
  - Enables IntelliSense/autocomplete for Excel objects
  - Files use `.d.vb` extension for type definitions

### File Structure Conventions

- **`.cls` files** - VBA Class modules (Sheet classes, ThisWorkbook, custom classes)
- **`.bas` files** - VBA Standard modules using modular structure:
  - **`modConstants.bas`** - System constants (table names, sheet names, status values)
  - **`modData.bas`** - Data access layer (GetTable, LogError, LogAudit functions)
  - **`modBusiness.bas`** - Business logic (validation, calculations, workflows)
  - **`modUI.bas`** - UI operations (form display, report generation, screen control)
- **`.frm` files** - UserForm files (if present)

### Modular VBA Architecture

**IMPORTANT**: This project uses a modular VBA architecture to avoid code bloat and VBA import limitations:

- **Benefits**: Better maintainability, easier VBA imports, improved code organization
- **Reference pattern**: Use `modModuleName.FunctionName` when calling functions across modules

## Development Workflow

1. Edit VBA source code in `customize/vba-files/` (UTF-8 encoding)
   - Use modular structure: organize code into mod*.bas files
   - Avoid creating large monolithic files
2. Run `.\xvba_pre_export.ps1` to prepare files for Excel
3. Import the generated Shift-JIS files from `vba-files/` into Excel
   - Import each mod*.bas file separately in VBA Editor
   - Smaller modular files prevent VBA import errors
4. The script automatically prepares the Excel workbook file based on config.json

## Excel Integration

- **Base File**: `basefile.xlsm` serves as the template workbook
- **Target File**: Configured via `config.json` `excel_file` property
- **Import Process**: Use Excel's VBA Editor to import modules from the `vba-files/` directory

## Encoding Handling

Critical aspect: VBA files must be in Shift-JIS encoding for proper Excel import, but UTF-8 is preferred for version control and editing. The pre-export script handles this conversion automatically.

## VBA Import Best Practices

### Avoiding Common Import Issues

1. **File Size Limitations**: VBA Editor has limitations importing large files
   - **Solution**: Use modular architecture with mod*.bas files

2. **Line Continuation Limits**: VBA has restrictions on consecutive line continuation characters (_)
   - **Limit**: Approximately 25-30 line continuations per statement
   - **Solution**: Break large Array definitions into helper function calls

3. **Memory Constraints**: Large sample data arrays can cause import failures
   - **Solution**: Use iterative data creation with helper functions instead of massive Array literals

### Current Project Structure

This project implements the modular architecture with:
- `modConstants.bas` (39 lines) - System constants
- `modData.bas` (153 lines) - Data access functions  
- `modBusiness.bas` (328 lines) - Business logic
- All sheet classes updated to reference the modular functions

This structure resolves VBA import size limitations while maintaining code organization and functionality.