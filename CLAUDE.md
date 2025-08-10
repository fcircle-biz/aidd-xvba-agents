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
- **`.bas` files** - VBA Standard modules  
- **`.frm` files** - UserForm files (if present)

## Development Workflow

1. Edit VBA source code in `customize/vba-files/` (UTF-8 encoding)
2. Run `.\xvba_pre_export.ps1` to prepare files for Excel
3. Import the generated Shift-JIS files from `vba-files/` into Excel
4. The script automatically prepares the Excel workbook file based on config.json

## Excel Integration

- **Base File**: `basefile.xlsm` serves as the template workbook
- **Target File**: Configured via `config.json` `excel_file` property
- **Import Process**: Use Excel's VBA Editor to import modules from the `vba-files/` directory

## Encoding Handling

Critical aspect: VBA files must be in Shift-JIS encoding for proper Excel import, but UTF-8 is preferred for version control and editing. The pre-export script handles this conversion automatically.