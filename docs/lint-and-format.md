# xlvbatools — Linter & Formatter Guide

This document describes the offline static linter and formatting rules of `xlvbatools`.

---

## Static Analysis (Linter)

The linter scans VBA code offline (without Excel or COM dependencies) to enforce clean coding practices.

### Lint Rules Reference

| Rule ID | Severity | Description | Rationale |
|:---|:---|:---|:---|
| **PF001** | WARNING | Executable statements before `Dim` | All variables should be declared at the top of the procedure for readability. |
| **PF002** | WARNING | Implicit Variant (no `As` clause) | Declaring variables without an explicit type defaults to Variant, consuming more memory and hiding bugs. |
| **PF003** | WARNING | ActiveSheet / ActiveCell usage | Fragile and prone to execution failure. Prefer explicit workbook/sheet object references. |
| **OE001** | WARNING | Option Explicit missing | Missing `Option Explicit` allows undeclared/misspelled variables to pass compilation, causing runtime bugs. |
| **LC001** | WARNING | Missing line continuation space | A line continuation character (`_`) must be preceded by exactly one space to form a valid continuation. |
| **UB001** | WARNING | Unbalanced Sub/Function blocks | A procedure is missing its corresponding closing statement (`End Sub`, `End Function`). |
| **MB001** | WARNING | Interactive MsgBox call | Headless automation runs will hang on blocking MsgBox calls. Prefer logging or conditional guards. |
| **CT001** | ERROR | VBE Compile Test failure | *(COM only)* Excel VBE compiler failed to compile the project. Includes code line context. |

---

## VBA Code Formatter

The linter formatter (`xlvba fmt`) normalizes code layout non-destructively:
* **Indentation:** Adjusts indent sizing (default 4 spaces) for Control Flow statements (`If...Then`, `For...Next`, `Do...Loop`, `Select Case`, `With`).
* **Header Preservation:** Skips indentation of VBE module attributes (e.g. `Attribute VB_Name`) and `Option` statements, maintaining clean column-0 placement.
* **Blank Lines:** Collapses multiple consecutive blank lines into a single blank line.
* **Dry-Run Mode:** Compares input vs formatted code and prints a diff without modifying the files.
