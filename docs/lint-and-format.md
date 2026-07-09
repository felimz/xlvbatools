# xlvbatools — Linter & Formatter Guide

This document describes the offline static linter and formatting rules of `xlvbatools`.

---

## Static Analysis (Linter)

The linter scans VBA code offline (without Excel or COM dependencies) to enforce clean coding practices.

### Lint Rules Reference

| Rule ID | Severity | Description | Rationale |
|:---|:---|:---|:---|
| **DS001** | ERROR | Dim after executable code | VBA allocates all local variables at procedure startup regardless of where they are placed. Declaring them inline is misleading. All `Dim` statements must be placed before any executable code. |
| **LC001** | WARNING | Orphaned line continuation | A line continuation character (` _`) must be preceded by exactly one space to form a valid continuation. |
| **SB001** | ERROR | Unbalanced blocks | A procedure is missing its corresponding closing statement (`End Sub`, `End Function`, `End Property`). |
| **PF001** | WARNING | Interactive MsgBox call | Headless automation runs will hang on blocking `MsgBox` calls. Prefer logging, conditional guards, or `Application.UserControl` checks. |
| **PF002** | WARNING | Implicit Variant (no `As` clause) | Declaring variables without an explicit type defaults to Variant, consuming more memory and hiding bugs. |
| **PF003** | WARNING | ActiveSheet / ActiveCell usage | Focus-dependent references can lead to flaky automation bugs if the active focus shifts during runtime. Prefer explicit workbook/sheet references. |
| **OE001** | WARNING | Option Explicit missing | Missing `Option Explicit` allows undeclared/misspelled variables to compile silently, causing runtime bugs. |
| **BK001** | WARNING | Block-level variable declaration | Declaring variables near-use inside loops (`For`, `Do`) or conditions (`If`, `Select`) falsely implies block scope, which VBA does not support (hoisting). |
| **SD002** | STYLE | Multiple variable declarations on one line | Commas in `Dim` statements are forbidden to prevent typing misconceptions (e.g., in `Dim a, b As Long`, only `b` is a Long; `a` is a Variant). |
| **PF004** | STYLE | Avoid Integer data type | VBA internally promotes Integers (16-bit) to Longs (32-bit) on 32-bit/64-bit systems, which can lead to overflow errors. Use `Long` instead. |
| **SD005** | STYLE | Missing parameter modifier | Arguments should explicitly specify `ByVal` or `ByRef` (instead of relying on default implicit `ByRef` behavior). |
| **SD006** | STYLE | Access modifier missing on procedure | Sub/Function/Property declarations should explicitly state `Public` or `Private` access. |
| **SD010** | STYLE | Line exceeds maximum length limit | Horizontal scrolling degrades code readability in VBE. Statements should be kept under 120 characters. |
| **SD014** | STYLE | Avoid Call keyword | The `Call` keyword is obsolete. Invoke procedures directly (e.g., `MyProc arg1` instead of `Call MyProc(arg1)`). |
| **SC001** | WARNING | Hardcoded secret or credential literal | Plain-text strings matching API keys, passwords, tokens, or hashes present a security vulnerability when committed to version control. |
| **SC002** | WARNING | Absolute file path usage | Absolute paths (e.g., `C:\Users\...`) break portability across developer environments. Prefer relative pathing or environment-derived paths. |
| **CL001** | STYLE | Hungarian type suffixes | Suffixes like `%`, `&`, `$`, `!`, `#`, `@` are obsolete type hints. Use explicit `As` clauses. |
| **CL002** | STYLE | Selection or ActiveWindow dependency | Relying on `.Select`, `Selection`, or `ActiveWindow` is slow and fragile. Reference ranges and sheets directly. |
| **SF001** | WARNING | Silent error suppression | Using `On Error Resume Next` without checking `Err.Number` hides runtime exceptions. Always check `Err.Number` and reset with `On Error GoTo 0`. |
| **DC001** | WARNING | Unused local variable | Declared variables that are never read or written clutter the code and represent dead allocations. |
| **DC002** | WARNING | Empty procedure declaration | Sub or Function bodies containing only comments or whitespace should be removed or implemented. |
| **DC003** | WARNING | Dead procedure (0 incoming calls) | Procedures that are never called from any other part of the project should be deleted or implemented. Call graph parsing is comment-aware, ignoring inline single-quote comments and `Rem` statements to prevent false-positive dependency edges. *(Requires call graph analysis).* |
| **SD015** | STYLE | Multiple consecutive blank lines | Excess blank lines increase vertical scrolling and clutter code readability. |
| **SD016** | STYLE | Double-spaced code blocks | Alternating blank lines after every line of code degrades vertical readability. |
| **RK001** | WARNING | Reserved keyword variable name | Declaring variables or parameters with the same name as a VBA reserved keyword (e.g. `Optional`, `Date`, `Error`, `Next`) leads to compilation errors or unexpected runtime behavior. |
| **IP001** | ERROR | Executable code outside procedure | Executable statements (calls, assignments, loops) must be placed inside a `Sub`, `Function`, or `Property` block. Only declarations are valid at module scope. |
| **DP001** | ERROR | Duplicate public procedure | Declaring multiple public procedures with the same name across different standard modules violates naming uniqueness and causes Excel compile errors. |
| **SM001** | WARNING | Invalid class or typed object member | Referencing a property or method that does not exist on a user-defined class module or a built-in typed object (e.g., calling `.SheetName` instead of `.Name` on a `Worksheet`) will fail at compile/runtime. |
| **CT001** | ERROR | VBE Compile Test failure | *(COM only)* Excel VBE compiler failed to compile the project. Includes code line context. |

---

## VBA Code Formatter

The formatter (`xlvba fmt`) normalizes code layout non-destructively:
* **Indentation:** Adjusts indent sizing (default 4 spaces) for Control Flow statements (`If...Then`, `For...Next`, `Do...Loop`, `Select Case`, `With`).
* **Header Preservation:** Skips indentation of VBE module attributes (e.g. `Attribute VB_Name`) and `Option` statements, maintaining clean column-0 placement.
* **Blank Lines:** Collapses multiple consecutive blank lines into a single blank line.
* **Double-Spacing removal:** Removes alternating double-spaced lines within procedures to keep code readable.
* **Dry-Run Mode:** Compares input vs formatted code and prints a diff without modifying the files.
