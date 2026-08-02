# Excel Date & Time Handling

This guide explains how dates and times are handled in Excel and how to use the helper conversion scripts.

## General Rule for Writing Dates to Excel

When writing dates to Excel via `AgentExcel` range write endpoints, **you usually do NOT need to convert dates to serial numbers**.

Simply pass the date as a formatted text string using standard date formats, preferably `yyyy/MM/dd` or `yyyy/MM/dd HH:mm:ss`:

```json
{
    "workbook": "sample.xlsx",
    "sheet": "Sheet1",
    "range": "A1",
    "values": [["2023/03/15"]]
}
```

Excel COM Interop automatically parses standard date strings like `yyyy/MM/dd` into Excel date cells.

## When Serial Numbers Are Needed

Excel internally stores dates as floating-point serial numbers (where `1` represents `1900-01-01 00:00:00`, calculated using base date `1899-12-30`). The integer portion represents days since the base date, and the decimal portion represents time of day.

Converting between date strings and serial numbers is only needed when:

1. **Inspecting raw values**: Range read APIs might return raw numeric serial values (e.g. `45000.5`) for date-formatted cells.
2. **Explicit Serial Calculation**: You specifically need to query, verify, or calculate the exact underlying Excel serial number.

## Helper Conversion Scripts

Two PowerShell helper scripts are provided in `scripts/` to convert between serial numbers and formatted date strings:

> **Important**: Always use `powershell -ExecutionPolicy Bypass -Command "..."` when passing arrays or pipeline inputs from the command line. Using `-File` treats CLI arguments as literal strings, so expressions like `@(44927, 45000.5)` with spaces get split into invalid positional arguments.

### 1. `ConvertFrom-ExcelSerial.ps1` (Serial Number → Date String)

Converts numeric serial numbers to formatted date strings:

- **Single serial number**:

```powershell
powershell -ExecutionPolicy Bypass -File "scripts\ConvertFrom-ExcelSerial.ps1" -Serial 45000.5
# Output: 45000.5 -> 2023/03/15 12:00:00
```

- **Multiple serial numbers via parameter with custom output format**:

```powershell
powershell -ExecutionPolicy Bypass -Command "& 'scripts\ConvertFrom-ExcelSerial.ps1' -Serial 44927, 45000.5 -Format 'yyyy-MM-dd HH:mm:ss'"
```

- **Multiple serial numbers via pipeline with custom output format**:

```powershell
powershell -ExecutionPolicy Bypass -Command "44927, 45000.5 | & 'scripts\ConvertFrom-ExcelSerial.ps1' -Format 'yyyy-MM-dd HH:mm:ss'"
```

### 2. `ConvertTo-ExcelSerial.ps1` (Date String → Serial Number)

Converts date strings to Excel serial numbers:

- **Single date string**:

```powershell
powershell -ExecutionPolicy Bypass -File "scripts\ConvertTo-ExcelSerial.ps1" -Date "2023/03/15 12:00:00"
# Output: 2023/03/15 12:00:00 -> 45000.5
```

- **Multiple date strings via parameter with custom output format**:

```powershell
powershell -ExecutionPolicy Bypass -Command "& 'scripts\ConvertTo-ExcelSerial.ps1' -Date '2023-01-01 00:00:00', '2023-03-15 12:00:00'"
```

- **Multiple date strings via pipeline with custom input format**:

```powershell
powershell -ExecutionPolicy Bypass -Command "'2023-01-01', '2023-03-15' | & 'scripts\ConvertTo-ExcelSerial.ps1' -Format 'yyyy-MM-dd'"
```

