# ExcelCompare

A small command-line tool that compares two sheets in two Excel files and writes
an .xlsx report that highlights every difference.

## How the comparison works

Each cell at the same row/column is compared on its **formatted string value**
(so `1`, `1.0`, and the formula `=1` all count as the same). The output report
copies the layout of the input and colours every cell that differs:

| Colour       | Meaning                                |
|--------------|----------------------------------------|
| Yellow       | Value changed between file1 and file2  |
| Light green  | Added in file2 (was empty in file1)    |
| Light pink   | Removed from file2 (was in file1 only) |

Each highlighted cell also has a comment with the exact old/new values.
A second sheet called **Summary** contains the totals.

## Build (one-time)

Requirements: [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)

```
cd ExcelCompare
dotnet publish -c Release
```

The final single-file .exe will be at:

```
bin\Release\net8.0\win-x64\publish\ExcelCompare.exe
```

That single file is fully self-contained — recipients do **not** need .NET
installed. Just ship that one `.exe`.

### Build for a different platform

Change `<RuntimeIdentifier>` in `ExcelCompare.csproj` or pass it on the
command line, e.g.:

```
dotnet publish -c Release -r linux-x64
dotnet publish -c Release -r osx-x64
dotnet publish -c Release -r osx-arm64
```

## Usage

```
ExcelCompare.exe <file1.xlsx> <file2.xlsx> [options]
```

### Options

| Option            | Description                                     |
|-------------------|-------------------------------------------------|
| `--sheet1 <name>` | Sheet name in file1 (default: first sheet)      |
| `--sheet2 <name>` | Sheet name in file2 (default: first sheet)      |
| `--sheet  <name>` | Shortcut for setting both `--sheet1` and `--sheet2` |
| `--out <path>`    | Output path (default: `Comparison_<timestamp>.xlsx`) |

### Examples

Compare the first sheet of each workbook:

```
ExcelCompare.exe old.xlsx new.xlsx
```

Compare a specific sheet by name in both files:

```
ExcelCompare.exe old.xlsx new.xlsx --sheet "Q1 Data" --out report.xlsx
```

Compare differently-named sheets:

```
ExcelCompare.exe old.xlsx new.xlsx --sheet1 "Sales_v1" --sheet2 "Sales_v2"
```

## Exit codes

| Code | Meaning                              |
|------|--------------------------------------|
| 0    | Success                              |
| 1    | Bad arguments (usage printed)        |
| 2    | Input file not found                 |
| 99   | Unexpected error (stack trace shown) |
