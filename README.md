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

From the repository root (this folder contains both `ExcelCompare.csproj` and
`Excel_Comparison.sln`, so always pass one of them to `dotnet`):

```powershell
dotnet restore ExcelCompare.csproj
dotnet publish ExcelCompare.csproj -c Release
```

If restore was skipped and you see `NETSDK1004` (missing `obj\project.assets.json`),
run `dotnet restore` first.

The self-contained single-file executable is produced at:

```
bin\Release\net8.0\win-x64\publish\ExcelCompare.exe
```

That build is **self-contained** — recipients do **not** need .NET installed.
Ship that `.exe` alone.

### Build for a different platform

The project defaults to `win-x64` in `ExcelCompare.csproj`. Override at publish time:

```powershell
dotnet publish ExcelCompare.csproj -c Release -r linux-x64
dotnet publish ExcelCompare.csproj -c Release -r osx-x64
dotnet publish ExcelCompare.csproj -c Release -r osx-arm64
```

You can also change `<RuntimeIdentifier>` in `ExcelCompare.csproj` if you always target one OS.

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
