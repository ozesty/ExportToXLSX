# Export-ToXLSX

`Export-ToXLSX` is a pipeline-friendly PowerShell function that creates a real
`.xlsx` workbook without Microsoft Excel, COM automation, or third-party
PowerShell modules.

It is designed for commands such as `Get-ADUser`, but it accepts any PowerShell
objects.

## Requirements

- Windows PowerShell 5.1
- .NET's built-in `System.IO.Compression` assembly
- No installed copy of Excel
- No third-party PowerShell modules

`Get-ADUser` requires Microsoft's Active Directory module, but that module is
not required by the exporter itself.

## Load the function

Clone or download this repository, then dot-source the script:

```powershell
. .\Export-ToXLSX.ps1
```

The `Export-ToXLSX` command is then available in the current PowerShell session.

## Basic usage

Pipe one or more objects to the function and provide an output path:

```powershell
Get-ADUser jsmith |
    Export-ToXLSX -Path 'X:\temp\jsmith.xlsx'
```

Without `-Property`, the function automatically creates a column for every
property it encounters. Properties first encountered on later objects are also
added as columns.

## Choose the columns

Use `Select-Object` to control exactly what enters the exporter:

```powershell
Get-ADUser -Filter * -Properties Mail, Department, Title |
    Select-Object SamAccountName, Name, Mail, Department, Title, Enabled |
    Export-ToXLSX -Path 'X:\temp\users.xlsx'
```

Alternatively, use the exporter's `-Property` parameter. The property order
becomes the spreadsheet column order:

```powershell
Get-ADUser -Filter * -Properties Mail, Department, Title |
    Export-ToXLSX -Path 'X:\temp\users.xlsx' `
        -Property SamAccountName, Name, Mail, Department, Title, Enabled
```

`Get-ADUser -Properties` retrieves additional attributes from Active Directory;
it does not restrict which properties are displayed. Use `Select-Object` or the
exporter's `-Property` parameter when you need an exact column list.

Calculated properties work as expected:

```powershell
Get-ADUser -Filter * |
    Select-Object Name, SamAccountName,
        @{Name='ExportedOn'; Expression={ Get-Date }} |
    Export-ToXLSX -Path 'X:\temp\users.xlsx'
```

The function is not specific to Active Directory:

```powershell
Get-Service |
    Select-Object Name, DisplayName, Status |
    Export-ToXLSX -Path '.\services.xlsx'
```

## How it works

An `.xlsx` file is an Open XML package containing XML documents inside a ZIP
container. The function builds the required workbook, worksheet, relationship,
content-type, and style parts using built-in .NET classes.

While pipeline objects are processed, finished worksheet rows are streamed to a
temporary buffer. At the end of the pipeline, the function writes the header,
copies the buffered rows into the worksheet, finalizes the ZIP package, and then
replaces the requested output file. This avoids keeping the full data set and
worksheet XML in memory.

The web-style namespace values inside the script are Open XML identifiers. They
do not cause network requests.

## Value handling

- Strings are XML-escaped and stored as inline strings.
- Strings beginning with `=` remain text and are not executed as formulas.
- Booleans are stored as Excel boolean cells.
- Safely representable numbers are stored as numeric cells.
- Integers, decimals, or floating-point values that Excel cannot preserve
  exactly are stored as text instead of being silently rounded.
- Dates from 1 March 1900 onward are stored as Excel dates.
- Earlier dates and `DateTimeOffset` values are stored as ISO-formatted text.
- Collections are joined with `; `.
- Byte arrays are Base64-encoded.
- XML control characters and literal `_xHHHH_` sequences are preserved using
  SpreadsheetML escaping.
- The header row is bold, and native date cells use a date/time format.

## Empty results and existing files

When `-Property` is supplied and the input is empty, the function creates a
header-only workbook. Without `-Property`, an empty input creates an empty
worksheet.

The workbook is first built in a temporary file beside the destination. An
existing output is replaced only after the new workbook has been completed. If
generation fails, the previous output remains intact. Filesystems that do not
support atomic replacement fail safely instead of deleting the existing file.

The destination directory is created automatically when it does not exist.

## Excel limits

The function validates Excel's worksheet limits:

- 1,048,575 data rows plus one header row
- 16,384 columns
- 32,767 characters per cell

Exceeding a limit stops the export with an explanatory error.

## Notes

- The output path must end in `.xlsx`.
- A `-Property` name that does not exist on any input object causes an error.
- Multivalued properties are flattened to text and cannot be reconstructed as
  their original collection from the workbook.
- In Windows PowerShell 5.1, cancelling an upstream pipeline with `Ctrl+C` can
  leave its delete-on-close temporary buffer locked until garbage collection or
  the PowerShell process exits. The destination workbook is not replaced.
