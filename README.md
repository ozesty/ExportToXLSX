# Export-ToXLSX

`Export-ToXLSX` is a pipeline-friendly PowerShell function that creates a real
`.xlsx` workbook without Microsoft Excel, COM automation, or third-party
PowerShell modules.

It is designed for commands such as `Get-ADUser`, but it accepts any PowerShell
objects.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
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

You can also pass an array directly via `-InputObject`:

```powershell
$users = Get-ADUser -Filter *
Export-ToXLSX -Path 'X:\temp\users.xlsx' -InputObject $users
```

Without `-Property`, the function automatically creates a column for every
property it encounters. Properties first encountered on later objects are also
added as columns. For Active Directory objects (`ADEntity`, such as `ADUser`),
internal wrapper properties (`PropertyNames`, `AddedProperties`,
`RemovedProperties`, `ModifiedProperties`, and `PropertyCount`) are skipped
automatically unless explicitly requested via `-Property`.

## Parameters

- `-Path` (Mandatory, Position 0): Output `.xlsx` file path on the `FileSystem`
  provider.
- `-InputObject` (Mandatory, Pipeline): One or more objects to export.
- `-Property` (Alias: `Properties`): Explicit ordered list of property names to
  export as columns.
- `-ExcludeProperty` (Alias: `ExcludeProperties`): Property names or wildcard
  patterns (case-insensitive, e.g. `SID`, `Object*`) to exclude from the
  exported columns.
- `-WorksheetName`: Custom worksheet name (default: `Export`, 1–31 characters,
  validated against Excel naming rules).
- `-NoClobber`: Prevents overwriting an existing `.xlsx` file.
- `-TruncateLongValues`: Truncates cell values exceeding Excel's 32,767-character
  limit (with a warning, preserving UTF-16 surrogate pairs) instead of failing
  the export. Useful for large AD attributes such as `thumbnailPhoto`,
  `userCertificate`, or large `MemberOf` lists.
- `-Strict`: Throws an error if any property specified in `-Property` is not
  present on any input object (by default, missing properties emit a warning and
  remain blank columns).
- `-BoldHeader`: Styles the header row (row 1) in bold. Off by default.
- `-AutoFilter` (Alias: `Sortable`): Enables Excel AutoFilter sort and filter
  dropdowns on the header row. Off by default.
- `-FreezeHeader`: Freezes the header row (row 1) while scrolling. Off by
  default.
- `-WhatIf` / `-Confirm`: Supported via `SupportsShouldProcess`.

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
it does not restrict which properties are displayed and does not preserve the
order of the attribute names you pass to it. Use `Select-Object` or the
exporter's `-Property` parameter when you need an exact column list and order
(the exporter preserves the exact property order from `-Property`,
`Select-Object`, or incoming `PSCustomObject`s).

To keep automatic column discovery while omitting a few default attributes
(such as `SID`, `ObjectGUID`, or `ObjectClass`), use `-ExcludeProperty`
(wildcards are supported):

```powershell
Get-ADUser -Filter * -Properties Mail, Department |
    Export-ToXLSX -Path 'X:\temp\users.xlsx' `
        -ExcludeProperty SID, Object*
```

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
    Export-ToXLSX -Path '.\services.xlsx' -WorksheetName 'Services'
```

## How it works

An `.xlsx` file is an Open XML package containing XML documents inside a ZIP
container. The function builds the required workbook, worksheet, relationship,
content-type, and style parts using built-in .NET classes.

While pipeline objects are processed, finished worksheet rows are streamed to a
temporary buffer. At the end of the pipeline, the function writes the worksheet
metadata (dimensions, optional frozen header pane via `-FreezeHeader`,
auto-sized column widths), header row, copies the buffered rows into the
worksheet, optionally adds an AutoFilter (`-AutoFilter`), finalizes the ZIP
package, and atomically replaces the requested output file. This avoids keeping
the full data set and worksheet XML in memory.

The web-style namespace values inside the script are Open XML identifiers. They
do not cause network requests.

## Value handling

- Strings are XML-escaped and stored as inline strings.
- Strings beginning with `=` remain text and are not executed as formulas.
- Booleans are stored as Excel boolean cells.
- Integers (`SByte`–`UInt32`), `Single`, and in-range `Double` values are stored
  as numeric cells using exact round-trip formatting (`G15`/`G17` for `Double`).
- `Int64`, `UInt64`, and `Decimal` values with more than 15 significant digits
  (such as 18-digit AD FileTimes like `lastLogon` or `accountExpires`), as well
  as out-of-range/non-finite floating-point values, are stored as text instead
  of being silently rounded.
- Dates from 1 March 1900 through 31 December 9999 are stored as Excel dates.
- Earlier/later dates and `DateTimeOffset` values are stored as ISO-8601 text.
- Collections are joined with `; ` (with nested `DateTime` and `DateTimeOffset`
  items formatted as ISO-8601, and nested `byte[]` items Base64-encoded).
- Empty collections and `$null`/`DBNull` values are omitted so Excel treats them
  as blank cells.
- Byte arrays are Base64-encoded.
- XML control characters and literal `_xHHHH_` sequences are preserved using
  SpreadsheetML escaping.
- Columns are auto-sized (clamped to 10–60 characters so `DateTime` values do
  not display as `########`), native date cells use `yyyy-mm-dd hh:mm:ss`, and
  bold headers (`-BoldHeader`), frozen headers (`-FreezeHeader`), and
  sort/filter dropdowns (`-AutoFilter` / `-Sortable`) can be enabled on demand.

## Empty results and existing files

When `-Property` is supplied and the input is empty, the function creates a
header-only workbook. Without `-Property`, an empty input creates an empty
worksheet.

The workbook is first built in a temporary file beside the destination. An
existing output is replaced atomically only after the new workbook has been
completed. If generation fails, the previous output remains intact. Filesystems
that do not support atomic replacement fail safely instead of deleting the
existing file.

The destination directory is created automatically when it does not exist.

## Excel limits

The function validates Excel's worksheet limits:

- 1,048,575 data rows plus one header row
- 16,384 columns
- 32,767 characters per cell

Exceeding row or column limits stops the export with an explanatory error.
Exceeding the per-cell character limit stops the export by default, or truncates
the value with a warning when `-TruncateLongValues` is specified.

## Notes

- The output path must use the `FileSystem` provider and end in `.xlsx`.
- By default, a `-Property` name not present on any input object emits a warning
  and produces a blank column (matching `Get-ADUser`, which omits unset extended
  attributes). Pass `-Strict` to throw an error instead.
- Multivalued properties are flattened to text and cannot be reconstructed as
  their original collection from the workbook.
- In Windows PowerShell 5.1, cancelling an upstream pipeline with `Ctrl+C` can
  leave its delete-on-close temporary buffer locked until garbage collection or
  the PowerShell process exits. The destination workbook is not replaced.

## Changelog

### 2026-11-25

- **AD & data handling improvements**
  - Changed missing `-Property` handling to warn and output blank columns by
    default so `Get-ADUser` queries with unset extended attributes do not abort;
    added `-Strict` to retain throwing behavior.
  - Fixed `[double]` export so fractional values (`1/3`, `[math]::PI`, etc.)
    remain numeric cells instead of falling back to text, while keeping
    out-of-range doubles and >15-digit `Int64`/`UInt64`/`Decimal` values as text.
  - Added upper `DateTime` bound (`9999-12-31T23:59:59`), ISO-8601 formatting
    for `DateTime` items inside collections, and omission of null/empty cells.
  - Added `-TruncateLongValues` with UTF-16 surrogate-safe truncation for large
    values such as `thumbnailPhoto` or large `MemberOf` lists.
  - Updated `$InputObject` to `[AllowNull()][psobject[]]` to support direct
    array arguments as well as pipeline input, skipped non-gettable properties,
    and automatically excluded internal AD wrapper properties (`PropertyNames`,
    `AddedProperties`, `RemovedProperties`, `ModifiedProperties`,
    `PropertyCount`) when `-Property` is omitted.
- **Excel UX & CLI enhancements**
  - Added `-ExcludeProperty` (alias `ExcludeProperties`) with case-insensitive
    exact and wildcard matching to drop columns such as `SID` or `Object*`
    while preserving automatic column discovery.
  - Added auto-sized column widths (`<cols>`), worksheet `<dimension>`, and
    opt-in `-BoldHeader`, `-FreezeHeader`, and `-AutoFilter` (alias `Sortable`)
    switches.
  - Added `-WorksheetName` (with Excel naming validation), `-NoClobber`, and
    `SupportsShouldProcess` (`-WhatIf`/`-Confirm`).
  - Added upfront `FileSystem` provider, directory-collision, and root-drive
    validation, XML declarations on all OOXML parts, and backup-free atomic
    replacement via `[IO.File]::Replace(..., [NullString]::Value)`.
- **Performance & code quality**
  - Added a compiled regex fast path in `ConvertTo-XmlString`, cached column
    names, and reused per-row buffers.
  - Cleaned up all PSScriptAnalyzer diagnostics (0 errors/warnings across
    Default, `PSGallery`, `CodeFormatting`, and PS 5.1/7.x compatibility rules),
    removed `'@)` here-strings and backtick continuations, switched helper calls
    to named parameters, and eliminated unused variables.
