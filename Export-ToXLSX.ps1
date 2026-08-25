function Export-ToXLSX {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [psobject]$InputObject,

        [Parameter(Mandatory, Position = 0)]
        [string]$Path,

        [Alias('Properties')]
        [string[]]$Property
    )

    begin {
        $maxDataRows = 1048575
        $maxColumns = 16384
        $maxCellCharacters = 32767
        $culture = [Globalization.CultureInfo]::InvariantCulture
        $utf8 = [Text.UTF8Encoding]::new($false)

        Add-Type -AssemblyName System.IO.Compression

        $resolvedPath = $ExecutionContext.SessionState.Path.
            GetUnresolvedProviderPathFromPSPath($Path)
        if ([IO.Path]::GetExtension($resolvedPath) -ine '.xlsx') {
            throw "Output path must use the .xlsx extension: $resolvedPath"
        }
        $directory = [IO.Path]::GetDirectoryName($resolvedPath)

        $columns = [Collections.Generic.List[string]]::new()
        $columnMap = [Collections.Generic.Dictionary[string, int]]::new(
            [StringComparer]::OrdinalIgnoreCase
        )
        $foundProperties = [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase
        )

        $explicitProperties = $PSBoundParameters.ContainsKey('Property')
        if ($explicitProperties) {
            if (-not $Property -or -not $Property.Count) {
                throw 'Property must contain at least one property name.'
            }
            foreach ($name in $Property) {
                if ([string]::IsNullOrWhiteSpace($name)) {
                    throw 'Property names cannot be empty.'
                }
                if ($columnMap.ContainsKey($name)) {
                    throw "Duplicate property '$name'."
                }
                $columnMap.Add($name, $columns.Count)
                $columns.Add($name)
            }
            if ($columns.Count -gt $maxColumns) {
                throw "Excel supports at most $maxColumns columns."
            }
        }

        function Close-RowBuffer {
            foreach ($resource in @($rowWriter, $rowStream)) {
                if ($resource) {
                    try { $resource.Dispose() }
                    catch { Write-Warning "Could not close an export buffer: $_" }
                }
            }
        }

        function ConvertTo-XmlString {
            param(
                [AllowNull()][string]$Text,
                [int]$Row,
                [string]$Column
            )

            if ($null -eq $Text) { $Text = '' }
            if ($Text.Length -gt $maxCellCharacters) {
                throw "Cell $Column$row exceeds Excel's $maxCellCharacters-character limit."
            }

            # Protect literal SpreadsheetML escape-looking text before creating
            # escapes for characters XML cannot represent directly.
            $Text = [regex]::Replace(
                $Text,
                '_(?=x[0-9A-Fa-f]{4}_)',
                '_x005F_'
            )

            $clean = [Text.StringBuilder]::new($Text.Length)
            for ($index = 0; $index -lt $Text.Length; $index++) {
                $character = $Text[$index]

                if ([char]::IsHighSurrogate($character)) {
                    if (
                        $index + 1 -lt $Text.Length -and
                        [char]::IsLowSurrogate($Text[$index + 1])
                    ) {
                        [void]$clean.Append($character)
                        $index++
                        [void]$clean.Append($Text[$index])
                    }
                    else {
                        [void]$clean.Append([char]0xFFFD)
                    }
                }
                elseif ([char]::IsLowSurrogate($character)) {
                    [void]$clean.Append([char]0xFFFD)
                }
                elseif ($character -eq "`r" -or -not [Xml.XmlConvert]::IsXmlChar($character)) {
                    [void]$clean.Append('_x{0:X4}_' -f [int]$character)
                }
                else {
                    [void]$clean.Append($character)
                }
            }

            [Security.SecurityElement]::Escape($clean.ToString())
        }

        function ConvertTo-TextValue {
            param($Value)

            if ($null -eq $Value -or $Value -is [DBNull]) { return '' }
            if ($Value -is [byte[]]) {
                return [Convert]::ToBase64String($Value)
            }
            if ($Value -is [DateTimeOffset]) {
                return $Value.ToString('o', $culture)
            }
            if ($Value -is [Collections.IEnumerable] -and $Value -isnot [string]) {
                $parts = foreach ($item in $Value) {
                    if ($null -eq $item) { '' }
                    elseif ($item -is [byte[]]) {
                        [Convert]::ToBase64String($item)
                    }
                    elseif ($item -is [DateTimeOffset]) {
                        $item.ToString('o', $culture)
                    }
                    else { [string]$item }
                }
                return [string]($parts -join '; ')
            }
            [string]$Value
        }

        function Get-ColumnName {
            param([int]$Number)

            $name = ''
            do {
                $name = [char](65 + $Number % 26) + $name
                $Number = [math]::Floor($Number / 26) - 1
            }
            while ($Number -ge 0)
            $name
        }

        function Get-InlineStringCell {
            param(
                [string]$Reference,
                $Value,
                [int]$Row,
                [string]$Column,
                [int]$Style = 0
            )

            $text = ConvertTo-XmlString (ConvertTo-TextValue $Value) $Row $Column
            $styleAttribute = if ($Style) { ' s="{0}"' -f $Style } else { '' }
            '<c r="{0}"{1} t="inlineStr"><is><t xml:space="preserve">{2}</t></is></c>' -f
                $Reference, $styleAttribute, $text
        }

        function Get-SignificantDigitCount {
            param([string]$Number)

            $mantissa = ($Number -split '[eE]', 2)[0]
            $digits = ($mantissa -replace '[^0-9]', '').TrimStart('0')
            $digits.Length
        }

        function Get-CellXml {
            param(
                [string]$Reference,
                $Value,
                [int]$Row,
                [string]$Column
            )

            if ($null -eq $Value -or $Value -is [DBNull]) {
                return '<c r="{0}"/>' -f $Reference
            }
            if ($Value -is [bool]) {
                return '<c r="{0}" t="b"><v>{1}</v></c>' -f
                    $Reference, [int]$Value
            }
            if ($Value -is [DateTime]) {
                if ($Value -lt [DateTime]'1900-03-01') {
                    return Get-InlineStringCell $Reference `
                        $Value.ToString('o', $culture) $Row $Column
                }
                $serialDate = $Value.ToOADate().ToString('G17', $culture)
                return '<c r="{0}" s="2"><v>{1}</v></c>' -f
                    $Reference, $serialDate
            }
            if ($Value.GetType().IsEnum) {
                return Get-InlineStringCell $Reference $Value $Row $Column
            }

            $typeCode = [Type]::GetTypeCode($Value.GetType())
            $number = $null
            $exactNumber = $null

            switch ($typeCode) {
                SByte   { $number = $Value.ToString($culture) }
                Byte    { $number = $Value.ToString($culture) }
                Int16   { $number = $Value.ToString($culture) }
                UInt16  { $number = $Value.ToString($culture) }
                Int32   { $number = $Value.ToString($culture) }
                UInt32  { $number = $Value.ToString($culture) }
                Int64   { $number = $Value.ToString($culture) }
                UInt64  { $number = $Value.ToString($culture) }
                Decimal { $number = $Value.ToString($culture) }
                Single {
                    if ([single]::IsNaN($Value) -or [single]::IsInfinity($Value)) {
                        return Get-InlineStringCell $Reference $Value $Row $Column
                    }
                    $number = $Value.ToString('G9', $culture)
                }
                Double {
                    if ([double]::IsNaN($Value) -or [double]::IsInfinity($Value)) {
                        return Get-InlineStringCell $Reference $Value $Row $Column
                    }
                    $number = $Value.ToString('G15', $culture)
                    try { $roundTrip = [double]::Parse($number, $culture) }
                    catch { $roundTrip = $null }
                    if ($null -eq $roundTrip -or $roundTrip -ne $Value) {
                        $exactNumber = $Value.ToString('G17', $culture)
                    }
                }
            }

            if ($number) {
                if (-not $exactNumber) { $exactNumber = $number }
                $magnitude = [math]::Abs([double]$Value)
                if (
                    (Get-SignificantDigitCount $exactNumber) -gt 15 -or
                    $magnitude -gt 9.99999999999999E+307 -or
                    ($magnitude -ne 0 -and $magnitude -lt 2.2250738585072E-308)
                ) {
                    return Get-InlineStringCell `
                        $Reference $exactNumber $Row $Column
                }
                return '<c r="{0}"><v>{1}</v></c>' -f $Reference, $number
            }

            Get-InlineStringCell $Reference $Value $Row $Column
        }

        $rowCount = 0
        $rowWriter = $null
        $rowStream = $null
        $rowPath = [IO.Path]::Combine(
            [IO.Path]::GetTempPath(),
            'xlsx-' + [guid]::NewGuid().ToString('N') + '.tmp'
        )
        $rowStream = [IO.FileStream]::new(
            $rowPath,
            [IO.FileMode]::CreateNew,
            [IO.FileAccess]::ReadWrite,
            [IO.FileShare]::None,
            65536,
            [IO.FileOptions]::DeleteOnClose
        )
        $rowWriter = [IO.StreamWriter]::new($rowStream, $utf8, 65536, $true)
    }

    process {
        try {
            if ($rowCount -ge $maxDataRows) {
                throw "Excel supports at most $maxDataRows data rows when a header is present."
            }

            $excelRow = $rowCount + 2
            $rowXml = [Text.StringBuilder]::new()
            [void]$rowXml.Append('<row r="{0}">' -f $excelRow)

            if ($explicitProperties) {
                for ($index = 0; $index -lt $columns.Count; $index++) {
                    $name = $columns[$index]
                    $itemProperty = $InputObject.PSObject.Properties[$name]
                    $value = $null
                    if ($null -ne $itemProperty) {
                        [void]$foundProperties.Add($name)
                        $value = $itemProperty.Value
                    }
                    $columnName = Get-ColumnName $index
                    [void]$rowXml.Append((
                        Get-CellXml ($columnName + $excelRow) `
                            $value $excelRow $columnName
                    ))
                }
            }
            else {
                $rowValues = [Collections.Generic.SortedDictionary[int, object]]::new()
                foreach ($itemProperty in $InputObject.PSObject.Properties) {
                    $name = $itemProperty.Name
                    if (-not $columnMap.ContainsKey($name)) {
                        if ($columns.Count -ge $maxColumns) {
                            throw "Excel supports at most $maxColumns columns."
                        }
                        $columnMap.Add($name, $columns.Count)
                        $columns.Add($name)
                    }
                    $rowValues[$columnMap[$name]] = $itemProperty.Value
                }

                foreach ($pair in $rowValues.GetEnumerator()) {
                    $columnName = Get-ColumnName $pair.Key
                    [void]$rowXml.Append((
                        Get-CellXml ($columnName + $excelRow) `
                            $pair.Value $excelRow $columnName
                    ))
                }
            }

            [void]$rowXml.Append('</row>')
            $rowWriter.Write($rowXml.ToString())
            $rowCount++
        }
        catch {
            Close-RowBuffer
            throw
        }
    }

    end {
        $outputStream = $null
        $zip = $null
        $temporaryOutput = $null
        $committed = $false

        try {
            if ($explicitProperties -and $rowCount) {
                $missing = @(
                    $columns | Where-Object {
                        -not $foundProperties.Contains($_)
                    }
                )
                if ($missing.Count) {
                    throw ('Properties not found in the input: {0}' -f
                        ($missing -join ', '))
                }
            }

            $rowWriter.Flush()
            $rowWriter.Dispose()
            $rowWriter = $null
            $rowStream.Position = 0

            [void][IO.Directory]::CreateDirectory($directory)
            $temporaryOutput = [IO.Path]::Combine(
                $directory,
                '.xlsx-' + [guid]::NewGuid().ToString('N') + '.tmp'
            )
            $outputStream = [IO.FileStream]::new(
                $temporaryOutput,
                [IO.FileMode]::CreateNew,
                [IO.FileAccess]::Write,
                [IO.FileShare]::None
            )
            $zip = [IO.Compression.ZipArchive]::new(
                $outputStream,
                [IO.Compression.ZipArchiveMode]::Create,
                $true
            )

            function Add-Part {
                param([string]$Name, [string]$Content)

                $entry = $zip.CreateEntry($Name)
                $writer = [IO.StreamWriter]::new($entry.Open(), $utf8)
                try { $writer.Write($Content) }
                finally { $writer.Dispose() }
            }

            Add-Part '[Content_Types].xml' @'
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>
'@
            Add-Part '_rels/.rels' @'
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>
'@
            Add-Part 'xl/workbook.xml' @'
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Export" sheetId="1" r:id="rId1"/></sheets></workbook>
'@
            Add-Part 'xl/_rels/workbook.xml.rels' @'
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>
'@
            Add-Part 'xl/styles.xml' @'
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>
'@

            $entryStream = $zip.CreateEntry(
                'xl/worksheets/sheet1.xml'
            ).Open()
            $worksheetWriter = [IO.StreamWriter]::new(
                $entryStream,
                $utf8,
                65536,
                $true
            )
            try {
                $worksheetWriter.Write(
                    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
                )
                if ($columns.Count) {
                    $worksheetWriter.Write('<row r="1">')
                    for ($index = 0; $index -lt $columns.Count; $index++) {
                        $columnName = Get-ColumnName $index
                        $worksheetWriter.Write((
                            Get-InlineStringCell ($columnName + '1') `
                                $columns[$index] 1 $columnName 1
                        ))
                    }
                    $worksheetWriter.Write('</row>')
                }
                $worksheetWriter.Flush()
                $rowStream.CopyTo($entryStream)
                $worksheetWriter.Write('</sheetData></worksheet>')
            }
            finally {
                $worksheetWriter.Dispose()
                $entryStream.Dispose()
            }

            $zip.Dispose()
            $zip = $null
            $outputStream.Dispose()
            $outputStream = $null

            if ([IO.File]::Exists($resolvedPath)) {
                $backupOutput = [IO.Path]::Combine(
                    $directory,
                    '.xlsx-' + [guid]::NewGuid().ToString('N') + '.bak'
                )
                [IO.File]::Replace(
                    $temporaryOutput,
                    $resolvedPath,
                    $backupOutput
                )
                $committed = $true
                try { [IO.File]::Delete($backupOutput) }
                catch {
                    Write-Warning "Export succeeded, but the backup remains: $backupOutput"
                }
            }
            else {
                [IO.File]::Move($temporaryOutput, $resolvedPath)
                $committed = $true
            }
        }
        finally {
            foreach ($resource in @($zip, $outputStream)) {
                if ($resource) {
                    try { $resource.Dispose() }
                    catch { Write-Warning "Could not close an output resource: $_" }
                }
            }
            if (
                -not $committed -and
                $temporaryOutput -and
                [IO.File]::Exists($temporaryOutput)
            ) {
                try { [IO.File]::Delete($temporaryOutput) }
                catch { Write-Warning "Could not remove temporary file: $temporaryOutput" }
            }
            Close-RowBuffer
        }
    }
}
