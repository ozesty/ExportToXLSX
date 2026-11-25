function Export-ToXLSX {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [AllowNull()]
        [psobject[]]$InputObject,

        [Parameter(Mandatory, Position = 0)]
        [string]$Path,

        [Alias('Properties')]
        [string[]]$Property,

        [Alias('ExcludeProperties')]
        [string[]]$ExcludeProperty,

        [string]$WorksheetName = 'Export',

        [switch]$NoClobber,

        [switch]$TruncateLongValues,

        [switch]$Strict,

        [switch]$BoldHeader,

        [Alias('Sortable')]
        [switch]$AutoFilter,

        [switch]$FreezeHeader
    )

    begin {
        $config = @{
            MaxDataRows           = 1048575
            MaxColumns            = 16384
            MaxCellCharacters     = 32767
            MinColumnWidth        = 10
            MaxColumnWidth        = 60
            ExcelMinDate          = [DateTime]'1900-03-01'
            ExcelMaxDate          = [DateTime]'9999-12-31T23:59:59'
            Culture               = [Globalization.CultureInfo]::InvariantCulture
            Utf8                  = [Text.UTF8Encoding]::new($false)
            TruncateLongValues    = $TruncateLongValues.IsPresent
            BoldHeader            = $BoldHeader.IsPresent
            AutoFilter            = $AutoFilter.IsPresent
            FreezeHeader          = $FreezeHeader.IsPresent
            XmlSpecialRegex       = [regex]::new(
                '[\x00-\x08\x0B-\x1F\uD800-\uDFFF\uFFFE\uFFFF]|_x[0-9A-Fa-f]{4}_',
                [Text.RegularExpressions.RegexOptions]::Compiled
            )
            XmlEscapePatternRegex = [regex]::new(
                '_(?=x[0-9A-Fa-f]{4}_)',
                [Text.RegularExpressions.RegexOptions]::Compiled
            )
            AdMetadataProperties  = [Collections.Generic.HashSet[string]]::new(
                [string[]]@(
                    'PropertyNames',
                    'AddedProperties',
                    'RemovedProperties',
                    'ModifiedProperties',
                    'PropertyCount'
                ),
                [StringComparer]::OrdinalIgnoreCase
            )
            ExcludedExact         = [Collections.Generic.HashSet[string]]::new(
                [StringComparer]::OrdinalIgnoreCase
            )
            ExcludedPatterns      = [Collections.Generic.List[Management.Automation.WildcardPattern]]::new()
        }

        Add-Type -AssemblyName System.IO.Compression

        $provider = $null
        $drive = $null
        $sessionPath = $ExecutionContext.SessionState.Path
        $resolvedPath = $sessionPath.GetUnresolvedProviderPathFromPSPath(
            $Path,
            [ref]$provider,
            [ref]$drive
        )
        [void]$drive

        if ($null -eq $provider -or $provider.Name -ne 'FileSystem') {
            throw "Output path must use the FileSystem provider: $Path"
        }
        if ([IO.Path]::GetExtension($resolvedPath) -ine '.xlsx') {
            throw "Output path must use the .xlsx extension: $resolvedPath"
        }
        if ([IO.Directory]::Exists($resolvedPath)) {
            throw "Output path is an existing directory: $resolvedPath"
        }
        if ($NoClobber -and [IO.File]::Exists($resolvedPath)) {
            throw "Output file already exists and -NoClobber was specified: $resolvedPath"
        }
        $directory = [IO.Path]::GetDirectoryName($resolvedPath)
        if ([string]::IsNullOrEmpty($directory)) {
            throw "Could not determine target directory for path: $resolvedPath"
        }
        $rootPath = [IO.Path]::GetPathRoot($resolvedPath)
        if (-not [string]::IsNullOrEmpty($rootPath) -and -not [IO.Directory]::Exists($rootPath)) {
            throw "Target drive or root path does not exist: $rootPath"
        }

        if (
            [string]::IsNullOrWhiteSpace($WorksheetName) -or
            $WorksheetName.Length -gt 31 -or
            $WorksheetName -match '[:\\/\?\*\[\]]' -or
            $WorksheetName.StartsWith("'") -or
            $WorksheetName.EndsWith("'")
        ) {
            throw "Invalid worksheet name '$WorksheetName'. Names must be 1-31 characters, cannot contain : \ / ? * [ ], and cannot start or end with a single quote."
        }

        function Get-ColumnName {
            param([int]$Number)

            $name = ''
            do {
                $name = [char](65 + $Number % 26) + $name
                $Number = [int][math]::Floor($Number / 26) - 1
            }
            while ($Number -ge 0)
            $name
        }

        $columns = [Collections.Generic.List[string]]::new()
        $columnNames = [Collections.Generic.List[string]]::new()
        $columnWidths = [Collections.Generic.List[int]]::new()
        $columnMap = [Collections.Generic.Dictionary[string, int]]::new(
            [StringComparer]::OrdinalIgnoreCase
        )
        $foundProperties = [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase
        )

        function Add-Column {
            param([string]$Name)

            if ($columns.Count -ge $config.MaxColumns) {
                throw "Excel supports at most $($config.MaxColumns) columns."
            }
            $index = $columns.Count
            $columnMap.Add($Name, $index)
            $columns.Add($Name)
            $columnNames.Add((Get-ColumnName -Number $index))
            $columnWidths.Add([math]::Min($Name.Length, $config.MaxColumnWidth))
            $index
        }

        function Test-ExcludedProperty {
            param([string]$Name)

            if ($config.ExcludedExact.Count -gt 0 -and $config.ExcludedExact.Contains($Name)) {
                return $true
            }
            foreach ($wildcard in $config.ExcludedPatterns) {
                if ($wildcard.IsMatch($Name)) {
                    return $true
                }
            }
            $false
        }

        if ($PSBoundParameters.ContainsKey('ExcludeProperty') -and $ExcludeProperty) {
            foreach ($pattern in $ExcludeProperty) {
                if ([string]::IsNullOrWhiteSpace($pattern)) {
                    throw 'ExcludeProperty names cannot be empty.'
                }
                if ([Management.Automation.WildcardPattern]::ContainsWildcardCharacters($pattern)) {
                    $config.ExcludedPatterns.Add(
                        [Management.Automation.WildcardPattern]::new(
                            $pattern,
                            [Management.Automation.WildcardOptions]::IgnoreCase
                        )
                    )
                }
                else {
                    [void]$config.ExcludedExact.Add($pattern)
                }
            }
        }

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
                if (Test-ExcludedProperty -Name $name) {
                    continue
                }
                [void](Add-Column -Name $name)
            }
            if ($columns.Count -eq 0) {
                throw 'All specified properties were excluded by -ExcludeProperty.'
            }
        }

        $shouldExport = $PSCmdlet.ShouldProcess($resolvedPath, 'Export XLSX')
        if (-not $shouldExport) {
            return
        }

        function Close-RowBuffer {
            foreach ($resource in @($rowWriter, $rowStream)) {
                if ($resource) {
                    try { $resource.Dispose() }
                    catch { Write-Warning "Could not close an export buffer: $_" }
                }
            }
        }

        function Measure-ColumnWidth {
            param(
                [int]$ColumnIndex,
                [int]$DisplayLength
            )

            if ($DisplayLength -gt $columnWidths[$ColumnIndex]) {
                $columnWidths[$ColumnIndex] = [math]::Min($DisplayLength, $config.MaxColumnWidth)
            }
        }

        function Get-DisplayLength {
            param([string]$Text)

            if ([string]::IsNullOrEmpty($Text)) { return 0 }
            if ($Text.IndexOf("`n") -lt 0) { return $Text.Length }

            $longest = 0
            foreach ($line in ($Text -split "`r?`n")) {
                if ($line.Length -gt $longest) { $longest = $line.Length }
            }
            $longest
        }

        function ConvertTo-XmlString {
            param(
                [AllowNull()][string]$Text,
                [int]$Row,
                [string]$Column
            )

            if ([string]::IsNullOrEmpty($Text)) { return '' }
            if ($Text.Length -gt $config.MaxCellCharacters) {
                if ($config.TruncateLongValues) {
                    $truncatedLength = $config.MaxCellCharacters
                    if (
                        $truncatedLength -gt 0 -and
                        [char]::IsHighSurrogate($Text[$truncatedLength - 1])
                    ) {
                        $truncatedLength--
                    }
                    $Text = $Text.Substring(0, $truncatedLength)
                    Write-Warning "Cell $Column$Row exceeded Excel's $($config.MaxCellCharacters)-character limit and was truncated."
                }
                else {
                    throw "Cell $Column$Row exceeds Excel's $($config.MaxCellCharacters)-character limit."
                }
            }

            if (-not $config.XmlSpecialRegex.IsMatch($Text)) {
                return [Security.SecurityElement]::Escape($Text)
            }

            # Protect literal SpreadsheetML escape-looking text before creating
            # escapes for characters XML cannot represent directly.
            $Text = $config.XmlEscapePatternRegex.Replace($Text, '_x005F_')

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
                    [void]$clean.Append(('_x{0:X4}_' -f [int]$character))
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
            if ($Value -is [string]) { return $Value }
            if ($Value -is [byte[]]) {
                return [Convert]::ToBase64String($Value)
            }
            if ($Value -is [DateTimeOffset] -or $Value -is [DateTime]) {
                return $Value.ToString('o', $config.Culture)
            }
            if ($Value -is [Collections.IEnumerable]) {
                $parts = [Collections.Generic.List[string]]::new()
                foreach ($item in $Value) {
                    if ($null -eq $item -or $item -is [DBNull]) {
                        $parts.Add('')
                    }
                    elseif ($item -is [byte[]]) {
                        $parts.Add([Convert]::ToBase64String($item))
                    }
                    elseif ($item -is [DateTimeOffset] -or $item -is [DateTime]) {
                        $parts.Add($item.ToString('o', $config.Culture))
                    }
                    else {
                        $parts.Add([string]$item)
                    }
                }
                if ($parts.Count -eq 0) { return '' }
                return [string]::Join('; ', $parts)
            }
            [string]$Value
        }

        function Get-InlineStringCell {
            param(
                [string]$Reference,
                $Value,
                [int]$Row,
                [string]$Column,
                [int]$ColumnIndex,
                [int]$Style = 0
            )

            $rawText = ConvertTo-TextValue -Value $Value
            if ($rawText.Length -eq 0 -and $Style -eq 0) {
                return ''
            }
            Measure-ColumnWidth -ColumnIndex $ColumnIndex -DisplayLength (Get-DisplayLength -Text $rawText)
            $text = ConvertTo-XmlString -Text $rawText -Row $Row -Column $Column
            $styleAttribute = if ($Style) { ' s="{0}"' -f $Style } else { '' }
            '<c r="{0}"{1} t="inlineStr"><is><t xml:space="preserve">{2}</t></is></c>' -f $Reference, $styleAttribute, $text
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
                [string]$Column,
                [int]$ColumnIndex
            )

            if ($null -eq $Value -or $Value -is [DBNull]) {
                return ''
            }
            if ($Value -is [string]) {
                return Get-InlineStringCell -Reference $Reference -Value $Value -Row $Row -Column $Column -ColumnIndex $ColumnIndex
            }
            if ($Value -is [bool]) {
                Measure-ColumnWidth -ColumnIndex $ColumnIndex -DisplayLength 5
                return '<c r="{0}" t="b"><v>{1}</v></c>' -f $Reference, [int]$Value
            }
            if ($Value -is [DateTime]) {
                if ($Value -lt $config.ExcelMinDate -or $Value -gt $config.ExcelMaxDate) {
                    $isoDate = $Value.ToString('o', $config.Culture)
                    return Get-InlineStringCell -Reference $Reference -Value $isoDate -Row $Row -Column $Column -ColumnIndex $ColumnIndex
                }
                Measure-ColumnWidth -ColumnIndex $ColumnIndex -DisplayLength 19
                $serialDate = $Value.ToOADate().ToString('G17', $config.Culture)
                return '<c r="{0}" s="2"><v>{1}</v></c>' -f $Reference, $serialDate
            }
            if ($Value.GetType().IsEnum) {
                return Get-InlineStringCell -Reference $Reference -Value $Value -Row $Row -Column $Column -ColumnIndex $ColumnIndex
            }

            $typeCode = [Type]::GetTypeCode($Value.GetType())

            switch ($typeCode) {
                { $_ -in 'SByte', 'Byte', 'Int16', 'UInt16', 'Int32', 'UInt32' } {
                    $number = $Value.ToString($config.Culture)
                    Measure-ColumnWidth -ColumnIndex $ColumnIndex -DisplayLength $number.Length
                    return '<c r="{0}"><v>{1}</v></c>' -f $Reference, $number
                }
                { $_ -in 'Int64', 'UInt64', 'Decimal' } {
                    $exactNumber = $Value.ToString($config.Culture)
                    if ((Get-SignificantDigitCount -Number $exactNumber) -gt 15) {
                        return Get-InlineStringCell -Reference $Reference -Value $exactNumber -Row $Row -Column $Column -ColumnIndex $ColumnIndex
                    }
                    Measure-ColumnWidth -ColumnIndex $ColumnIndex -DisplayLength $exactNumber.Length
                    return '<c r="{0}"><v>{1}</v></c>' -f $Reference, $exactNumber
                }
                Single {
                    if ([single]::IsNaN($Value) -or [single]::IsInfinity($Value)) {
                        return Get-InlineStringCell -Reference $Reference -Value $Value -Row $Row -Column $Column -ColumnIndex $ColumnIndex
                    }
                    $number = $Value.ToString('G9', $config.Culture)
                    Measure-ColumnWidth -ColumnIndex $ColumnIndex -DisplayLength $number.Length
                    return '<c r="{0}"><v>{1}</v></c>' -f $Reference, $number
                }
                Double {
                    if ([double]::IsNaN($Value) -or [double]::IsInfinity($Value)) {
                        return Get-InlineStringCell -Reference $Reference -Value $Value -Row $Row -Column $Column -ColumnIndex $ColumnIndex
                    }
                    $magnitude = [math]::Abs($Value)
                    if (
                        $magnitude -gt 9.99999999999999E+307 -or
                        ($magnitude -ne 0 -and $magnitude -lt 2.2250738585072E-308)
                    ) {
                        $exactDouble = $Value.ToString('G17', $config.Culture)
                        return Get-InlineStringCell -Reference $Reference -Value $exactDouble -Row $Row -Column $Column -ColumnIndex $ColumnIndex
                    }
                    $number = $Value.ToString('G15', $config.Culture)
                    try { $roundTrip = [double]::Parse($number, $config.Culture) }
                    catch { $roundTrip = $null }
                    if ($null -eq $roundTrip -or $roundTrip -ne $Value) {
                        $number = $Value.ToString('G17', $config.Culture)
                    }
                    Measure-ColumnWidth -ColumnIndex $ColumnIndex -DisplayLength ([math]::Min($number.Length, 15))
                    return '<c r="{0}"><v>{1}</v></c>' -f $Reference, $number
                }
            }

            Get-InlineStringCell -Reference $Reference -Value $Value -Row $Row -Column $Column -ColumnIndex $ColumnIndex
        }

        $rowCount = 0
        $rowWriter = $null
        $rowStream = $null
        $rowXml = [Text.StringBuilder]::new(4096)
        $rowValues = [Collections.Generic.SortedDictionary[int, object]]::new()
        $rowPath = [IO.Path]::Combine(
            [IO.Path]::GetTempPath(),
            'xlsx-' + [guid]::NewGuid().ToString('N') + '.tmp'
        )
        try {
            $rowStream = [IO.FileStream]::new(
                $rowPath,
                [IO.FileMode]::CreateNew,
                [IO.FileAccess]::ReadWrite,
                [IO.FileShare]::None,
                65536,
                [IO.FileOptions]::DeleteOnClose
            )
            $rowWriter = [IO.StreamWriter]::new($rowStream, $config.Utf8, 65536, $true)
        }
        catch {
            Close-RowBuffer
            throw
        }
    }

    process {
        if (-not $shouldExport -or $null -eq $InputObject) {
            return
        }

        try {
            foreach ($item in $InputObject) {
                if ($null -eq $item) { continue }

                if ($rowCount -ge $config.MaxDataRows) {
                    throw "Excel supports at most $($config.MaxDataRows) data rows when a header is present."
                }

                $excelRow = $rowCount + 2
                [void]$rowXml.Clear()
                [void]$rowXml.Append(('<row r="{0}">' -f $excelRow))

                if ($explicitProperties) {
                    for ($index = 0; $index -lt $columns.Count; $index++) {
                        $name = $columns[$index]
                        $itemProperty = $item.PSObject.Properties[$name]
                        if ($null -eq $itemProperty -or -not $itemProperty.IsGettable) {
                            continue
                        }
                        [void]$foundProperties.Add($name)
                        $value = $itemProperty.Value
                        if ($null -eq $value -or $value -is [DBNull]) {
                            continue
                        }
                        $columnName = $columnNames[$index]
                        $cellXml = Get-CellXml -Reference ($columnName + $excelRow) -Value $value -Row $excelRow -Column $columnName -ColumnIndex $index
                        if ($cellXml.Length -gt 0) {
                            [void]$rowXml.Append($cellXml)
                        }
                    }
                }
                else {
                    $isAdEntity = $false
                    foreach ($typeName in $item.PSObject.TypeNames) {
                        if ($typeName -like '*Microsoft.ActiveDirectory.Management.ADEntity*') {
                            $isAdEntity = $true
                            break
                        }
                    }

                    $rowValues.Clear()
                    foreach ($itemProperty in $item.PSObject.Properties) {
                        if (-not $itemProperty.IsGettable) { continue }
                        $name = $itemProperty.Name
                        if ($isAdEntity -and $config.AdMetadataProperties.Contains($name)) {
                            continue
                        }
                        if (
                            ($config.ExcludedExact.Count -gt 0 -and $config.ExcludedExact.Contains($name)) -or
                            ($config.ExcludedPatterns.Count -gt 0 -and (Test-ExcludedProperty -Name $name))
                        ) {
                            continue
                        }
                        $columnIndex = 0
                        if (-not $columnMap.TryGetValue($name, [ref]$columnIndex)) {
                            $columnIndex = Add-Column -Name $name
                        }
                        $rowValues[$columnIndex] = $itemProperty.Value
                    }

                    foreach ($pair in $rowValues.GetEnumerator()) {
                        $value = $pair.Value
                        if ($null -eq $value -or $value -is [DBNull]) {
                            continue
                        }
                        $columnIndex = $pair.Key
                        $columnName = $columnNames[$columnIndex]
                        $cellXml = Get-CellXml -Reference ($columnName + $excelRow) -Value $value -Row $excelRow -Column $columnName -ColumnIndex $columnIndex
                        if ($cellXml.Length -gt 0) {
                            [void]$rowXml.Append($cellXml)
                        }
                    }
                }

                [void]$rowXml.Append('</row>')
                $rowWriter.Write($rowXml.ToString())
                $rowCount++
            }
        }
        catch {
            Close-RowBuffer
            throw
        }
    }

    end {
        if (-not $shouldExport) {
            return
        }

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
                    $missingMessage = 'Properties not found in the input: {0}' -f ($missing -join ', ')
                    if ($Strict) {
                        throw $missingMessage
                    }
                    Write-Warning $missingMessage
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
                param(
                    [string]$Name,
                    [string]$Content
                )

                $entry = $zip.CreateEntry($Name)
                $writer = [IO.StreamWriter]::new($entry.Open(), $config.Utf8)
                try { $writer.Write($Content) }
                finally { $writer.Dispose() }
            }

            $xmlHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            $escapedSheetName = [Security.SecurityElement]::Escape($WorksheetName)

            $contentTypesXml = $xmlHeader + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
            $rootRelsXml = $xmlHeader + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
            $workbookXml = $xmlHeader + ('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="{0}" sheetId="1" r:id="rId1"/></sheets></workbook>' -f $escapedSheetName)
            $workbookRelsXml = $xmlHeader + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
            $stylesXml = $xmlHeader + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'

            Add-Part -Name '[Content_Types].xml' -Content $contentTypesXml
            Add-Part -Name '_rels/.rels' -Content $rootRelsXml
            Add-Part -Name 'xl/workbook.xml' -Content $workbookXml
            Add-Part -Name 'xl/_rels/workbook.xml.rels' -Content $workbookRelsXml
            Add-Part -Name 'xl/styles.xml' -Content $stylesXml

            $entryStream = $zip.CreateEntry('xl/worksheets/sheet1.xml').Open()
            $worksheetWriter = [IO.StreamWriter]::new(
                $entryStream,
                $config.Utf8,
                65536,
                $true
            )
            try {
                $worksheetWriter.Write(
                    $xmlHeader +
                    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                )
                $rangeRef = $null
                if ($columns.Count) {
                    $lastColumnName = $columnNames[$columns.Count - 1]
                    $lastRow = $rowCount + 1
                    $rangeRef = 'A1:{0}{1}' -f $lastColumnName, $lastRow
                    $worksheetWriter.Write(('<dimension ref="{0}"/>' -f $rangeRef))
                    if ($config.FreezeHeader) {
                        $worksheetWriter.Write(
                            '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
                        )
                    }
                    $worksheetWriter.Write('<cols>')
                    for ($index = 0; $index -lt $columns.Count; $index++) {
                        $width = [math]::Min(
                            [math]::Max($columnWidths[$index] + 3, $config.MinColumnWidth),
                            $config.MaxColumnWidth
                        )
                        $colNumber = $index + 1
                        $worksheetWriter.Write(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f $colNumber, $width))
                    }
                    $worksheetWriter.Write('</cols>')
                }

                $worksheetWriter.Write('<sheetData>')
                if ($columns.Count) {
                    $headerStyle = if ($config.BoldHeader) { 1 } else { 0 }
                    $worksheetWriter.Write('<row r="1">')
                    for ($index = 0; $index -lt $columns.Count; $index++) {
                        $columnName = $columnNames[$index]
                        $headerCell = Get-InlineStringCell -Reference ($columnName + '1') -Value $columns[$index] -Row 1 -Column $columnName -ColumnIndex $index -Style $headerStyle
                        $worksheetWriter.Write($headerCell)
                    }
                    $worksheetWriter.Write('</row>')
                }
                $worksheetWriter.Flush()
                $rowStream.CopyTo($entryStream)
                $worksheetWriter.Write('</sheetData>')
                if ($config.AutoFilter -and $rangeRef) {
                    $worksheetWriter.Write(('<autoFilter ref="{0}"/>' -f $rangeRef))
                }
                $worksheetWriter.Write('</worksheet>')
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
                [IO.File]::Replace(
                    $temporaryOutput,
                    $resolvedPath,
                    [NullString]::Value
                )
                $committed = $true
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
