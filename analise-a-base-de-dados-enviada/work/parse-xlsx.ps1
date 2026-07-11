param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [Parameter(Mandatory = $true)]
    [string]$OutJson
)

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-ColumnIndex {
    param([string]$CellRef)
    $letters = ($CellRef -replace '[0-9]', '').ToUpperInvariant()
    $n = 0
    foreach ($ch in $letters.ToCharArray()) {
        $n = ($n * 26) + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $n - 1
}

function Get-EntryXml {
    param($Zip, [string]$Name)
    $entry = $Zip.GetEntry($Name)
    if ($null -eq $entry) { return $null }
    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream)
        try {
            [xml]$xml = $reader.ReadToEnd()
            return $xml
        } finally {
            $reader.Dispose()
        }
    } finally {
        $stream.Dispose()
    }
}

function Get-SharedStrings {
    param($Zip)
    $xml = Get-EntryXml -Zip $Zip -Name 'xl/sharedStrings.xml'
    $items = @()
    if ($null -eq $xml) { return $items }
    foreach ($si in $xml.GetElementsByTagName('si')) {
        $parts = @()
        foreach ($t in $si.GetElementsByTagName('t')) {
            $parts += $t.InnerText
        }
        $items += ($parts -join '')
    }
    return $items
}

function Get-WorkbookSheets {
    param($Zip)
    $workbook = Get-EntryXml -Zip $Zip -Name 'xl/workbook.xml'
    $rels = Get-EntryXml -Zip $Zip -Name 'xl/_rels/workbook.xml.rels'
    $relMap = @{}
    foreach ($rel in $rels.GetElementsByTagName('Relationship')) {
        $relMap[$rel.Id] = 'xl/' + $rel.Target.TrimStart('/')
    }
    $sheets = @()
    foreach ($sheet in $workbook.GetElementsByTagName('sheet')) {
        $rid = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $sheets += [pscustomobject]@{
            name = $sheet.name
            path = $relMap[$rid]
        }
    }
    return $sheets
}

function Convert-CellValue {
    param($Cell, $SharedStrings)
    $type = $Cell.t
    $valueNode = $Cell.GetElementsByTagName('v') | Select-Object -First 1
    if ($type -eq 'inlineStr') {
        $parts = @()
        foreach ($t in $Cell.GetElementsByTagName('t')) { $parts += $t.InnerText }
        return ($parts -join '')
    }
    if ($null -eq $valueNode) { return $null }
    $raw = $valueNode.InnerText
    if ($type -eq 's') {
        $idx = [int]$raw
        if ($idx -ge 0 -and $idx -lt $SharedStrings.Count) { return $SharedStrings[$idx] }
        return $raw
    }
    if ($type -eq 'b') { return ($raw -eq '1') }
    $num = 0.0
    if ([double]::TryParse($raw, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$num)) {
        return $num
    }
    return $raw
}

$zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $Path))
try {
    $sharedStrings = Get-SharedStrings -Zip $zip
    $sheetDefs = Get-WorkbookSheets -Zip $zip
    $workbookOut = @()
    foreach ($sheetDef in $sheetDefs) {
        $xml = Get-EntryXml -Zip $zip -Name $sheetDef.path
        $rows = @()
        foreach ($row in $xml.GetElementsByTagName('row')) {
            $cells = @{}
            foreach ($cell in $row.GetElementsByTagName('c')) {
                $idx = Get-ColumnIndex -CellRef $cell.r
                $cells[$idx] = Convert-CellValue -Cell $cell -SharedStrings $sharedStrings
            }
            if ($cells.Count -gt 0) {
                $max = ($cells.Keys | Measure-Object -Maximum).Maximum
                $values = for ($i = 0; $i -le $max; $i++) {
                    if ($cells.ContainsKey($i)) { $cells[$i] } else { $null }
                }
                $rows += ,$values
            }
        }
        $workbookOut += [pscustomobject]@{
            name = $sheetDef.name
            rows = $rows
        }
    }
    $workbookOut | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $OutJson -Encoding UTF8
} finally {
    $zip.Dispose()
}
