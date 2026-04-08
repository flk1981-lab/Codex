param(
    [Parameter(Mandatory = $true)]
    [string]$MarkdownPath,
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,
    [string]$Title = ''
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Escape-XmlText {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return [System.Security.SecurityElement]::Escape($Text)
}

function New-ParagraphXml {
    param(
        [string]$Text,
        [string]$Style = 'Normal'
    )

    $safe = Escape-XmlText $Text
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return '<w:p/>'
    }

    $styleXml = ''
    if ($Style -and $Style -ne 'Normal') {
        $styleXml = "<w:pPr><w:pStyle w:val=`"$Style`"/></w:pPr>"
    }

    return "<w:p>$styleXml<w:r><w:t xml:space=`"preserve`">$safe</w:t></w:r></w:p>"
}

if (-not (Test-Path -LiteralPath $MarkdownPath)) {
    throw "Markdown file not found: $MarkdownPath"
}

$lines = Get-Content -LiteralPath $MarkdownPath -Encoding UTF8
$paragraphs = New-Object System.Collections.Generic.List[string]

if ($Title) {
    $paragraphs.Add((New-ParagraphXml -Text $Title -Style 'Title'))
}

$inCodeBlock = $false
foreach ($rawLine in $lines) {
    $line = $rawLine.TrimEnd()

    if ($line -match '^```') {
        $inCodeBlock = -not $inCodeBlock
        continue
    }

    if ($inCodeBlock) {
        $paragraphs.Add((New-ParagraphXml -Text $line -Style 'Code'))
        continue
    }

    if ([string]::IsNullOrWhiteSpace($line)) {
        $paragraphs.Add((New-ParagraphXml -Text ''))
        continue
    }

    if ($line -match '^# (.+)$') {
        $paragraphs.Add((New-ParagraphXml -Text $Matches[1] -Style 'Heading1'))
        continue
    }

    if ($line -match '^## (.+)$') {
        $paragraphs.Add((New-ParagraphXml -Text $Matches[1] -Style 'Heading2'))
        continue
    }

    if ($line -match '^### (.+)$') {
        $paragraphs.Add((New-ParagraphXml -Text $Matches[1] -Style 'Heading3'))
        continue
    }

    if ($line -match '^- (.+)$') {
        $paragraphs.Add((New-ParagraphXml -Text ("- " + $Matches[1]) -Style 'Normal'))
        continue
    }

    if ($line -match '^[0-9]+\.\s+(.+)$') {
        $prefix = $line.Substring(0, $line.IndexOf('.') + 1)
        $text = $line.Substring($line.IndexOf('.') + 1).Trim()
        $paragraphs.Add((New-ParagraphXml -Text ("$prefix $text") -Style 'Normal'))
        continue
    }

    $paragraphs.Add((New-ParagraphXml -Text $line -Style 'Normal'))
}

$documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
 xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:w10="urn:schemas-microsoft-com:office:word"
 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
 xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
 xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
 xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
 mc:Ignorable="w14 wp14">
  <w:body>
    $($paragraphs -join "`n    ")
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
      <w:cols w:space="708"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>
"@

$stylesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:eastAsia="宋体" w:hAnsi="Calibri" w:cs="Calibri"/>
      <w:sz w:val="22"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:after="240"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:eastAsia="黑体" w:hAnsi="Calibri" w:cs="Calibri"/>
      <w:b/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:eastAsia="黑体" w:hAnsi="Calibri" w:cs="Calibri"/>
      <w:b/>
      <w:sz w:val="28"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr><w:spacing w:before="180" w:after="80"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:eastAsia="黑体" w:hAnsi="Calibri" w:cs="Calibri"/>
      <w:b/>
      <w:sz w:val="26"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:eastAsia="黑体" w:hAnsi="Calibri" w:cs="Calibri"/>
      <w:b/>
      <w:sz w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Code">
    <w:name w:val="Code"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr>
      <w:rFonts w:ascii="Consolas" w:eastAsia="等线" w:hAnsi="Consolas" w:cs="Consolas"/>
      <w:sz w:val="20"/>
    </w:rPr>
  </w:style>
</w:styles>
"@

$contentTypesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"@

$relsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"@

$docRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"@

$now = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$safeTitle = if ($Title) { Escape-XmlText $Title } else { Escape-XmlText ([System.IO.Path]::GetFileNameWithoutExtension($OutputPath)) }

$coreXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>$safeTitle</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">$now</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">$now</dcterms:modified>
</cp:coreProperties>
"@

$appXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
</Properties>
"@

$dir = Split-Path -Parent $OutputPath
if ($dir -and -not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir | Out-Null
}

$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("tdsc_docx_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tmpRoot | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmpRoot '_rels') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmpRoot 'docProps') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmpRoot 'word') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmpRoot 'word\_rels') | Out-Null

try {
    Set-Content -LiteralPath (Join-Path $tmpRoot '[Content_Types].xml') -Value $contentTypesXml -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmpRoot '_rels\.rels') -Value $relsXml -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmpRoot 'docProps\core.xml') -Value $coreXml -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmpRoot 'docProps\app.xml') -Value $appXml -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmpRoot 'word\document.xml') -Value $documentXml -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmpRoot 'word\styles.xml') -Value $stylesXml -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmpRoot 'word\_rels\document.xml.rels') -Value $docRelsXml -Encoding UTF8

    if (Test-Path -LiteralPath $OutputPath) {
        Remove-Item -LiteralPath $OutputPath -Force
    }

    $zipPath = [System.IO.Path]::ChangeExtension($OutputPath, '.zip')
    if (Test-Path -LiteralPath $zipPath) {
        Remove-Item -LiteralPath $zipPath -Force
    }

    [System.IO.Compression.ZipFile]::CreateFromDirectory($tmpRoot, $zipPath)
    Move-Item -LiteralPath $zipPath -Destination $OutputPath
}
finally {
    if (Test-Path -LiteralPath $tmpRoot) {
        Remove-Item -LiteralPath $tmpRoot -Recurse -Force
    }
}
