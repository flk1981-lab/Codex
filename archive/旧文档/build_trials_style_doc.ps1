param(
    [Parameter(Mandatory = $true)]
    [string]$MarkdownPath,
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,
    [string]$FontName = 'Times New Roman',
    [string]$HeadingFontName = '',
    [string]$TitleFontName = '',
    [string]$CoverNote = 'Study protocol formatted for journal-style submission preparation',
    [string]$TocTitle = 'Table of Contents',
    [string]$RunningTitle = ''
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if ([string]::IsNullOrWhiteSpace($HeadingFontName)) {
    $HeadingFontName = $FontName
}
if ([string]::IsNullOrWhiteSpace($TitleFontName)) {
    $TitleFontName = $HeadingFontName
}

function Clean-InlineMarkdown {
    param([string]$Text)

    if ($null -eq $Text) { return '' }

    $clean = $Text.TrimEnd()
    $clean = $clean -replace '\[(.+?)\]\((.+?)\)', '$1'
    $clean = $clean -replace '\*\*(.+?)\*\*', '$1'
    $clean = $clean -replace '\*(.+?)\*', '$1'
    $clean = $clean -replace '`(.+?)`', '$1'
    $clean = $clean -replace '\s{2,}$', ''
    return $clean.Trim()
}

function Add-Paragraph {
    param(
        $Document,
        [string]$Text,
        [int]$StyleId = -1,
        [string]$FontName = 'Times New Roman',
        [double]$FontSize = 12,
        [bool]$Bold = $false,
        [bool]$Italic = $false,
        [int]$Alignment = 3,
        [double]$SpaceBefore = 0,
        [double]$SpaceAfter = 6,
        [double]$LineSpacing = 18,
        [bool]$KeepWithNext = $false
    )

    $range = $Document.Range($Document.Content.End - 1, $Document.Content.End - 1)
    $range.Text = ($Text + "`r")
    $range.Style = $StyleId

    $paragraph = $range.Paragraphs.Item(1)
    $paragraph.Range.Font.Name = $FontName
    $paragraph.Range.Font.Size = $FontSize
    $paragraph.Range.Font.Bold = [int]$Bold
    $paragraph.Range.Font.Italic = [int]$Italic
    $paragraph.Format.Alignment = $Alignment
    $paragraph.Format.SpaceBefore = $SpaceBefore
    $paragraph.Format.SpaceAfter = $SpaceAfter
    $paragraph.Format.LineSpacingRule = 0
    $paragraph.Format.LineSpacing = $LineSpacing
    $paragraph.Format.KeepWithNext = [int]$KeepWithNext
    return $paragraph
}

function Add-BulletParagraph {
    param(
        $Document,
        [string]$Text
    )

    $paragraph = Add-Paragraph -Document $Document -Text $Text -StyleId -1 -FontName 'Times New Roman' -FontSize 12 -Alignment 3 -SpaceAfter 3 -LineSpacing 18
    $paragraph.Range.ListFormat.ApplyBulletDefault()
    return $paragraph
}

function Insert-BreakAtEnd {
    param(
        $Document,
        [int]$BreakType
    )

    $range = $Document.Range($Document.Content.End - 1, $Document.Content.End - 1)
    $range.InsertBreak($BreakType)
}

if (-not (Test-Path -LiteralPath $MarkdownPath)) {
    throw "Markdown file not found: $MarkdownPath"
}

$lines = Get-Content -LiteralPath $MarkdownPath -Encoding UTF8
$title = ''
$runningTitle = if ([string]::IsNullOrWhiteSpace($RunningTitle)) { 'TDSC study protocol' } else { $RunningTitle }
$titlePairs = New-Object System.Collections.Generic.List[object]
$bodyItems = New-Object System.Collections.Generic.List[object]

$mode = 'scan'
$pendingLabel = $null
$pendingValue = New-Object System.Collections.Generic.List[string]
$h1 = 0
$h2 = 0
$h3 = 0

foreach ($rawLine in $lines) {
    $line = $rawLine.TrimEnd()

    if (-not $title -and $line -match '^# (.+)$') {
        $title = Clean-InlineMarkdown $Matches[1]
        continue
    }

    if ($line -match '^## Title page$') {
        $mode = 'title'
        continue
    }

    if ($line -match '^## Table of contents$') {
        if ($pendingLabel) {
            $titlePairs.Add([PSCustomObject]@{
                Label = $pendingLabel
                Value = (Clean-InlineMarkdown ($pendingValue -join ' '))
            })
            if (($pendingLabel -eq 'Running title' -or $pendingLabel -eq '页眉短题') -and $pendingValue.Count -gt 0) {
                $runningTitle = Clean-InlineMarkdown ($pendingValue -join ' ')
            }
            $pendingLabel = $null
            $pendingValue = New-Object System.Collections.Generic.List[string]
        }
        $mode = 'toc'
        continue
    }

    if ($mode -eq 'title') {
        if ($line -match '^\*\*(.+?)\*\*') {
            if ($pendingLabel) {
                $titlePairs.Add([PSCustomObject]@{
                    Label = $pendingLabel
                    Value = (Clean-InlineMarkdown ($pendingValue -join ' '))
                })
                if (($pendingLabel -eq 'Running title' -or $pendingLabel -eq '页眉短题') -and $pendingValue.Count -gt 0) {
                    $runningTitle = Clean-InlineMarkdown ($pendingValue -join ' ')
                }
            }
            $pendingLabel = Clean-InlineMarkdown $Matches[1]
            $pendingValue = New-Object System.Collections.Generic.List[string]
            continue
        }

        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        if ($pendingLabel) {
            $pendingValue.Add((Clean-InlineMarkdown $line))
        }
        continue
    }

    if ($mode -eq 'toc') {
        if ($line -match '^## (.+)$') {
            $mode = 'body'
        } else {
            continue
        }
    }

    if ($mode -eq 'body') {
        if ([string]::IsNullOrWhiteSpace($line)) {
            $bodyItems.Add([PSCustomObject]@{ Type = 'blank'; Text = '' })
            continue
        }

        if ($line -match '^## (.+)$') {
            $h1 += 1
            $h2 = 0
            $h3 = 0
            $bodyItems.Add([PSCustomObject]@{
                Type = 'heading1'
                Text = ('{0} {1}' -f $h1, (Clean-InlineMarkdown $Matches[1]))
            })
            continue
        }

        if ($line -match '^### (.+)$') {
            $h2 += 1
            $h3 = 0
            $bodyItems.Add([PSCustomObject]@{
                Type = 'heading2'
                Text = ('{0}.{1} {2}' -f $h1, $h2, (Clean-InlineMarkdown $Matches[1]))
            })
            continue
        }

        if ($line -match '^#### (.+)$') {
            $h3 += 1
            $bodyItems.Add([PSCustomObject]@{
                Type = 'heading3'
                Text = ('{0}.{1}.{2} {3}' -f $h1, $h2, $h3, (Clean-InlineMarkdown $Matches[1]))
            })
            continue
        }

        if ($line -match '^- (.+)$') {
            $bodyItems.Add([PSCustomObject]@{
                Type = 'paragraph'
                Text = ('- ' + (Clean-InlineMarkdown $Matches[1]))
            })
            continue
        }

        if ($line -match '^[0-9]+\.\s+(.+)$') {
            $prefix = $line.Substring(0, $line.IndexOf('.') + 1)
            $suffix = $line.Substring($line.IndexOf('.') + 1).Trim()
            $bodyItems.Add([PSCustomObject]@{
                Type = 'paragraph'
                Text = (Clean-InlineMarkdown ("$prefix $suffix"))
            })
            continue
        }

        $bodyItems.Add([PSCustomObject]@{
            Type = 'paragraph'
            Text = (Clean-InlineMarkdown $line)
        })
    }
}

if (-not $title) {
    throw 'No document title was found in the markdown file.'
}

$outputDir = Split-Path -Parent $OutputPath
if ($outputDir -and -not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

$wdStyleNormal = -1
$wdStyleHeading1 = -2
$wdStyleHeading2 = -3
$wdStyleHeading3 = -4
$wdSectionBreakNextPage = 2
$wdPageBreak = 7
$wdAlignLeft = 0
$wdAlignCenter = 1
$wdAlignJustify = 3
$wdHeaderFooterPrimary = 1
$wdSaveFormatDocx = 16
$wdFieldPage = 33
$wdFieldNumPages = 26
$wdFieldEmpty = -1

$word = $null
$document = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $document = $word.Documents.Add()
    $document.PageSetup.PaperSize = 7
    $document.PageSetup.TopMargin = $word.CentimetersToPoints(2.54)
    $document.PageSetup.BottomMargin = $word.CentimetersToPoints(2.54)
    $document.PageSetup.LeftMargin = $word.CentimetersToPoints(2.54)
    $document.PageSetup.RightMargin = $word.CentimetersToPoints(2.54)

    $document.Styles.Item($wdStyleNormal).Font.Name = $FontName
    $document.Styles.Item($wdStyleNormal).Font.Size = 12
    $document.Styles.Item($wdStyleNormal).ParagraphFormat.Alignment = $wdAlignJustify
    $document.Styles.Item($wdStyleNormal).ParagraphFormat.LineSpacingRule = 0
    $document.Styles.Item($wdStyleNormal).ParagraphFormat.LineSpacing = 18
    $document.Styles.Item($wdStyleNormal).ParagraphFormat.SpaceAfter = 6

    $document.Styles.Item($wdStyleHeading1).Font.Name = $HeadingFontName
    $document.Styles.Item($wdStyleHeading1).Font.Size = 14
    $document.Styles.Item($wdStyleHeading1).Font.Bold = 1
    $document.Styles.Item($wdStyleHeading1).ParagraphFormat.Alignment = $wdAlignLeft
    $document.Styles.Item($wdStyleHeading1).ParagraphFormat.SpaceBefore = 12
    $document.Styles.Item($wdStyleHeading1).ParagraphFormat.SpaceAfter = 6
    $document.Styles.Item($wdStyleHeading1).ParagraphFormat.KeepWithNext = 1

    $document.Styles.Item($wdStyleHeading2).Font.Name = $HeadingFontName
    $document.Styles.Item($wdStyleHeading2).Font.Size = 12
    $document.Styles.Item($wdStyleHeading2).Font.Bold = 1
    $document.Styles.Item($wdStyleHeading2).ParagraphFormat.Alignment = $wdAlignLeft
    $document.Styles.Item($wdStyleHeading2).ParagraphFormat.SpaceBefore = 6
    $document.Styles.Item($wdStyleHeading2).ParagraphFormat.SpaceAfter = 3
    $document.Styles.Item($wdStyleHeading2).ParagraphFormat.KeepWithNext = 1

    $document.Styles.Item($wdStyleHeading3).Font.Name = $HeadingFontName
    $document.Styles.Item($wdStyleHeading3).Font.Size = 12
    $document.Styles.Item($wdStyleHeading3).Font.Bold = 1
    $document.Styles.Item($wdStyleHeading3).Font.Italic = 1
    $document.Styles.Item($wdStyleHeading3).ParagraphFormat.Alignment = $wdAlignLeft
    $document.Styles.Item($wdStyleHeading3).ParagraphFormat.SpaceBefore = 3
    $document.Styles.Item($wdStyleHeading3).ParagraphFormat.SpaceAfter = 3
    $document.Styles.Item($wdStyleHeading3).ParagraphFormat.KeepWithNext = 1

    Add-Paragraph -Document $document -Text $title -StyleId $wdStyleNormal -FontName $TitleFontName -FontSize 16 -Bold $true -Alignment $wdAlignCenter -SpaceAfter 12 -LineSpacing 20 | Out-Null
    Add-Paragraph -Document $document -Text $CoverNote -StyleId $wdStyleNormal -FontName $FontName -FontSize 12 -Italic $true -Alignment $wdAlignCenter -SpaceAfter 18 -LineSpacing 18 | Out-Null

    if ($titlePairs.Count -gt 0) {
        $tableRange = $document.Range($document.Content.End - 1, $document.Content.End - 1)
        $table = $document.Tables.Add($tableRange, $titlePairs.Count, 2)
        $table.Borders.Enable = 0
        $table.AllowAutoFit = $false
        $table.Columns.Item(1).Width = $word.CentimetersToPoints(5.2)
        $table.Columns.Item(2).Width = $word.CentimetersToPoints(10.8)

        for ($i = 1; $i -le $titlePairs.Count; $i++) {
            $pair = $titlePairs[$i - 1]
            $leftCell = $table.Cell($i, 1).Range
            $leftCell.Text = $pair.Label
            $leftCell.Font.Name = $HeadingFontName
            $leftCell.Font.Size = 11
            $leftCell.Font.Bold = 1
            $leftCell.ParagraphFormat.Alignment = $wdAlignLeft
            $leftCell.ParagraphFormat.SpaceAfter = 0

            $rightCell = $table.Cell($i, 2).Range
            $rightCell.Text = $pair.Value
            $rightCell.Font.Name = $FontName
            $rightCell.Font.Size = 11
            $rightCell.Font.Bold = 0
            $rightCell.ParagraphFormat.Alignment = $wdAlignLeft
            $rightCell.ParagraphFormat.SpaceAfter = 0
        }

        $afterTableRange = $document.Range($table.Range.End, $table.Range.End)
        $afterTableRange.InsertParagraphAfter()
        $afterTableRange.InsertParagraphAfter()
    }

    Insert-BreakAtEnd -Document $document -BreakType $wdSectionBreakNextPage

    Add-Paragraph -Document $document -Text $TocTitle -StyleId $wdStyleNormal -FontName $HeadingFontName -FontSize 14 -Bold $true -Alignment $wdAlignLeft -SpaceAfter 6 -LineSpacing 18 | Out-Null
    $tocRange = $document.Range($document.Content.End - 1, $document.Content.End - 1)
    $null = $document.TablesOfContents.Add($tocRange, $true, 1, 3, $false, '', $true, $true, '', $true, $false, $true)
    $postTocRange = $document.Range($document.Content.End - 1, $document.Content.End - 1)
    $postTocRange.InsertParagraphAfter()
    Insert-BreakAtEnd -Document $document -BreakType $wdPageBreak

    foreach ($item in $bodyItems) {
        switch ($item.Type) {
            'blank' {
                Add-Paragraph -Document $document -Text '' -StyleId $wdStyleNormal -FontName $FontName -FontSize 12 -Alignment $wdAlignLeft -SpaceAfter 0 -LineSpacing 18 | Out-Null
            }
            'heading1' {
                Add-Paragraph -Document $document -Text $item.Text -StyleId $wdStyleHeading1 -FontName $HeadingFontName -FontSize 14 -Bold $true -Alignment $wdAlignLeft -SpaceAfter 6 -LineSpacing 18 -KeepWithNext $true | Out-Null
            }
            'heading2' {
                Add-Paragraph -Document $document -Text $item.Text -StyleId $wdStyleHeading2 -FontName $HeadingFontName -FontSize 12 -Bold $true -Alignment $wdAlignLeft -SpaceAfter 3 -LineSpacing 18 -KeepWithNext $true | Out-Null
            }
            'heading3' {
                Add-Paragraph -Document $document -Text $item.Text -StyleId $wdStyleHeading3 -FontName $HeadingFontName -FontSize 12 -Bold $true -Italic $true -Alignment $wdAlignLeft -SpaceAfter 3 -LineSpacing 18 -KeepWithNext $true | Out-Null
            }
            default {
                Add-Paragraph -Document $document -Text $item.Text -StyleId $wdStyleNormal -FontName $FontName -FontSize 12 -Alignment $wdAlignJustify -SpaceAfter 6 -LineSpacing 18 | Out-Null
            }
        }
    }

    if ($document.Sections.Count -ge 2) {
        $section2 = $document.Sections.Item(2)
        $section2.Headers.Item($wdHeaderFooterPrimary).LinkToPrevious = $false
        $section2.Footers.Item($wdHeaderFooterPrimary).LinkToPrevious = $false

        $headerRange = $section2.Headers.Item($wdHeaderFooterPrimary).Range
        $headerRange.Text = $runningTitle
        $headerRange.Font.Name = $HeadingFontName
        $headerRange.Font.Size = 9
        $headerRange.Font.Italic = 1
        $headerRange.ParagraphFormat.Alignment = $wdAlignCenter

        $section2.Footers.Item($wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = $true
        $section2.Footers.Item($wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1

        $footerRange = $section2.Footers.Item($wdHeaderFooterPrimary).Range
        $footerRange.Text = 'Page '
        $footerRange.Font.Name = $FontName
        $footerRange.Font.Size = 9
        $footerRange.ParagraphFormat.Alignment = $wdAlignCenter
        $footerRange.Collapse(0)
        $document.Fields.Add($footerRange, $wdFieldPage, '', $false) | Out-Null
        $footerRange = $section2.Footers.Item($wdHeaderFooterPrimary).Range
        $footerRange.Collapse(0)
        $footerRange.InsertAfter(' of ')
        $footerRange.Collapse(0)
        $document.Fields.Add($footerRange, $wdFieldNumPages, '', $false) | Out-Null
    }

    $document.Repaginate()
    foreach ($toc in $document.TablesOfContents) {
        $toc.Update()
    }
    $document.Fields.Update() | Out-Null
    $document.Repaginate()

    $document.SaveAs([ref]$OutputPath, [ref]$wdSaveFormatDocx)
    try {
        $document.Close()
    }
    catch {
    }
    try {
        $word.Quit()
    }
    catch {
    }
    $document = $null
    $word = $null
}
finally {
    if ($document -ne $null) {
        try {
            $document.Close()
        }
        catch {
        }
    }
    if ($word -ne $null) {
        try {
            $word.Quit()
        }
        catch {
        }
    }
}
