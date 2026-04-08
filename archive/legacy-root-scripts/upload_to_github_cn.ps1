$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function T([string]$base64) {
    return [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($base64))
}

Set-Location -LiteralPath $PSScriptRoot

Write-Host ((T '5b2T5YmN5LuT5bqT55uu5b2V77ya') + $PSScriptRoot)
Write-Host ""

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host (T '5pyq5qOA5rWL5YiwIEdpdO+8jOivt+WFiOWuieijhSBHaXQg5ZCO5YaN6K+V44CC')
    exit 1
}

try {
    git rev-parse --is-inside-work-tree | Out-Null
} catch {
    Write-Host (T '5b2T5YmN5paH5Lu25aS55LiN5pivIEdpdCDku5PlupPjgII=')
    exit 1
}

$commitMessage = Read-Host (T '6K+36L6T5YWl5pys5qyh5o+Q5Lqk6K+05piO77yM55u05o6l5Zue6L2m5YiZ5L2/55So6buY6K6k5YC84oCc5pu05paw4oCd')
if ([string]::IsNullOrWhiteSpace($commitMessage)) {
    $commitMessage = T '5pu05paw'
}

Write-Host ""
Write-Host (T '56ys5LiA5q2l77ya5pW055CG5pys5qyh5pS55YqoLi4u')
git add -A
if ($LASTEXITCODE -ne 0) {
    Write-Host (T '5omn6KGMIGdpdCBhZGQgLUEg5aSx6LSl44CC')
    exit 1
}

$statusLines = @(git status --porcelain)
if ($LASTEXITCODE -ne 0) {
    Write-Host (T '6K+75Y+WIEdpdCDnirbmgIHlpLHotKXjgII=')
    exit 1
}

if ($statusLines.Count -eq 0) {
    Write-Host ""
    Write-Host (T '5b2T5YmN5rKh5pyJ5Y+v5o+Q5Lqk55qE5pS55Yqo44CC')
    git status --short --branch
    exit 0
}

Write-Host ""
Write-Host (T '56ys5LqM5q2l77ya55Sf5oiQ5pys5Zyw5o+Q5LqkLi4u')
git commit -m $commitMessage
if ($LASTEXITCODE -ne 0) {
    Write-Host (T '5omn6KGMIGdpdCBjb21taXQg5aSx6LSl44CC')
    exit 1
}

Write-Host ""
Write-Host (T '56ys5LiJ5q2l77ya5LiK5Lyg5YiwIEdpdEh1Yi4uLg==')
git push
if ($LASTEXITCODE -ne 0) {
    Write-Host (T '5omn6KGMIGdpdCBwdXNoIOWksei0peOAgg==')
    Write-Host (T '5aaC5p6c6L+c56uv5bey5pyJ5paw5YaF5a6577yM5bu66K6u5YWI6L+Q6KGM4oCc5LuOR2l0SHVi5LiL6L295pyA5pawLmJhdOKAneWGjemHjeivleOAgg==')
    exit 1
}

Write-Host ""
Write-Host (T '5LiK5Lyg5a6M5oiQ44CC')
git status --short --branch
