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

$statusLines = @(git status --porcelain)
if ($LASTEXITCODE -ne 0) {
    Write-Host (T '6K+75Y+WIEdpdCDnirbmgIHlpLHotKXjgII=')
    exit 1
}

if ($statusLines.Count -gt 0) {
    Write-Host (T '5qOA5rWL5Yiw5L2g5pys5Zyw6L+Y5pyJ5pyq5o+Q5Lqk55qE5pS55Yqo77ya')
    git status --short --branch
    Write-Host ""
    Write-Host (T '5Li66YG/5YWN6KaG55uW5oiW5Yay56qB77yM6K+35YWI6L+Q6KGM4oCc5LiK5Lyg5YiwR2l0SHViLmJhdOKAneaPkOS6pOacrOWcsOaUueWKqO+8jOWGjeS4i+i9veacgOaWsOWGheWuueOAgg==')
    exit 1
}

Write-Host (T '5q2j5Zyo5LuOIEdpdEh1YiDkuIvovb3mnIDmlrDlhoXlrrkuLi4=')
git pull --rebase
if ($LASTEXITCODE -ne 0) {
    Write-Host (T '5LiL6L295aSx6LSl77yM6K+35qOA5p+l572R57uc5oiW5aSE55CG5Yay56qB5ZCO6YeN6K+V44CC')
    exit 1
}

Write-Host ""
Write-Host (T '5LiL6L295a6M5oiQ44CC')
git status --short --branch
