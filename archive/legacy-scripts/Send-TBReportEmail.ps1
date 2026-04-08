param(
    [Parameter(Mandatory = $true)]
    [string]$MarkdownPath,

    [Parameter(Mandatory = $true)]
    [string]$ReportTitle,

    [string]$Subject,
    [string]$PlainTextSummary = '',
    [string]$ConfigPath = 'C:\Users\Administrator\.codex\secrets\tb-report-mail.json',
    [string]$OutputRoot = 'D:\Codex\reports\tb-research'
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if (-not (Test-Path -LiteralPath $MarkdownPath)) {
    throw "Markdown report not found: $MarkdownPath"
}

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    throw "Mail config not found: $ConfigPath"
}

$config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$credentialPath = [string]$config.credential_path

if (-not (Test-Path -LiteralPath $credentialPath)) {
    throw "Encrypted credential not found: $credentialPath"
}

$credential = Import-Clixml -LiteralPath $credentialPath
$now = Get-Date
$dateFolder = $now.ToString('yyyy-MM')
$safeTitle = [regex]::Replace($ReportTitle, '[^0-9A-Za-z\u4e00-\u9fff._-]+', '_').Trim('_')
$safeBaseName = $now.ToString('yyyy-MM-dd_HHmmss') + '_' + $safeTitle
$targetDir = Join-Path $OutputRoot $dateFolder

if (-not (Test-Path -LiteralPath $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir | Out-Null
}

$docxPath = Join-Path $targetDir ($safeBaseName + '.docx')

$mdToWordScript = Join-Path (Split-Path -Parent $PSScriptRoot) 'md_to_word.ps1'
if (-not (Test-Path -LiteralPath $mdToWordScript)) {
    throw "Markdown-to-Word converter not found: $mdToWordScript"
}

& $mdToWordScript -MarkdownPath $MarkdownPath -OutputPath $docxPath -Title $ReportTitle

$mailMessage = New-Object System.Net.Mail.MailMessage
try {
    $mailMessage.From = [System.Net.Mail.MailAddress]::new([string]$config.sender_email)
    $mailMessage.To.Add([string]$config.recipient_email)
    $mailMessage.Subject = $(if ($Subject) { $Subject } else { $ReportTitle })
    $mailMessage.SubjectEncoding = [System.Text.Encoding]::UTF8
    $mailMessage.BodyEncoding = [System.Text.Encoding]::UTF8
    $mailMessage.IsBodyHtml = $false

    $defaultBody = @(
        'Attached is the latest TB learning report in Word format.',
        '',
        $(if ($PlainTextSummary) { $PlainTextSummary } else { 'This message was generated automatically by Codex.' }),
        '',
        "Sent at: $($now.ToString('yyyy-MM-dd HH:mm:ss'))"
    ) -join [Environment]::NewLine
    $mailMessage.Body = $defaultBody

    $attachment = [System.Net.Mail.Attachment]::new($docxPath)
    $attachment.NameEncoding = [System.Text.Encoding]::UTF8
    $mailMessage.Attachments.Add($attachment)

    $smtpClient = [System.Net.Mail.SmtpClient]::new([string]$config.smtp_host, [int]$config.port)
    $smtpClient.EnableSsl = [bool]$config.use_ssl
    $smtpClient.Credentials = $credential
    $smtpClient.Send($mailMessage)

    Write-Output "Sent report email to $($config.recipient_email)"
    Write-Output "Attachment: $docxPath"
}
finally {
    $mailMessage.Dispose()
}
