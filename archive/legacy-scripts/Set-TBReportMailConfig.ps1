param(
    [Parameter(Mandatory = $true)]
    [string]$SenderEmail,

    [Parameter(Mandatory = $true)]
    [string]$RecipientEmail,

    [Parameter(Mandatory = $true)]
    [string]$AppPassword,

    [string]$SmtpHost = 'smtp.gmail.com',
    [int]$Port = 587,
    [bool]$UseSsl = $true,
    [string]$ConfigPath = 'C:\Users\Administrator\.codex\secrets\tb-report-mail.json',
    [string]$CredentialPath = 'C:\Users\Administrator\.codex\secrets\tb-report-mail-credential.xml'
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Ensure-ParentDirectory {
    param([string]$Path)

    $parent = Split-Path -Parent $Path
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent | Out-Null
    }
}

Ensure-ParentDirectory -Path $ConfigPath
Ensure-ParentDirectory -Path $CredentialPath

$securePassword = ConvertTo-SecureString $AppPassword -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($SenderEmail, $securePassword)
$credential | Export-Clixml -LiteralPath $CredentialPath

$config = [ordered]@{
    sender_email = $SenderEmail
    recipient_email = $RecipientEmail
    smtp_host = $SmtpHost
    port = $Port
    use_ssl = $UseSsl
    credential_path = $CredentialPath
}

$config | ConvertTo-Json | Set-Content -LiteralPath $ConfigPath -Encoding UTF8

Write-Output "Saved mail config to $ConfigPath"
Write-Output "Saved encrypted credential to $CredentialPath"
