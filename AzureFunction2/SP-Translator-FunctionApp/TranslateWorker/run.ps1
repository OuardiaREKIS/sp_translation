# Worker de file : traduit UNE copie de langue par message.
# Message JSON : { siteUrl, page, sourceTitle, topicHeader, commentsDisabled,
#                  language: { code, azCode, id } }
# En cas d'exception, le message est rejoue par la file (5 tentatives par
# defaut, puis file poison) ; le statut Pending/Failed permet la reprise
# par un nouveau clic dans tous les cas.

param($QueueItem, $TriggerMetadata)

. "$PSScriptRoot/../shared/translate-common.ps1"

$job = if ($QueueItem -is [string]) { $QueueItem | ConvertFrom-Json } else { $QueueItem }
Write-Host "Job recu : $($job.page) -> $($job.language.code) (tentative $($TriggerMetadata.DequeueCount))"

# Connexion app-only par certificat (memes reglages que l'orchestrateur)
$connectWatch = [System.Diagnostics.Stopwatch]::StartNew()
if ($env:SP_CERT_PATH) {
    $certPassword = ConvertTo-SecureString $env:SP_CERT_PASSWORD -AsPlainText -Force
    Connect-PnPOnline -Url $job.siteUrl -ClientId $env:SP_CLIENT_ID -Tenant $env:SP_TENANT -CertificatePath $env:SP_CERT_PATH -CertificatePassword $certPassword -ErrorAction Stop
}
else {
    Connect-PnPOnline -Url $job.siteUrl -ClientId $env:SP_CLIENT_ID -Tenant $env:SP_TENANT -Thumbprint $env:SP_THUMBPRINT -ErrorAction Stop
}
Write-Host ("[MOSAIC PERF] {0} {1} {2:N1} s" -f $job.language.code, "Connect-PnP".PadRight(18), $connectWatch.Elapsed.TotalSeconds)

try {
    $ok = Invoke-CopyTranslation -SiteUrl $job.siteUrl -PagePath $job.page `
        -LanguageSp $job.language.code -LanguageAz $job.language.azCode `
        -SourceTitle $job.sourceTitle -TopicHeader ([string]$job.topicHeader) `
        -CommentsDisabled ([bool]$job.commentsDisabled)
    if (-not $ok) {
        # Echec DETERMINISTE (validation canvas) : statut Failed deja pose,
        # inutile de faire rejouer le message par la file.
        Write-Warning "Job termine en Failed (validation) : $($job.page) -> $($job.language.code)"
    }
}
finally {
    Disconnect-PnPOnline
}
