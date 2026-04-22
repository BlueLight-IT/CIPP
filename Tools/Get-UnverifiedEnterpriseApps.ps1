[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$ExportPath,
    [switch]$IncludeMicrosoftApps
)

$MicrosoftOwnedPublisherIds = @(
    'f8cdef31-a31e-4b4a-93e4-5f571e91255a',
    '72f988bf-86f1-41af-91ab-2d7cd011db47'
)

$requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications')
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        throw "Required module '$module' is not installed. Run: Install-Module $module -Scope CurrentUser"
    }
    Import-Module $module -ErrorAction Stop
}

$connectParams = @{ Scopes = 'Application.Read.All', 'Directory.Read.All'; NoWelcome = $true }
if ($TenantId) { $connectParams.TenantId = $TenantId }

Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Cyan
Connect-MgGraph @connectParams | Out-Null

$context = Get-MgContext
Write-Host ("Connected to tenant: {0} ({1})" -f $context.TenantId, $context.Account) -ForegroundColor Green

Write-Host 'Retrieving enterprise applications (service principals)...' -ForegroundColor Cyan
$servicePrincipals = Get-MgServicePrincipal -All -Property Id, AppId, DisplayName, PublisherName, AppOwnerOrganizationId, VerifiedPublisher, SignInAudience, Tags, CreatedDateTime, ServicePrincipalType

$enterpriseApps = $servicePrincipals | Where-Object {
    $_.ServicePrincipalType -eq 'Application' -and $_.Tags -contains 'WindowsAzureActiveDirectoryIntegratedApp'
}

if (-not $IncludeMicrosoftApps) {
    $enterpriseApps = $enterpriseApps | Where-Object {
        $_.AppOwnerOrganizationId -and ($MicrosoftOwnedPublisherIds -notcontains $_.AppOwnerOrganizationId.ToString())
    }
}

$unverified = $enterpriseApps | Where-Object {
    -not $_.VerifiedPublisher -or [string]::IsNullOrWhiteSpace($_.VerifiedPublisher.VerifiedPublisherId)
}

$results = $unverified | ForEach-Object {
    [PSCustomObject]@{
        DisplayName            = $_.DisplayName
        AppId                  = $_.AppId
        ObjectId               = $_.Id
        PublisherName          = $_.PublisherName
        AppOwnerOrganizationId = $_.AppOwnerOrganizationId
        SignInAudience         = $_.SignInAudience
        CreatedDateTime        = $_.CreatedDateTime
    }
} | Sort-Object DisplayName

Write-Host ("Found {0} unverified enterprise app(s)." -f @($results).Count) -ForegroundColor Yellow
$results | Format-Table -AutoSize

if ($ExportPath) {
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host ("Exported results to {0}" -f $ExportPath) -ForegroundColor Green
}

Disconnect-MgGraph | Out-Null
