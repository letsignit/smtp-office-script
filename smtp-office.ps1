###############################################################################
# AUTHENTICATION
###############################################################################

$credential = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange `
                             -ConnectionUri "https://outlook.office365.com/powershell-liveid/" `
                             -Credential $credential `
                             -Authentication Basic `
                             -AllowRedirection -ErrorVariable CmdError
if ($CmdError)
{
    throw [System.Exception]::new('CreateSessionError', $CmdError)
}

if ($Session)
{

    Import-PSSession -Name Get-AcceptedDomain, New-AcceptedDomain, Remove-TransportRule, Remove-OutboundConnector,    `
                               Remove-InboundConnector, Get-InboundConnector, Get-OutboundConnector,       `
                               Get-TransportRule, New-OutboundConnector, New-TransportRule, New-InboundConnector,       `
                               Get-RemoteDomain, Set-RemoteDomain, Get-HostedConnectionFilterPolicy, Set-HostedConnectionFilterPolicy,       `
                               Get-DistributionGroup, Set-DistributionGroup, Get-DynamicDistributionGroup, Get-Group `
                               $Session -ErrorVariable CmdError
    Write-Output ""

    if ($CmdError)
    {
        throw [System.Exception]::new('ImportSessionError', $CmdError)
    }
}
else
{
    throw [System.Exception]::new('AuthenticationFailed', "Cannot create a session")
}

###############################################################################
# GLOBAL VARIABLES
###############################################################################


$lsiSMTP = "lsicloud-smtp.letsignit.com"
$ip1 = "40.66.63.89"
$ip2 = "40.66.63.90"
$ip3 = "40.66.63.91"


###############################################################################
# SET DOMAIN
###############################################################################


Write-Output "#### Domain"

$defaultDomain = Get-AcceptedDomain  | where Default -eq $true
$domain = if (($domain = Read-Host "Your domain [$defaultDomain]") -eq '')
{
    $defaultDomain.Id
}
else
{
    $domain
}

$acceptedDomain = Get-AcceptedDomain -Identity $domain

if (!$acceptedDomain)
{
    throw [System.Exception]::new('DomainNotFound', "$domain is not an accepted domain")
}
else
{
    Write-Output "Using $domain"
    Write-Output ""
}


###############################################################################
# SET EMAIL
###############################################################################


Write-Output "#### Email"

$email = Read-Host "Your email "

Write-Output "Using $email"
Write-Output ""


###############################################################################
# CREATE INBOUND CONNECTOR
###############################################################################



Write-Output "#### InboundConnector"

$inbound = Get-InboundConnector | where { $ip1 -in $_.SenderIPAddresses  -and $ip2 -in $_.SenderIPAddresses -and $ip3  -in $_.SenderIPAddresses}
if ($inbound)
{
    Write-Output "Inbound connector found: do nothing"
    $inbound | ft Identity
}
else
{
    New-InboundConnector `
        -Name "LSI to o365" `
        -SenderDomains "*" `
        -ConnectorType OnPremises `
        -RequireTls $true `
        -SenderIPAddresses $ip1, $ip2, $ip3 `
        -CloudServicesMailEnabled $true `
        -ErrorVariable CmdError

    if ($CmdError)
    {
        throw [System.Exception]::new('CreateInboundConnector', $CmdError)
    }

    Write-Output "Inbound by ip connector added"
}


###############################################################################
# CREATE OUTBOUND CONNECTOR
###############################################################################



Write-Output "#### OutboundConnector"

$outbound = Get-OutboundConnector |  Where-Object SmartHosts -match $lsiSMTP
if ($outbound)
{
    Write-Output "Outbound connector found: do nothing"
    $outbound | ft Identity
}
else
{

    New-OutboundConnector `
        -Name "o365 to LSI" `
        -ConnectorType OnPremises `
        -IsTransportRuleScoped $true `
        -UseMxRecord $false `
        -SmartHosts $lsiSMTP `
        -TlsSettings DomainValidation `
        -TlsDomain $lsiSMTP `
        -RouteAllMessagesViaOnPremises $false `
        -ErrorVariable CmdError

    if ($CmdError)
    {
        throw [System.Exception]::new('CreateOutboundConnector', $CmdError)
    }

    write-host "Outbound connector added"

}

###############################################################################
# CREATE TRANSPORT RULE
###############################################################################



Write-Output "#### TransportRule"

$outbound = Get-OutboundConnector |  Where-Object SmartHosts -match $lsiSMTP

$rule = Get-TransportRule | where { $_.RouteMessageOutboundConnector -eq $outbound -and ($_.SenderDomainIs -eq $domain -or $_.FromAddressMatchesPatterns -eq "@$domain$")}
if ($rule)
{
    Write-Output "Outbound transport rule found: do nothing"
    $rule | ft Identity
}
else
{
    New-TransportRule `
        -Name "Route email to LSI $domain" `
        -Enabled $false `
        -FromScope InOrganization `
        -SenderAddressLocation Envelope `
        -FromAddressMatchesPatterns "@$domain$" `
        -From $email `
        -RouteMessageOutboundConnector $outbound.Id `
        -ExceptIfHeaderContainsMessageHeader "X-LSI-Version" `
        -ExceptIfHeaderContainsWords "1.0" `
        -ExceptIfFromAddressMatchesPatterns "<>" `
        -ExceptIfMessageTypeMatches "Calendaring" `
        -StopRuleProcessing $true `
        -ErrorVariable CmdError

    if ($CmdError)
    {
        throw [System.Exception]::new('CreateOutboundTransportRule', $CmdError)
    }

    Write-Output "Outbound transport rules added"
}


###############################################################################
# ADD  HOSTED CONNECTION FILTER POLICY
###############################################################################


Write-Output "#### HostedConnectionFilterPolicy"

$connectionFilter = Get-HostedConnectionFilterPolicy | where { $_.IPAllowList -match $ip1 -and $_.IPAllowList -match $ip2 -and $_.IPAllowList -match $ip3 }
if ($connectionFilter)
{
    Write-Output "Connexion filter found: do nothing"
    $connectionFilter | ft Identity
}
else
{
    Set-HostedConnectionFilterPolicy -Identity Default `
        -IPAllowList @{ Add = $ip1, $ip2, $ip3 } `
        -ErrorVariable CmdError

    if ($CmdError)
    {
        throw [System.Exception]::new('SetHostedConnectionFilterPolicy', $CmdError)
    }

    Write-Output "Connexion filter added"
}


###############################################################################
# DISABLING RICH TEXT FORMAT
###############################################################################


Write-Output "#### Disabling Rich text format"


# use '-ne' because there is 3 value
# $true : always
# $false : never
# '' : use use settings
Get-RemoteDomain |  Where { $_.TNEFEnabled -ne $false } | foreach {
    Set-RemoteDomain -Identity $_.Id -TNEFEnabled $false
}

###############################################################################
# CONFIGURE DISTRIBUTION GROUP
###############################################################################


Write-Output "#### Set ReportToOriginator into distribution list"


Get-DistributionGroup | where { $_.isDirSynced -eq $false -and $_.ReportToOriginatorEnabled -eq $false } | foreach {
    Set-DistributionGroup -Identity $_.Id -ReportToOriginatorEnabled $true
}

Get-DynamicDistributionGroup | where { $_.ReportToOriginatorEnabled -eq $false } | foreach {
    Set-DistributionGroup -Identity $_.Id -ReportToOriginatorEnabled $true
}

Get-DistributionGroup | Sort-Object DisplayName | ft DisplayName, isDirSynced, ReportToOriginatorEnabled

Get-DynamicDistributionGroup | Sort-Object DisplayName | ft DisplayName, ReportToOriginatorEnabled


