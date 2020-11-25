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
                               Get-DistributionGroup, Set-DistributionGroup, Get-DynamicDistributionGroup, Get-Group, `
                               Get-OrganizationConfig, Set-TransportRule, Set-InboundConnector, Set-OutboundConnector `
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
# HYDRATATION
###############################################################################

$dehydrated = Get-OrganizationConfig  | Select -ExpandProperty IsDehydrated

if ($dehydrated) {
    Enable-OrganizationCustomization
    Write-Output "Organization hydration done"
} else {
    Write-Output "Organization hydration not needed"
}


###############################################################################
# GLOBAL VARIABLES
###############################################################################


$lsiSMTP = "smtp-ca.letsignit.com"
$ip1 = "40.80.240.101"
$ip2 = "52.155.24.145"
$ip3 = "165.22.231.153"
$ip4 = "165.22.231.150"


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

$inbound = Get-InboundConnector | where { $ip1 -in $_.SenderIPAddresses  -and $ip2 -in $_.SenderIPAddresses }
if ($inbound)
{
    Set-InboundConnector -Identity $inbound.Id `
        -SenderDomains "*" `
        -ConnectorType OnPremises `
        -RequireTls $true `
        -SenderIPAddresses $ip1, $ip2, $ip3, $ip4 `
        -CloudServicesMailEnabled $true `
        -ErrorVariable CmdError

    if ($CmdError)
    {
        throw [System.Exception]::new('UpdateInboundConnector', $CmdError)
    }

    Write-Output "Inbound connector updated"
    $inbound | ft Identity
}
else
{
    New-InboundConnector `
        -Name "LSI_USE to o365" `
        -SenderDomains "*" `
        -ConnectorType OnPremises `
        -RequireTls $true `
        -SenderIPAddresses $ip1, $ip2, $ip3, $ip4 `
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
    Set-OutboundConnector -Identity $outbound.Id `
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
        throw [System.Exception]::new('UpdateOutboundConnector', $CmdError)
    }

    Write-Output "Outbound connector found: do nothing"
    $outbound | ft Identity
}
else
{

    New-OutboundConnector `
        -Name "o365 to LSI_USE" `
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


Write-Output "#### Exception TransportRule"

$exceptRule = Get-TransportRule | where { $_.MessageTypeMatches -eq "OOF" -and $_.SetHeaderName -eq "X-LSI-Version" }
if ($exceptRule)
{
    Set-TransportRule -Identity $exceptRule.Name `
        -MessageTypeMatches "OOF" `
        -SetHeaderName "X-LSI-Version" `
        -SetHeaderValue "1.0" `
        -ErrorVariable CmdError

    if ($CmdError)
    {
        throw [System.Exception]::new('UpdateOutboundTransportRuleExcept', $CmdError)
    }

    Write-Output "Outbound transport rule exception updated"
}
else
{
    New-TransportRule `
        -Name "LSI exception for automatic reply" `
        -Priority 0 `
        -Enabled $true `
        -MessageTypeMatches OOF `
        -SetHeaderName "X-LSI-Version" `
        -SetHeaderValue "1.0" `
        -ErrorVariable CmdError

    if ($CmdError)
    {
        throw [System.Exception]::new('CreateOutboundTransportRuleExcept', $CmdError)
    }

    Write-Output "Outbound transport exception rules added"
}

Write-Output "#### TransportRule"

$outbound = Get-OutboundConnector |  Where-Object SmartHosts -match $lsiSMTP

$rule = Get-TransportRule | where { $_.RouteMessageOutboundConnector -eq $outbound -and ($_.SenderDomainIs -eq $domain -or $_.FromAddressMatchesPatterns -eq "@$domain$")}
if ($rule)
{
    Set-TransportRule -Identity $rule.Name `
        -FromScope InOrganization `
        -SenderAddressLocation Envelope `
        -RouteMessageOutboundConnector $outbound.Id `
        -ExceptIfHeaderContainsMessageHeader "X-LSI-Version" `
        -ExceptIfHeaderContainsWords "1.0" `
        -ExceptIfFromAddressMatchesPatterns "<>" `
        -ExceptIfMessageTypeMatches "Calendaring" `
        -ExceptIfMessageSizeOver "30MB" `
        -StopRuleProcessing $true `
        -ErrorVariable CmdError

    if ($CmdError)
    {
        throw [System.Exception]::new('CreateOutboundTransportRule', $CmdError)
    }

    Write-Output "Outbound transport rules updated"
}
else
{
    New-TransportRule `
        -Name "Route email to LSI_USE $domain" `
        -Priority 1 `
        -FromScope InOrganization `
        -SenderAddressLocation Envelope `
        -FromAddressMatchesPatterns "@$domain$" `
        -From $email `
        -RouteMessageOutboundConnector $outbound.Id `
        -ExceptIfHeaderContainsMessageHeader "X-LSI-Version" `
        -ExceptIfHeaderContainsWords "1.0" `
        -ExceptIfFromAddressMatchesPatterns "<>" `
        -ExceptIfMessageTypeMatches "Calendaring" `
        -ExceptIfMessageSizeOver "30MB" `
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

$connectionFilter = Get-HostedConnectionFilterPolicy | where { $_.IPAllowList -match $ip1 -and $_.IPAllowList -match $ip2 }
if ($connectionFilter)
{
    Set-HostedConnectionFilterPolicy -Identity Default `
        -IPAllowList @{ Add = $ip1, $ip2, $ip3, $ip4} `
        -ErrorVariable CmdError

    if ($CmdError)
    {
        throw [System.Exception]::new('SetHostedConnectionFilterPolicy', $CmdError)
    }

    Write-Output "Connexion filter updated"
    $connectionFilter | ft Identity
}
else
{
    Set-HostedConnectionFilterPolicy -Identity Default `
        -IPAllowList @{ Add = $ip1, $ip2, $ip3, $ip4} `
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


try
{

    Write-Output "#### Set ReportToOriginator into distribution list"


    Get-DistributionGroup | where { $_.isDirSynced -eq $false -and $_.ReportToOriginatorEnabled -eq $false } | foreach {
        Set-DistributionGroup -Identity $_.Id -ReportToOriginatorEnabled $true
    }

    Get-DynamicDistributionGroup | where { $_.ReportToOriginatorEnabled -eq $false } | foreach {
        Set-DistributionGroup -Identity $_.Id -ReportToOriginatorEnabled $true
    }

    Get-DistributionGroup | Sort-Object DisplayName | ft DisplayName, isDirSynced, ReportToOriginatorEnabled

    Get-DynamicDistributionGroup | Sort-Object DisplayName | ft DisplayName, ReportToOriginatorEnabled

}
catch
{
    Write-Output "Ran into an issue: $PSItem"
}
