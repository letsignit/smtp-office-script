###############################################################################
# INSTALL MFA MODULE
###############################################################################

$MFAExchangeModule = ((Get-ChildItem -Path $( $env:LOCALAPPDATA + "\Apps\2.0\" ) -Filter CreateExoPSSession.ps1 -Recurse).FullName | Select-Object -Last 1)
If ($MFAExchangeModule -eq $null)
{
    Write-Host  `nPlease install Exchange Online MFA Module.  -ForegroundColor yellow
    Write-Host You can install module using below blog : `nLink `nOR you can install module directly by entering "Y"`n
    $Confirm = Read-Host Are you sure you want to install module directly? [Y] Yes [N] No
    if ($Confirm -match "[yY]")
    {
        Write-Host Yes
        Start-Process "iexplore.exe" "https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application"
    }
    else
    {
        Start-Process 'https://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/'
        Exit
    }
    $Confirmation = Read-Host Have you installed Exchange Online MFA Module? [Y] Yes [N] No
    if ($Confirmation -match "[yY]")
    {
        $MFAExchangeModule = ((Get-ChildItem -Path $( $env:LOCALAPPDATA + "\Apps\2.0\" ) -Filter CreateExoPSSession.ps1 -Recurse).FullName | Select-Object -Last 1)
        If ($MFAExchangeModule -eq $null)
        {
            Write-Host Exchange Online MFA module is not available -ForegroundColor red
            Exit
        }
    }
    else
    {
        Write-Host Exchange Online PowerShell Module is required
        Start-Process 'https://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/'
        Exit
    }
}

# load module
. "$MFAExchangeModule"


###############################################################################
# AUTHENTICATION
###############################################################################


$Session = Connect-EXOPSSession -WarningAction SilentlyContinue  -ErrorVariable CmdError

Write-Output ""

if ($CmdError)
{
    throw [System.Exception]::new('ConnectionFailed', $CmdError)
}

###############################################################################
# HYDRATATION
###############################################################################

$dehydrated = Get-OrganizationConfig  | Select -ExpandProperty IsDehydrated

if ($dehydrated)
{
    Enable-OrganizationCustomization
    Write-Output "Organization hydration done"
}
else
{
    Write-Output "Organization hydration not needed"
}


###############################################################################
# GLOBAL VARIABLES
###############################################################################


$lsiSMTP = "smtp-fr.letsignit.com"
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

$inbound = Get-InboundConnector | where { $ip1 -in $_.SenderIPAddresses -and $ip2 -in $_.SenderIPAddresses -and $ip3 -in $_.SenderIPAddresses }
if ($inbound)
{
    Set-InboundConnector -Identity $inbound.Id `
        -SenderDomains "*" `
        -ConnectorType OnPremises `
        -RequireTls $true `
        -SenderIPAddresses $ip1, $ip2, $ip3 `
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

$outbound = Get-OutboundConnector |  Where-Object { $_.SmartHosts -match $lsiSMTP -or $_.SmartHosts -match "lsicloud-smtp.letsignit.com" }
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

    Write-Output "Outbound connector found: updated"
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

$rule = Get-TransportRule | where { $_.RouteMessageOutboundConnector -eq $outbound -and ($_.SenderDomainIs -eq $domain -or $_.FromAddressMatchesPatterns -eq "@$domain$") }
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
        -Name "Route email to LSI $domain" `
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

$connectionFilter = Get-HostedConnectionFilterPolicy | where { $_.IPAllowList -match $ip1 -and $_.IPAllowList -match $ip2 -and $_.IPAllowList -match $ip3 }
if ($connectionFilter)
{
    Set-HostedConnectionFilterPolicy -Identity Default `
        -IPAllowList @{ Add = $ip1, $ip2, $ip3 } `
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

