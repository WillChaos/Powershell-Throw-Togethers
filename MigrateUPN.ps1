#requires -version 4
#Requires -RunAsAdministrator
<#
.SYNOPSIS
  Changes UPN,Mail,MailNicName + Proxy addresses in bulk (Based on OU selection)
    
.INPUTS
  N/A

.OUTPUTS
  N/A

.NOTES
  Version:        0.7
  Author:         willf
  Creation Date:  20/2/20
  Purpose/Change: Projects

.INFO
  Script doesnt handle any real advanced error checking yet. however, it has been stablly used on multiple migrations.
  (Example: no null or sanity checks on OU/UPN setters - And no try/catch on set-aduser yet)
  Reccomended to target a test OU first. 
#>

# -------------------------------------------------  PreReqs  --------------------------------------------------------- #
Import-Module *ActiveDirectory*

# ------------------------------------------------- Functions --------------------------------------------------------- #
Function Write-Banner()
{
    # swank
    Write-Host "-------------------------------------------------" -ForegroundColor Magenta
    Write-Host "|                UPN mod script                  " -ForegroundColor Magenta
    Write-Host "-------------------------------------------------" -ForegroundColor Magenta
    Write-Host "Primary UPNS Will be set as: $UPN "                -ForegroundColor Green
    Write-Host "-------------------------------------------------" -ForegroundColor Magenta
    Read-Host "Press enter to begin..."

}
Function Select-OU()
{
    return (Get-ADOrganizationalUnit -Filter * | select name, ObjectName , DistinguishedName| Out-GridView -PassThru -Title "Select OU" ).DistinguishedName
}
Function Select-UsersInOU()
{
    return (Get-ADUser -Filter * -SearchBase (Select-OU))
}
Function Select-UPN()
{
    return (Get-adforest | select UPNSuffixes -ExpandProperty UPNSuffixes -Unique | Out-GridView -Title "Select UPN" -PassThru).ToString()
}
# -------------------------------------------------   Exec    --------------------------------------------------------- #

# Main
$UPN = Select-UPN
Write-Banner

foreach($User in Select-UsersInOU)
{
    Write-Host "[*] Currently looking at: "$User.UserPrincipalName -ForegroundColor Gray
    
    # set working vars
    $SamName = $User.SamAccountName
    $NewUPN  =  $SamName + "@" + $UPN
    
    # Set UPN
    Write-Host "-[+] Setting UPN to: $NewUPN" -ForegroundColor DarkGray
    Set-ADUser -Identity $User -UserPrincipalName $NewUPN 

    # Set Proxy Address to lower
    Write-Host "-[+] Setting all Proxy addresses entries to lowercase" -ForegroundColor DarkGray
    $OldProxyList = (Get-ADUser $User -Properties proxyaddresses).proxyaddresses
    $NewProxyList = 
    foreach($ProxyAddress in $OldProxyList)
    {
        $ProxyAddress.tolower()
    }

    set-aduser $User -Replace @{proxyaddresses = $NewProxyList}

    # Set Primary SMTP to Upper
    Write-Host "-[+] Removing proxy Address: smtp:$NewUPN" -ForegroundColor DarkGray
    Set-ADUser -Identity $User -Remove @{proxyAddresses = ("smtp:"+$NewUPN)}

    Write-Host "-[+] Adding proxy Address: SMTP:$NewUPN" -ForegroundColor DarkGray
    Set-ADUser -Identity $User -Add @{proxyAddresses = ("SMTP:"+$NewUPN)}

    # Set email address attribute
    Write-Host "-[+] Setting Email Address: $NewUPN" -ForegroundColor DarkGray
    Set-ADUser -Identity $User -EmailAddress $NewUPN
    
    # Set MailNic attribute
    Write-Host "-[+] Setting MailNicname Attribute: $NewUPN" -ForegroundColor DarkGray
    Set-ADUser -Identity $User -Replace @{mailNickname=$NewUPN}

}
