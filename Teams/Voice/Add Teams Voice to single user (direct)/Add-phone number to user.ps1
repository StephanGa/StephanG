<#
.SYNOPSIS
    This script adds a voice number to one user
.INPUTS
    the script will ask for the users alias, the office and the phone number
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    This script needs Global Admin rights and the newest Teams Powershell (2.6.1-preview+) Module
    You will need to change 
#>


function SetPhoneNumber {
    param (
        $Mail,
        $phonenumber        
    )
    Write-Host "Setting phone number - if you have removed it from another user - it may take up to 24h until you can reassign!" -ForegroundColor Yellow
    # Add defined phonenumber
    Set-CsPhoneNumberAssignment -Identity $Mail -PhoneNumber $phonenumber -PhoneNumberType DirectRouting
    Set-CsPhoneNumberAssignment -Identity $Mail -EnterpriseVoiceEnabled $true 
 
    
}

function SupplierDependentSetting {
    param (
        $supplier,
        $mail
    )
    switch (${supplier}) {
        Colt { $policyname = "Colt"; break }
        Swisscom { $policyname = "SwisscomET4T"; break }
        Default { $policyname = "Colt"; break }        
    }
    #Remove it first
    Grant-CsOnlineVoiceRoutingPolicy -Identity $mail -PolicyName $Null
    #Grant the one that is matching
    Grant-CsOnlineVoiceRoutingPolicy -Identity $mail -PolicyName $policyname
}

function PhoneNumberInformation {
    param (
        $Mail,
        $phonenumber
    )
    $checkuser = Get-CsOnlineUser $Mail | Select-Object Firstname,Lastname,Office,LineURI
    If ($checkuser.LineURI -ne ""){
        Write-Host "User has already a Line assigned - you can quit the script by hitting strg+c" $($checkuser.LineURI) -ForegroundColor Yellow
        Pause
        }
        Write-Host "Checking for doubles"
        $double= Get-CsOnlineUser | Select-Object Firstname,Lastname,LineURI,UserPrincipalName | Where-Object {$_.LineURI -eq "tel:$phonenumber"}
        If ($double) {
            Write-Host "User $mail will get the phone number from $double - you can quit the script by hitting strg+c" -ForegroundColor Yellow
            pause
            Write-Host "Removing phonenumber from $double"  -ForegroundColor DarkRed -BackgroundColor White
            Remove-CsPhoneNumberAssignment -Identity $double.UserPrincipalName -PhoneNumber $phonenumber -PhoneNumberType DirectRouting
            }
        Else
            {
            Write-host "No doubles"
            }      
        
}

## Start
Write-Host "Welcome to the Teams Voice provisioning tool -- this tool can add and change phone numbers" -ForegroundColor Blue -BackgroundColor White

$mail = Read-Host "Copy and paste the mail of the user"
$check = Get-ADUser -Filter 'userPrincipalName -eq $mail'
If (!$check)
    {
    Write-host "User: $mail not existant" -ForegroundColor Yellow
    break;
    }
Else
    {
$Supplier = Read-Host "Enter the supplier (Colt or Swisscom)"
$phonenumber = Read-Host "Enter the phone number (e.g. +43123456789) - no copy and paste from Excel etc - notepad first"
$phonenumber = $phonenumber.Replace(" ","")
Connect-MicrosoftTeams

PhoneNumberInformation -Mail $Mail -phonenumber $phonenumber;
SetPhoneNumber -Mail $Mail -phonenumber $phonenumber;
SupplierDependentSetting -Supplier $Supplier -mail $mail
Get-CsOnlineUser -identity $Mail | Select-Object Firstname,Lastname,Office,LineURI
    }