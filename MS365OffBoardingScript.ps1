# Microsoft Custom 365 User off-boarding script. Version 1.0 By: Komail Chaudhry
# The script is updated regularly to keep up to date with Microsoft APIs (Current revision April 2024)
# Ensure you're running this script in PowerShell 5.1 or newer.
# Ensure you have installed all the preq Module's needed

# Import AzureAD and Exchange Online Management V3 module
Import-Module AzureAD
Import-Module ExchangeOnlineManagement

# Connect to Azure AD and Exchange Online
Connect-AzureAD
Connect-ExchangeOnline

# Get user(s) from the administrator
$UserPrincipalNames = Read-Host "Enter the User Principal Name(s) of the user(s) to manage, separated by commas"
$UsersArray = $UserPrincipalNames -split ','

# Function to remove a user from all groups
function Remove-UserFromAllGroups {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserObjectId
    )

    $Groups = Get-AzureADUserMembership -ObjectId $UserObjectId

    foreach ($Group in $Groups) {
        if ($Group.DisplayName -eq "All Users") {
            Write-Host "[Default] All Users Group Bypassed for $UserObjectId"
        } else {
            try {
                Remove-AzureADGroupMember -ObjectId $Group.ObjectId -MemberId $UserObjectId
                Write-Host "Removed from group $($Group.DisplayName)"
            } catch {
                Write-Host "Failed to remove from $($Group.DisplayName). Error: $($_.Exception.Message)"
            }
        }
    }
}

# Process each user
foreach ($UserPrincipalName in $UsersArray) {
    $UserPrincipalName = $UserPrincipalName.Trim()
    Write-Host "Processing $UserPrincipalName" -ForegroundColor Cyan

    try {
        # Block the user from signing in
        Set-AzureADUser -ObjectId $UserPrincipalName -AccountEnabled $false
        Write-Host "Sign-in blocked for $UserPrincipalName"

        # Convert user mailbox to a shared mailbox
        $Mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue
        if ($Mailbox) {
            Set-Mailbox -Identity $UserPrincipalName -Type Shared
            Write-Host "Converted mailbox to shared for $UserPrincipalName"
        } else {
            Write-Host "No mailbox found for $UserPrincipalName"
        }

        # Call function to remove user from all groups
        $UserObject = Get-AzureADUser -ObjectId $UserPrincipalName
        Remove-UserFromAllGroups -UserObjectId $UserObject.ObjectId

        # Unassign all licenses
        $Licenses = Get-AzureADUser -ObjectId $UserPrincipalName | Select -ExpandProperty AssignedLicenses
        $LicenseIds = $Licenses | Select -ExpandProperty SkuId
        if ($LicenseIds.Count -gt 0) {
            $LicensesToRemove = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
            $LicensesToRemove.AddLicenses = @()
            $LicensesToRemove.RemoveLicenses = $LicenseIds
            Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $LicensesToRemove
            Write-Host "Licenses unassigned for $UserPrincipalName"
        }

    } catch {
        Write-Host "Error processing $UserPrincipalName. Error: $_" -ForegroundColor Red
    }
}

# Disconnect from Exchange Online and Azure AD
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-AzureAD

Write-Host "Script execution completed. If you received 'Failed to remove from All Users. Error[This is normal and can be ignored]' " -ForegroundColor Green
Read-Host -Prompt "Press Enter to exit"
