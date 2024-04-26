# Microsoft Custom 365 User off-boarding script. Version 1.1 By: Komail Chaudhry
# The script is updated regularly to keep up to date with Microsoft APIs (Current revision April 2024)
# Ensure you're running this script in PowerShell 5.1 or newer.
# Ensure you have installed all the prerequisite Modules needed

# Import AzureAD and Exchange Online Management V3 module
Import-Module AzureAD
Import-Module ExchangeOnlineManagement

# Connect to Azure AD and Exchange Online
Connect-AzureAD
Connect-ExchangeOnline

# Read user principal names from a file
$UserPrincipalNames = Get-Content "UserEmails.txt" -Raw
$UsersArray = $UserPrincipalNames -split ','

# Initialize result variables
$ProcessedCount = 0
$ErrorLog = New-Object System.Collections.Generic.List[string]

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
                $ErrorLog.Add("Failed to remove $($UserObjectId) from $($Group.DisplayName). Error: $($_.Exception.Message)")
            }
        }
    }
}

# Process each user
foreach ($UserPrincipalName in $UsersArray) {
    $UserPrincipalName = $UserPrincipalName.Trim()
    if (-not [string]::IsNullOrWhiteSpace($UserPrincipalName)) {
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
            
            $ProcessedCount += 1

        } catch {
            Write-Host "Error processing $UserPrincipalName. Error: $_" -ForegroundColor Red
            $ErrorLog.Add("Error processing $UserPrincipalName. Error: $_")
        }
    } else {
        Write-Host "Skipped an empty or invalid user entry"
    }
}

# Disconnect from Exchange Online and Azure AD
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-AzureAD

# Summary of results
Write-Host "Script execution completed. Total Processed Users: $ProcessedCount" -ForegroundColor Green
Write-Host "Total Errors: $($ErrorLog.Count)" -ForegroundColor Red
if ($ErrorLog.Count -gt 0) {
   
}
Read-Host -Prompt "Press Enter to exit"

