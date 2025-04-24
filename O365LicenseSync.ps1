# Requires MS Graph Powershell module (via PowershellGet)

# Remember: this looks at group memberships in Azure Entra, NOT in AD directly. 
# So don't assume you  run this immediately after a group change in AD in order to assign a license. 
# You need to wait for the DirectorySync (or similar) tool to catch up first, and the timing for that is non-deterministic.
 
# TODO: Add notes about certificate rotation to this location.

# base data
$UsageLocation = 'US' #United States
$tenant = '<tenant name here>'
$log = "path to log file here"

# This depends on app defined in Entra with appropriate permissions 
# Authentication to the app is via AppID and certificate.
$appClient = "Entra Application (client) ID here (will be a guid string)" # This is analogous to the username, rather than password, so safe to leave in the script
# The certficiate should be installed to local computer store, as well as set in the Entra application.
$cert = "O365License" # TODO: select by thumbprint instead. (Getting thumbprints is awkward, but worth it at cert renewal time). Also make sure the script account has access to manage the private key

$data = @{
    'Student' = @{ 
        LicenseSKUs = @(  "OFFICESUBSCRIPTION_STUDENT", "STANDARDWOFFPACK_STUDENT")
        Groups = @('Students', 'OnlineStudents')
        GroupIDs = @()
        Members = @()
        IDs = @()       
    }                         
             
    'Employee' = @{ 
        LicenseSKUs = @("STANDARDWOFFPACK_FACULTY","OFFICESUBSCRIPTION_FACULTY", "POWER_BI_STANDARD_FACULTY" )
		# LicenseSKUs = @("$($tenant):OFFICESUBSCRIPTION_FACULTY")
        Groups = @('FacultyAdmin','FACULTY')  # Do not include GAs or EMPLOYEES (GAs can be students, and Employees is not always cleaned out fast enough)
        GroupIDs = @()
        Members = @()
        IDs = @()
     }
}

ac -Path $log "$(get-date) - Starting"

write-output "Connecting to Entra"
ac -Path $log "$(get-date) - Connecting to Entra"

try {
    Connect-MgGraph -ClientID $appClient -TenantId "$tenant.onmicrosoft.com"  -CertificateName "CN=$cert"
}
catch {
    ac -Path $log "$(get-date) - Error connecting to Entra. Message below. Exiting early"
    ac -Path $log $($_.Exception.Message)
    exit
}

Write-Output "Prepping group membership data"
ac -Path $log "$(get-date) - Prepping group membership data"

# TODO: We're a school, for goodness sake. How likely is it we end up with more than one group at different points the AD OU heirarcy named "Students"? VERY.
#  Therefore we may need to be more specific than "DisplayName" in the future.
$data.Student.GroupIDs = (Get-MgGroup -All | ?{ $data.Student.Groups -contains $_.DisplayName } ).Id
$data.Employee.GroupIDs = (Get-MgGroup -All | ?{ $data.Employee.Groups -contains $_.DisplayName } ).Id

# Student License users
foreach( $id in $data.Student.GroupIDs)
{
    $data.Student.IDs =  $data.Student.IDs + (Get-MgGroupMember -GroupId $id -All).Id
}
$data.Student.IDs = $data.Student.IDs | select -u

# Employee license users -- TODO: convert to method
foreach( $id in $data.Employee.GroupIDs)
{
    $data.Employee.IDs =  $data.Employee.IDs + (Get-MgGroupMember -GroupId $id -All).Id
}
$data.Employee.IDs = $data.Employee.IDs | select -u

# Remove GAs from the employees list. At least for this purpose, they should be treated as students, and can and will get those licenses instead.
# TODO: make this part of the data
$GAGroup = (Get-MgGroup -All | ?{ "GAs" -eq $_.Displayname }).Id
$GAs = (Get-MgGroupMember -GroupId $GAGroup -All).Id
$data.Employee.IDs = $data.Employee.IDs | ?{ -not ( $GAs -contains $_) }

# #############
# Some employees are also students. In these cases, only assign the employee license SKUs (exclude from student if also in employee)
$data.Student.IDs = $data.Student.IDs | ?{ -not ( $data.Employee.IDs -contains $_) }
# #############

$SKUs  = @{}
$errors = ""
$errorCount = 0

# TODO: this organizes the operations by license, which may require several operations per user.
#      If we load them into a dictionary instead (or something sortable, so we sort for removals first),
#      then the new APIs may let us get this to one operation per user that handles any additions/removals all at once.

Write-Output "Computing Student changes"
ac -Path $log "$(get-date) - Computing Student changes"
foreach ($sku in $data.Student.LicenseSKUs)
{
    # TODO: Try/catch blocks if an API call fails
    $AccountSKU = Get-MgSubscribedsku | Where-Object {$_.SkuPartNumber -eq $sku}

    $LicensedIDs  = (Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq $($AccountSKU.SkuId) )" -All).Id

    $Extra = $LicensedIDs |  ?{ -not ( $data.Student.IDs -contains $_)} | ?{ -not ( $SpecialsIDs -contains $_ )}
    $NeedLicense = $data.Student.IDs | ?{-not  ($LicensedIDs -contains $_)}

    $SKUs.Add($sku, @{
            'AccountSKU'=$AccountSKU
            'Extra' = $Extra
            'Needed' = $NeedLicense
        })

	# These are warnings: don't add to error count (there'll be enough of those later)
	if ($AccountSKU.PrepaidUnits.Enabled -le 0) {
		Write-Warning "No $sku licenses found!"
		$errors += "`nNo $sku license found!"
	}
    elseif ($AccountSKU.PrepaidUnits.Enabled -lt $data.Student.IDs.Length) { 
		Write-Warning "Not enough $sku licenses for all students (need $($data.Student.IDs.Length), have $($AccountSKU.PrepaidUnits.Enabled)). Remove user licenses or buy more licenses"
		$errors += "Not enough $sku licenses for all students (need $($data.Student.IDs.Length), have $($AccountSKU.PrepaidUnits.Enabled)). Remove user licenses or buy more licenses`n"
		#TODO: Create an alert of some kind for this
	}
}

Write-Output "Computing Employee changes"
ac -Path $log "$(get-date) - Computing Employee changes"
foreach ($sku in $data.Employee.LicenseSKUs)
{
    $AccountSKU = Get-MgSubscribedsku | Where-Object {$_.SkuPartNumber -eq $sku}

    $LicensedIDs  = (Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq $($AccountSKU.SkuId) )" -All ).Id

    $Extra = $LicensedIDs |  ?{ -not ( $data.Employee.IDs -contains $_)} | ?{ -not ( $SpecialsIDs -contains $_ )}
    $NeedLicense = $data.Employee.IDs | ?{-not  ($LicensedIDs -contains $_)}

    $SKUs.Add($sku, @{
            'AccountSKU'=$AccountSKU
            'Extra' = $Extra
            'Needed' = $NeedLicense
        })

	# These are warnings: don't add to error count (there'll be enough of those later)
	if ($AccountSKU.PrepaidUnits.Enabled -le 0) {
		Write-Warning "No $sku licenses found!"
		$errors += "`nNo $sku license found!"
	}
    elseif ($AccountSKU.PrepaidUnits.Enabled -lt $data.Employee.IDs.Length) { 
		Write-Warning "Not enough $sku licenses for all employees (need $($data.Employee.IDs.Length), have $($AccountSKU.PrepaidUnits.Enabled)). Remove user licenses or buy more licenses"
		$errors += "Not enough $sku licenses for all employees (need $($data.Employee.IDs.Length), have $($AccountSKU.PrepaidUnits.Enabled)). Remove user licenses or buy more licenses`n"
		#TODO: Create an alert of some kind for this
    }
	
}

# Need to avoid both license conflicts and exceeding license counts, so remove everywhere before adding
Write-Output "Syncing license assignments"
ac -Path $log "$(get-date) - Syncing license assignments"
foreach ($s in $SKUs.Keys)
{
    $sku = $SKUs[$s]
    Write-Output "Removing $($sku.Extra.Length) users from $s license." 
    ac -Path $log "$(get-date) - Removing $($sku.Extra.Length) users from $s license." 

    foreach ($User in $sku.Extra) {		
        Try {                                
            $u = Set-MgUserLicense -UserId $User -AddLicenses @() -RemoveLicenses @($sku.AccountSKU.SkuId) -ErrorAction Stop -WarningAction Stop        
            Write-Output "Successfully removed $s license for user $($u.UserPrincipalName)" 
            # Don't log each user change -- too noisy -- but do write it to the console if we're running interactively, and do track the error if we get one
        } catch {
			$errorCount += 1
            try {
                $userName = (Get-MgUser -UserID $User).UserPrincipalName  
            } catch {$userName = $User} # At least we'll have the Id

			$msg = "Error when removing $s license from user $userName : `n $($_.Exception.Message)`r`n"
            Write-Warning $msg
			$errors += "$msg"
        } 
    }
}

# Now add missing users
foreach ($s in $SKUs.Keys)
{
    $sku = $SKUs[$s]  
    Write-Output "Adding $($sku.Needed.Length) users to $s license."
    ac -Path $log "$(get-date) - Adding $($sku.Needed.Length) users to $s license." 

    $toAdd = @(  @{SkuId = $sku.AccountSKU.SkuId} ) # Different from Remove: need array of objects instead of just IDs, 

    foreach ($User in $sku.Needed) {
		# Users are missing this by default, and it blocks adding licenses... doesn't hurt to keep setting it
		Update-MGUser -UserID $User -UsageLocation "US"
        Try {            
            $u = Set-MgUserLicense -UserId $User -AddLicenses $toAdd -RemoveLicenses @() -ErrorAction Stop -WarningAction Stop  
            Write-Output "Successfully added $s license for user $($u.UserPrincipalName)" 
            # Don't log each user change -- too noisy -- but do write it to the console if we're running interactively, and do track the error if we get one
        } catch {
			$errorCount += 1
            try {
                $userName = (Get-MgUser -UserID $User).UserPrincipalName 
            } catch { $userName = $User  } # At least we'll have the Id
			$msg = "Error when adding $s license for user $userName : $($_.Exception.Message)`n"
            Write-Warning $msg
			$errors += "$msg"
        }   
    }
 }
 
 if ($errorCount -gt 0)
 {
	ac -Path $log  "$(get-date) - Found $errorCount errors:`n$errors"
 }

 ac -Path $log  "$(get-date) - Finished`n"
