#Copyright 2015 Joel Coehoorn. 

#Acknowledgements:
# This script was original based on the work by Johan Dahlbom located here:
# http://365lab.net/2014/04/15/office-365-assign-licenses-based-on-groups-using-powershell/
# However, it has been almost completely re-written twice since then, and boasts significant added functionality.

#This script syncs Office 365 license assignments from group membership in the Azure Active Directory Tenant.
# Adjust the variable assignments in the first section to match your environment.
# This script assumes you already have your Active Directory syncing to Azure, and does NOT address setting that up.
# This script should be set to run as a scheduled task in your local environment on a machine that has the Office 365 commandlets installed.
# No support is provided for configuring your local Active Directory to the Azure tenant, and no support is provided for installing the powershell commandlets. Support inquiries in those areas will be ignored.
# A general guide to configuring your local environment is available here: https://support.office.com/en-ca/article/Managing-Office-365-and-Exchange-Online-with-Windows-PowerShell-06a743bb-ceb6-49a9-a61d-db4ffdf54fa6?ui=en-US&rs=en-CA&ad=CA

#General TODOs: 
# Need better logging and alerts
# A lot of duplicated code for the two tiers could be consolidated

#Note - the Script Requires PowerShell 3.0
Import-Module MSOnline

#Office 365 Admin Credentials
$CloudUsername = '<username here>'  #often something like admin@<tenant>.onmicrosoft.com
$CloudPassword = ConvertTo-SecureString '<password here>' -AsPlainText -Force #yes, you have to hard-code a password :(. Be sure to keep the script somewhere protected
$UsageLocation = 'US' #United States. More codes here: https://www.iso.org/obp/ui/#search/code/
$tenant = '<tenant name here>'

#Special IDs that may be licensed for, say, testing, but are not in the synced group, and should never be removed (at least not automatically removed)
$SpecialsIDs = @([GUID]("00000000-0000-0000-0000-000000000000"))

#This is intended for two tiers (student/employee), such that users will never need a mix of the tiers. Multiple AD Groups and SKUs are supported for each tier.
# Users in the 2nd tier will always only get those license SKUs... if they are in both tiers, only the 2nd tier licenses will be applied.
# A long list of possible SKUs is available here: http://blogs.technet.com/b/treycarlee/archive/2013/11/01/list-of-powershell-licensing-sku-s-for-office-365.aspx
# Mostly, though, I found the SKUs I needed by assigning the license to my own account via the GUI and then asking Powershell what licenses I have
# TODO: Generalize Student/Employee names
$data = @{
	'Student' = @{ 
		LicenseSKUs = @("$($tenant):OFFICESUBSCRIPTION_STUDENT", "$($tenant):STANDARDWOFFPACK_STUDENT")
		Groups = @('StudentsGroup')
		GroupIDs = @()
		Members = @()
		IDs = @()       
	}

	'Employee' = @{ 
		LicenseSKUs = @("$($tenant):STANDARDWOFFPACK_FACULTY","$($tenant):OFFICESUBSCRIPTION_FACULTY" )
		Groups = @('EmployeeGroup')  
		GroupIDs = @()
		Members = @()
		IDs = @()
	}
}
######################################
#End of settings

Write-Output "Preparing synchronization data"

#Connect to Office 365 
$CloudCred = New-Object System.Management.Automation.PSCredential $CloudUsername, $CloudPassword
Connect-MsolService -Credential $CloudCred

$data.Student.GroupIDs = (Get-MsolGroup -All | ?{ $data.Student.Groups -contains $_.DisplayName } ).ObjectId
$data.Employee.GroupIDs = (Get-MsolGroup -All | ?{ $data.Employee.Groups -contains $_.DisplayName } ).ObjectId

#Student license users
foreach( $id in $data.Student.GroupIDs)
{
	$data.Student.Members =  $data.Student.Members + (Get-MsolGroupMember -GroupObjectId $id -All)
}
$data.Student.Members = $data.Student.Members | Sort-Object -property EmailAddress -Unique
$data.Student.IDs = ($data.Student.Members).ObjectId

#Employee license users
#TODO: abstract to method, reuse code from student section
foreach( $id in $data.Employee.GroupIDs)
{
	$data.Employee.Members =  $data.Employee.Members + (Get-MsolGroupMember -GroupObjectId $id -All)
}
$data.Employee.Members = $data.Employee.Members | Sort-Object -property EmailAddress -Unique
$data.Employee.IDs = ($data.Employee.Members).ObjectId

##############
#Some employees are also students. In these cases, only assign the employee license SKUs (exclude from student if also in employee)
# If you don't want this behaviour (you want the employee group to be additive, the union of the two permissions) just comment out these two lines below
# Be warned that a number of SKU's conflict with each other. It's not always possible to just add a license, and in many markets it can be expensive.
$data.Student.Members = $data.Student.Members | ?{ -not ( $data.Employee.IDs -contains $_.ObjectId)}
$data.Student.IDs = $data.Student.IDs | ?{ -not ( $data.Employee.IDs -contains $_) }
##############

$SKUs  = @{}
$AllUsers = Get-MsolUser -All | ?{ $_.IsLicensed -eq "TRUE"} #Can be slow. Unfortunately, we need this to know who has what specific licenses already assigned

#Compute changes to Student tier
foreach ($sku in $data.Student.LicenseSKUs)
{
	$AccountSKU = Get-MsolAccountSku | Where-Object {$_.AccountSKUID -eq $sku}
	$LicensedIDs  = ($ALLUsers | ?{  ($_.Licenses | ?{ $_.AccountSkuId -eq $sku}).Length -gt 0}).ObjectId 

	$Extra = $LicensedIDs |  ?{ -not ( $data.Student.IDs -contains $_)} | ?{ -not ( $SpecialsIDs -contains $_ )}
	$NeedLicense = $data.Student.IDs | ?{-not  ($LicensedIDs -contains $_)}

	$SKUs.Add($sku, @{
		'AccountSKU'=$AccountSKU
		'Extra' = $Extra
		'Needed' = $NeedLicense
	})

	if ($AccountSKU.ActiveUnits -lt $data.Student.IDs.Length) { 
		Write-Warning "Not enough $sku licenses for all students, please remove user licenses or buy more licenses"
		#TODO: Create an alert of some kind for this
	}
}

#Compute changes to Employee tier 
#TODO: re-use code from student section
foreach ($sku in $data.Employee.LicenseSKUs)
{
	$AccountSKU = Get-MsolAccountSku | Where-Object {$_.AccountSKUID -eq $sku}
	$LicensedIDs  = ($AllUsers | ?{  ($_.Licenses | ?{ $_.AccountSkuId -eq $sku}).Length -gt 0}).ObjectId

	$Extra = $LicensedIDs |  ?{ -not ( $data.Employee.IDs -contains $_)} | ?{ -not ( $SpecialsIDs -contains $_ )}
	$NeedLicense = $data.Employee.IDs | ?{-not  ($LicensedIDs -contains $_)}

	$SKUs.Add($sku, @{
			'AccountSKU'=$AccountSKU
			'Extra' = $Extra
			'Needed' = $NeedLicense
		})

	if ($AccountSKU.ActiveUnits -lt $data.Employee.IDs.Length) { 
		Write-Warning "Not enough $sku licenses for all employees, please remove user licenses or buy more licenses"
		#TODO: Create an alert of some kind for this
	}
}

#Need to avoid both license conflicts and exceeding license counts, so remove everywhere before adding
foreach ($s in $SKUs.Keys)
{
	$sku = $SKUs[$s]
	Write-Output "Removing $($sku.Extra.Length) users from $s license." 

	foreach ($User in $sku.Extra) {
		Try {
			$userName = (Get-MsolUser -ObjectId $User).UserPrincipalName
			Set-MsolUserLicense -ObjectId $User -RemoveLicenses $sku.AccountSKU.AccountSkuId -ErrorAction Stop -WarningAction Stop        
			Write-Output "Successfully removed $s license for $userName with ID: $User"
		} catch {
			Write-Warning "Error when removing licensing from $userName`r`n$_"
		}
	}
}

#Now add missing users
foreach ($s in $SKUs.Keys)
{
	$sku = $SKUs[$s]  
	Write-Output "Adding $($sku.Needed.Length) users to $s license."

	foreach ($User in $sku.Needed) {
		Try {
			$userName = (Get-MsolUser -ObjectId $User).UserPrincipalName
			Set-MsolUser -ObjectId $User -UsageLocation $UsageLocation -ErrorAction Stop -WarningAction Stop
			Set-MsolUserLicense -ObjectId $User -AddLicenses $sku.AccountSKU.AccountSkuId -ErrorAction Stop -WarningAction Stop
			Write-Output "Successfully added $s license for user $userName with ID: $User"
		} catch {
			Write-Warning "Error when licensing user: $Username`r`n$_"
		}
	}
}
