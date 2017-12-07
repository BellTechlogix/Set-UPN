<#
Set-UPN.ps1
Created By - Kristopher Roy
Ceated On - 11/29/2017
Last Modified - 12/6/2017
Purpose - The purpose of this script is to set a new UPN for users, initially based upon csv list, but eventually will include options to manually input users, or grab and license all!
#>

function Get-FileName
{
  param(
      [Parameter(Mandatory=$false)]
      [string] $Filter,
      [Parameter(Mandatory=$false)]
      [switch]$Obj,
      [Parameter(Mandatory=$False)]
      [string]$Title = "Select A File"
    )
 
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
  $OpenFileDialog.initialDirectory = $initialDirectory
  $OpenFileDialog.FileName = $Title
  #can be set to filter file types
  IF($Filter -ne $null){
  $FilterString = '{0} (*.{1})|*.{1}' -f $Filter.ToUpper(), $Filter
	$OpenFileDialog.filter = $FilterString}
  if(!($Filter)) { $Filter = "All Files (*.*)| *.*"
  $OpenFileDialog.filter = $Filter
  }
  $OpenFileDialog.ShowDialog() | Out-Null
  ## dont bother asking, just give back the object
  IF($OBJ){
  $fileobject = GI -Path $OpenFileDialog.FileName.tostring()
  Return $fileObject
  }
  else{Return $OpenFileDialog.FileName}
}

#import the Active Directory Module
Import-Module ActiveDirectory

#get and import csv list of users
$userlist = Import-Csv (get-filename -Filter "csv" -Title "Select CSV list with Accounts to Change" -Obj)

#New UPN Suffix
$upnsuffix = "@studentconnections.org"

$count = 0
#loop to license each user in list
FOREACH($user in $userlist)
{
	$count++
	#shows status bar
	Write-Progress -Activity ("Setting new UPN for . . ."+$user.UserPrincipalName) -Status "Scanned: $count of $($userlist.UserPrincipalName.Count)" -PercentComplete ($count/$userlist.UserPrincipalName.Count*100)
	#checks if field contains email address
	if($user.EmailAddress -ne $null -and $user.EmailAddress -ne "")
	{
		#remove the domain from name for get and set of UPN
		$upnname = (($user.EmailAddress).split('@'))[0]
		#get the AD User Object
		$aduser = Get-ADUser $user.sAMAccountName
		#Verify that aduser has a User Object before running the set command
		if($aduser -ne $null -and $aduser -ne "")
		{
            write-host "Setting new UPN $upnname$upnsuffix"
			$aduser|set-aduser -UserPrincipalName "$upnname$upnsuffix"
			$aduser|set-aduser -EmailAddress "$upnname$upnsuffix"
        }
		elseif($aduser -eq $null -or $aduseruser -eq ""){Write-Host "unable to find aduser for $user.sAMAccountName"}
	}
	#displays message if no email address was provided
	elseif($user.EmailAddress -eq $null -or $user.EmailAddress -eq ""){Write-Host "no EmailAddress provided for $user"}
	$user = $null
}