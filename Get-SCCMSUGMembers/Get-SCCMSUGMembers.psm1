<#
.SYNOPSIS
   Connect to SCCM Drive
   Get SCCM Software Update Group ArticleID Members

.DESCRIPTION
	Invoke-ConnectSCCM to Connect
	Get-SCCMSUGMembers reads Article IDs from a given software update group
	
	
.PARAMETER 
    $SUG
	
.EXAMPLE
     Get-SCCMSUGMembers -SUG "All Software Updates"

.INPUTS
   -SUG <SCCM Collection Name>

.OUTPUTS
    KB Article IDs as $SUGMembers

.NOTES
    Author: Carl Melander	
	
#>

#Get current location to use as return after doing SCCM Functions
$global:userSiteCode = Get-Location 

Function Invoke-ConnectSCCM { 

	# Site configuration
	$global:SiteCode = "Site Code" # Site code 
	$ProviderMachineName = "sccm.machine.com" # SMS Provider machine name

	# Import the ConfigurationManager.psd1 module 
	if((Get-Module ConfigurationManager) -eq $null) {
		Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" 
	}

	# Connect to the site's drive if it is not already present
	if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
		New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName 
	}

}
Function Get-SCCMSUGMembers {

	param(
			[Parameter(Mandatory=$true)]
			[String]$SUG #This is the software update group target	 	 
		)
	
	$SUGMembers = @()
	
	#Status Notification
	Write-host -nonewline -foregroundcolor Black -backgroundcolor DarkYellow ("   LOADING:")
	write-host (" Reading Software Update Group: $SUG")
	
	#Connect to SCCM site
	Invoke-ConnectSCCM; Set-Location "$($SiteCode):\" 

	#Get KB list from SUG
	$CMPSSuppressFastNotUsedCheck = $true
	$SUGdata = Get-CMSoftwareUpdateGroup -Name $SUG | Get-CMSoftwareUpdate
		
	$SUGHash = @{
		ArticleID = $SUGdata | Select-Object -ExpandProperty ArticleID 
		Title = $SUGdata | Select-Object -ExpandProperty LocalizedDisplayName
		DatePosted = $SUGdata | Select-Object -ExpandProperty DatePosted
	}
	
	$SUGMembers += for( $i=0; $i -lt $SUGHash.ArticleID.Count; $i++) {
		[PSCustomObject]@{ 
			ArticleID = $SUGHash.ArticleID[$i]
			DatePosted = $SUGHash.DatePosted[$i]
			Title = $SUGHash.Title[$i]
		}
	}
			
$global:SUGMembers = $SUGMembers

#Returns to user site for quality of life
$n = Get-Location; if ($n -ne $userSiteCode ) {Set-Location $userSiteCode}

}

