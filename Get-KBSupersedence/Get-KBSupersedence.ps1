<#
.SYNOPSIS
    Reads the MSRC API and parses out the Microsoft Security Update details as KBArticle, Supersedence value pairs

.DESCRIPTION
	MSRCSecurityUpdates is a PSGallery resource that collects data through the Microsoft Security Response Center API
    https://www.powershellgallery.com/packages/MsrcSecurityUpdates/1.9.6
	
	Get-KBSupersedenceValuePairs function parses this information into value pairs
	
	Get-KBSupersedence returns recursive Supersedence information for <KBArticle> as an object
	
.PARAMETER 
    (Mandatory) -testID <KBArticle> 
	
.EXAMPLE
     Get-KBSupersedence -testID 3172514

.INPUTS
   <KBArticle>

.OUTPUTS
    $SupersedenceValuePairs
	$Supersedence

.NOTES
    Author: Carl Melander	

#>

Function Get-KBSupersedenceValuePairs {
	
	Write-host -nonewline -foregroundcolor Black -backgroundcolor Green ("  DOWNLOAD:")
	write-host (" Getting MS Msrc Data")
	
	#default search range '2019-Jan' YYYY-mmm
	$searchDepth ='2019-Jan'

	#Check to see if MsrcSecurityUpdates module is installed
	$MsrcInstalled = Get-InstalledModule | Where-object { $_ -match "MsrcSecurityUpdates" }
	if($MsrcInstalled -eq $null){Install-Module MSRCSecurityUpdates -Force -Scope CurrentUser}

	#Import MsrcSecurityUpdates Module (Microsoft Update Catalog API - Source PSGallery)
	Import-Module -Name MsrcSecurityUpdates -Force

	#Get all published Security Updates within the search depth range
	$Months = Get-MsrcSecurityUpdate 
	$Months = $Months.ID | Where-Object { $_ -gt $searchDepth }

	for($p=0; $p -lt $Months.count;$p++ ){
		
		#Status Message for process
		$progress = ([int](($p/$months.count)*10000)/100)
		Write-Progress -Activity "Retreiving Security Update Information" -Status "$($progress)% Complete" -PercentComplete $progress
		
		$MSlist = Get-MsrcCvrfDocument -ID $months[$p] 

		$MsrcHashTable =  @{
			KBArticle = $MSlist.Vulnerability.Remediations.Description.Value | Foreach-Object { if($_ -match "^[\d\.]+$"){$_}else{"NA"} } 
			Supersedence = $MSList.Vulnerability.Remediations.Supercedence | Foreach-Object { if($_ -eq $null){ "NA" } else { $_} } 
		}

		$SupersedenceValuePairs += For( $i=0; $i -lt $MsrcHashTable.KBArticle.Count; $i++ ) {
			 [PSCustomObject]@{
				KBArticle = $MsrcHashTable.KBArticle[$i]
				Supersedence = $MsrcHashTable.Supersedence[$i]
			} 
		}
	}
	
	#Closes Status Message for process
	Write-Progress -Activity "Retreiving Security Updates" -Completed
	
	#Optimize: Remove Duplicates of KB & Supersedence value pairs
	$global:SupersedenceValuePairs = $SupersedenceValuePairs | Sort-Object -Property KBArticle, Supersedence -Unique
}

Function Get-KBSupersedence {

	param(
			[Parameter(Mandatory=$true)]
			[String]$testID		 
		)

	#Generates Supersedence Value Pairs to search against
	if( $SupersedenceValuePairs -eq $null ){ Get-KBSupersedenceValuePairs }
	
	Write-host -nonewline -foregroundcolor Black -backgroundcolor Cyan (" SEARCHING:")
	write-host (" Looking for Supersedence Information on Article ID: $testID")
	
	#Create Empty array object
	$Supersedence=@()

		#Search $SupersedenceValuePairs for $testID; add to $Supersedence
		for($i=0; $i -lt $SupersedenceValuePairs.Supersedence.Count;$i++){
			if( ($SupersedenceValuePairs.Supersedence[$i] -match $testID) -and ($SupersedenceValuePairs.Supersedence[$i] -ne "NA") ){		
				$Supersedence += [PSCustomObject] @{					
						KBArticle = "KB$($testID)"
						Supersedence = "Superseded"
						SupersededBy = "KB$($SupersedenceValuePairs.KBArticle[$i])"		
					}	
				#Recursive Search of Superseding KB					
				$testID = $SupersedenceValuePairs.KBArticle[$i]
				$i=-1
				$result = 1
			}	
		}
				
		#Validate Superseding KB as 'Current'
		for($i=0;$i -lt $SupersedenceValuePairs.KBArticle.Count;$i++){
			if($testID -match $SupersedenceValuePairs.KBArticle[$i]){
				$Supersedence += [PSCustomObject] @{
						
						KBArticle = "KB$($testID)"
						Supersedence = "Current"
						SupersededBy = "NA"
						
					}
					$i=$SupersedenceValuePairs.KBArticle.Count
					$result = 1
			}
			
		}
		
	#if no results tag, log 'No Data' 
	if($result -ne 1){
		$Supersedence += [PSCustomObject] @{
						
						KBArticle = "KB$($testID)"
						Supersedence = "Not in MSRC"
						SupersededBy = "NA"
						
					}
	}
	
#Output Supersedence as global
$global:Supersedence = $Supersedence

}