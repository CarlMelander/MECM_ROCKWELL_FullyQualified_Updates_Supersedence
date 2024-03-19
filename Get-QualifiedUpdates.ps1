<#
.SYNOPSIS
    Determines Which Rockwell Qualified patches are not in the SUG, but that are in ALL Updates

.DESCRIPTION
	Downloads the Rockwell Qualified patch XLS, selects a given tab, then filters for those that are Fully Qualified. 
	Get Supersedence data from the Microsoft Microsoft Security Response Center API
	Connect with SCCM to determine all sofware updates / SUG updates
	
.PARAMETER 
    None
	
.EXAMPLE
     .\Get-QualifiedUpdates.ps1 

.INPUTS
   None

.OUTPUTS
    Checked Values are exported to Script Root as "Qualified.CSV"

.NOTES
    Author: Carl Melander	
	
#>
Write-host -nonewline -foregroundcolor Black -backgroundcolor DarkYellow ("   VERSION:")
write-host (" Script V1.1")

$RockwellGroup = "CPR9 SR7" #Target Qualified Patch Release Group
$SUG = "Software Update Group" #SUG
$AllUpdates = "All Updates Group" #All Updates Reference Location
$PatchesSince = "1/1/2019" #Search Depth (1/1/2019 is the depth of the Supersedence data)

#Import Modules
$scriptroot = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
import-module -Name "$($scriptroot)\Get-SCCMSUGMembers\Get-SCCMSUGMembers"
import-module -Name "$($scriptroot)\Get-KBSupersedence\Get-KBSupersedence"
import-module -Name "$($scriptroot)\Get-RockwellQualifiedPatches\Get-RockwellQualifiedPatches"

#Get Rockwell Resource data
Get-RockwellQualifiedPatches -Group $RockwellGroup

#Remove KB from Rockwell KBs to match ArticleID
$RockwellKBs = $RockwellKBs.replace('KB','') 

#Get all updates
Get-SCCMSUGMembers $AllUpdates



#Filter for MS, Patched since $PatchesSince, and is number
$filteredAllPatches = $SUGMembers |
	 #Where-Object { $_.Title -match "Microsoft" } | 
	 Where-Object { $_.DatePosted -gt $PatchesSince } | 
	 Where-Object { $_.ArticleID -match "^[\d\.]+$" }
	

#$filteredAllPatches.ArticleID | Out-File "$($scriptroot)\SUG.csv"  #Error Checking

#Get SUG updates
Get-SCCMSUGMembers $SUG 



#Filter for those not in SUG, that are in rockwell
$output = $filteredAllPatches | 
	Where-Object { $_.ArticleID -notin $SUGMembers.ArticleID } |
	Where-Object { $_.ArticleID -in $RockwellKBs } | Sort-Object ArticleID -Unique
	
#$Output.ArticleID | Out-File "$($scriptroot)\Filtered.csv" #Error Checking

#Check Supersedence 
$CheckedValues = @()
For($i=0; $i -lt $output.ArticleID.Count;$i++){
	
	Get-KBSupersedence $output.ArticleID[$i]
	 
	$CheckedValues += [PSCustomObject]@{ 
		KBArticleID = "KB$($output.ArticleID[$i])"
		DatePosted = $output.DatePosted[$i]
		Title = $output.Title[$i]
		Status = if($Supersedence.Supersedence.Count -eq 1 ){$Supersedence.Supersedence}else{Supersedence = $Supersedence.Supersedence[0]}
		SupersededBy = if($Supersedence.Supersedence.Count -eq 1 ){$Supersedence.SupersededBy}else{Supersedence = $Supersedence.SupersededBy[0]}
			
	}
		
}
$CheckedValues | Export-CSV -path "$($scriptroot)\Qualified.csv" -NoTypeInformation
