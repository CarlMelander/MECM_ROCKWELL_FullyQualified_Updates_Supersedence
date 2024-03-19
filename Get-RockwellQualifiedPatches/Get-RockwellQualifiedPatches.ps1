<#
.SYNOPSIS
    Downloads and Reads the Rockwell Qualified XLS and parses out the Fully Qualified Microsoft updates for a given Tab Name

.DESCRIPTION
	Downloads the Rockwell Qualified patch XLS, selects a given tab, the filters for those that are Fully Qualified. 
	Closes the Excel process, and removes the XLS from the default temp location $OutputPatchPath
	
.PARAMETER 
    (Mandatory) -Group <Tab Name>
	
.EXAMPLE
     Get-RockwellQualifiedPatches -Group "CPR9 SR7"

.INPUTS
   <Tab Name>

.OUTPUTS
    $RockWellKBs

.NOTES
    Author: Carl Melander	
	
#>

Function Get-RockwellQualifiedPatches {
		
		param(
				[Parameter(Mandatory=$true)]
				[String]$Group #This is qualified patch group tab 	 	 
			)
		
		#This is the Tab on the Rockwell Patch XLS to filter against
		[String]$tab = $Group

		#Rockwell Patch File Name (This is the default name from Rockwell and shouldn't change)
		[String]$inputfile = "PQual_Queries.xls" 

		#Rockwell patch list location
		[String]$Rockwell_site = "https://www.rockwellautomation.com/ms-patch-qualification/Monthly%20Download/"
		[String]$Rockwell_target = "$($Rockwell_site)$inputfile"

		#Temp Location for Rockwell XLS
		[string]$OutputPatchPath = "c:\temp\" 

		#Combines Output Patch Path with Rockwell XLS file name 
		[String]$fullpath = "$($OutputPatchPath)$inputfile"

	#Create Objects, collect current state
	Write-host -nonewline -foregroundcolor Black -backgroundcolor DarkYellow ("   LOADING:")
	write-host (" Initializing Rockwell Script")
			
		#Create new-objects for KB list
		$RockWellKBs = @()
		
		#Get all XLS instances (script stops all xls not in this list)
		$excelID = get-process excel -ErrorAction SilentlyContinue -ErrorVariable ProcessError | select -expandproperty ID
		if($processerror){$excelID="none"}

		#Int XLS ComObject
		$XLSfile = New-Object -ComObject Excel.Application 
		$XLSfile.visible= $False

	#download rockwell patch xls from website
	Write-host -nonewline -foregroundcolor Black -backgroundcolor Green ("  DOWNLOAD:")
	write-host (" Downloading Rockwell Patch List to $($OutputPatchPath)$inputfile")

		Start-BitsTransfer -Source $Rockwell_target -Destination $OutputPatchPath

	#Open XLS, Select Tab, Set Range 
	Write-host -nonewline -foregroundcolor Black -backgroundcolor DarkYellow ("   LOADING:")
	write-host (" Reading $tab Patches from $fullpath")

		$xls = $XLSfile.workbooks.open($fullpath)
		$xlstab = $xls.sheets.item($tab) 
		$objRange = $xlstab.UsedRange
		$RowCount = $objRange.Rows.Count
		$kbrange = $xlstab.range("A4:A$RowCount") 
		$qrange = $xlstab.range("E4:E$RowCount") 
		$kbrange = $kbrange.formula2
		$qrange = $qrange.formula2
		
	Write-host -nonewline -foregroundcolor Black -backgroundcolor Cyan (" FILTERING:")
	write-host (" Filtering Fully Qualified Patches")


	for ($i=0; $i -lt $qrange.count; $i++){ 
		if ($qrange[$i,1] -match "Fully Qualified"){
			$RockWellKBs += $kbrange[$i,1]
		}
	}
	$global:RockWellKBs = $RockWellKBs

	##### Close and clean up XLS Loose ends
	get-process excel | where-object ID -notin $excelID | stop-process -force 
	Write-host -nonewline -foregroundcolor Black -backgroundcolor Yellow ("  CLEAN-UP:")
	write-host (" Removing $fullpath")
	start-sleep -s 5
	Remove-Item $fullpath
}

