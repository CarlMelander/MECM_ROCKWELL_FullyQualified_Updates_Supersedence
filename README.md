# Description

This script was created to automate the population of a Software Update Group with Fully Qualified Rockwell Patches, including Supersedence 

# Functions
- Downloads the Rockwell Qualified patch XLS, selects a given tab, then filters for those that are Fully Qualified. 
- Get Supersedence data from the Microsoft Microsoft Security Response Center API
-	Connect with SCCM/MECM to determine all sofware updates / SUG updates

# Modules
## Get-QualifiedUpdates
Orchastrates the tasks and outputs KB Article IDs with supersedence
## Get-SCCMSUGMembers
Get-SCCMSUGMembers reads Article IDs from a given software update group
## Get-RockwellQualifiedPatches
Downloads the Rockwell Qualified patch XLS, selects a given tab, the filters for those that are Fully Qualified. 
## Get-KBSupersedence
MSRCSecurityUpdates is a PSGallery resource that collects data through the Microsoft Security Response Center API https://www.powershellgallery.com/packages/MsrcSecurityUpdates/1.9.6
This Module then parses this information into value pairs and returns recursive Supersedence information


	
