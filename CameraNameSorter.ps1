# ==============================================================================================
# 
# Microsoft PowerShell Source File 
# 
# This script will organize photo and video files by renaming the file based on the date the
# file was created and moving them into folders based on the year and month. It will also append
# a random number to the end of the file name just to avoid name collisions. The script will
# look in the SourceRootPath (recursing through all subdirectories) for any files matching
# the extensions in FileTypesToOrganize. It will rename the files and move them to folders under
# DestinationRootPath, e.g. DestinationRootPath\2011\02_February\2011-02-09_21-41-47_680.jpg
# 
# JPG files contain EXIF data which has a DateTaken value. Other media files have a MediaCreated
# date. 
# 
# Original Author: ToddRopog
# Edited: KimHietikko
#
# ============================================================================================== 


# These value work for Windows 10:
[reflection.assembly]::loadfile( "C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Drawing.dll") 

# For older versions of Windos you may need to use these values:
#[reflection.assembly]::loadfile( "C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.Drawing.dll") 
#$MediaCreatedColumn = 191

$SourceRootPath = "E:\Kuvia vuosien saatossa"
$DestinationRootPath = "D:\Organisoidut"
$FileTypesToOrganize = @("*.jpg","*.jpeg")
$global:ConfirmAll = $false


function GetCameraFromExifData($File) {
	$FileDetail = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $File.Fullname 
	$CameraManufacturerPropertyItem
	
    if ($FileDetail.PropertyIdList.Contains(271)) {
		$CameraManufacturerPropertyItem = ConvertAsciiArrayToString $FileDetail.GetPropertyItem(271).Value[0..($FileDetail.GetPropertyItem(271).Value.Length-1)]
		$CameraManufacturerPropertyItem -replace '[\W]', ''
    }

    return $CameraManufacturerPropertyItem
}
function ConvertAsciiArrayToString($CharArray) {
	$ReturnVal = ""
	foreach ($Char in $CharArray) {
		$ReturnVal += [char]$Char
	}
	return $ReturnVal
}

function GetCamera($File) {
	switch ($File.Extension) {
        default { $Camera = GetCameraFromExifData($File) }
    }
	return $Camera
}

function BuildDesinationPath($Path, $CameraName) {
	return [String]::Format("{0}\{1}", $Path, $CameraName)
}

function BuildNewFilePath($Path, $FileName) {
	return [String]::Format("{0}\{1}", $Path, $FileName)
}

function CreateDirectory($Path){
	if (!(Test-Path $Path)) {
		New-Item $Path -Type Directory
	}
}

function ConfirmContinueProcessing() {
	if ($global:ConfirmAll -eq $false) {
		$Response = Read-Host "Continue? (Y/N/A)"
		if ($Response.Substring(0,1).ToUpper() -eq "A") {
			$global:ConfirmAll = $true
		}
		if ($Response.Substring(0,1).ToUpper() -eq "N") { 
			break 
		}
	}
}

function GetAllSourceFiles() {
	return @(Get-ChildItem $SourceRootPath -Recurse -Include $FileTypesToOrganize)
}

Write-Host "Begin Organize"
$Files = GetAllSourceFiles
foreach ($File in $Files) {
	$CameraName = GetCamera($File)
	if ($CameraName -eq "FUJIFILM") {
		$DestinationPath = BuildDesinationPath $DestinationRootPath $CameraName[1]
		CreateDirectory $DestinationPath
		$NewFilePath = BuildNewFilePath $DestinationPath $File.Name
		
		Write-Host $File.FullName -> $NewFilePath
		if (!(Test-Path $NewFilePath)) {
			Copy-Item -literalpath $File.FullName $NewFilePath
		} else {
			Write-Host "Unable to rename file. File already exists. "
			ConfirmContinueProcessing
		}
	}
} 

Write-Host "Done"