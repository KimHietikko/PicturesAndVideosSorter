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
$MediaCreatedColumn = 208
$MediaModifiedColumn = 3

# For older versions of Windos you may need to use these values:
#[reflection.assembly]::loadfile( "C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.Drawing.dll") 
#$MediaCreatedColumn = 191

$SourceRootPath = "C:\Users\kimhi\Desktop\Kuvat ja videot\Organisoitavat"
$DestinationRootPath = "C:\Users\kimhi\Desktop\Kuvat ja videot\Organisoitu"
$FileTypesToOrganize = @("*.jpg","*.jpeg","*.avi","*.mp4", "*.3gp", "*.mov", "*.png", "*.MTS", "*.gif", "*.M2TS")
$global:ConfirmAll = $false

function GetMediaCreatedDate($File) {
	$Shell = New-Object -ComObject Shell.Application
	$Folder = $Shell.Namespace($File.DirectoryName)
	$CreatedDate = $Folder.GetDetailsOf($Folder.Parsename($File.Name), $MediaCreatedColumn).Replace([char]8206, ' ').Replace([char]8207, ' ')

    if (!$CreatedDate) {
        if($File.Name.Contains('VID-')) {

            $Year = $File.Name.substring(4,4)
	        $Month = $File.Name.substring(8,2)
	        $Day = $File.Name.substring(10,2)
	        $Hour = '00'
	        $Minute = '00'
	        $Second = '00'
        
            $CreatedDate = [String]::Format("{0}/{1}/{2} {3}:{4}:{5}", $Year, $Month, $Day, $Hour, $Minute, $Second)
        }

        if($File.Name.Contains('WhatsApp Video ')) {

            $Year = $File.Name.substring(15,4)
	        $Month = $File.Name.substring(20,2)
	        $Day = $File.Name.substring(23,2)
	        $Hour = $File.Name.substring(29,2)
	        $Minute = $File.Name.substring(32,2)
	        $Second = $File.Name.substring(35,2)
        
            $CreatedDate = [String]::Format("{0}/{1}/{2} {3}:{4}:{5}", $Year, $Month, $Day, $Hour, $Minute, $Second)
        }

        if ($File.Name.IndexOf('-') -eq 4 -and $File.Name.substring(7,1).contains('-') -and $File.Name.substring(10,1).contains('_') -and $File.Name.substring(13,1).contains('-') -and $File.Name.substring(16,1).contains('-')) {
            
            $Year = $File.Name.substring(0,4)
	        $Month = $File.Name.substring(5,2)
	        $Day = $File.Name.substring(8,2)
	        $Hour = $File.Name.substring(11,2)
	        $Minute = $File.Name.substring(14,2)
	        $Second = $File.Name.substring(17,2)
        
            $CreatedDate = [String]::Format("{0}/{1}/{2} {3}:{4}:{5}", $Year, $Month, $Day, $Hour, $Minute, $Second)
        }

        if ($File.Name.Contains('.M2TS')) {
            
            $Year = $File.Name.substring(0,4)
	        $Month = $File.Name.substring(4,2)
	        $Day = $File.Name.substring(7,2)
	        $Hour = $File.Name.substring(10,2)
	        $Minute = $File.Name.substring(13,2)
	        $Second = $File.Name.substring(16,2)
        
            $CreatedDate = [String]::Format("{0}/{1}/{2} {3}:{4}:{5}", $Year, $Month, $Day, $Hour, $Minute, $Second)
        }
    }

    if (!$CreatedDate) {
        $CreatedDate = $Folder.GetDetailsOf($Folder.Parsename($File.Name), $MediaModifiedColumn).Replace([char]8206, ' ').Replace([char]8207, ' ')
    }

	if ($null -ne ($CreatedDate -as [DateTime])) {
		return [DateTime]::Parse($CreatedDate).getDateTimeFormats()[8]
	} else {
		return $null
	}
}

function ConvertAsciiArrayToString($CharArray) {
	$ReturnVal = ""
	foreach ($Char in $CharArray) {
		$ReturnVal += [char]$Char
	}
	return $ReturnVal
}

function GetDateTakenFromExifData($File) {
    $ascii = $true
    $lastWriteTime = $false

	$FileDetail = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $File.Fullname 
	$DateTimePropertyItem

    if ($FileDetail.PropertyIdList.Contains(36867)) {
        $DateTimePropertyItem = $FileDetail.GetPropertyItem(36867)
    }
	
	if (!$DateTimePropertyItem) {
        if ($FileDetail.PropertyIdList.Contains(306)) {
            $DateTimePropertyItem = $FileDetail.GetPropertyItem(306)
        }
	}

    if (!$DateTimePropertyItem) {

        if($File.Name.Contains('WhatsApp Image ')) {

            $Year = $File.Name.substring(15,4)
	        $Month = $File.Name.substring(20,2)
	        $Day = $File.Name.substring(23,2)
	        $Hour = $File.Name.substring(29,2)
	        $Minute = $File.Name.substring(32,2)
	        $Second = $File.Name.substring(35,2)
        
            $ascii = $false

            $DateTimePropertyItem = [String]::Format("{0}/{1}/{2} {3}:{4}:{5}", $Year, $Month, $Day, $Hour, $Minute, $Second)
        }

        if($File.Name.Contains('Screenshot_') -and $File.Name.Contains('.jpg')) {

            $Year = $File.Name.substring(11,4)
	        $Month = $File.Name.substring(15,2)
	        $Day = $File.Name.substring(17,2)
	        $Hour = $File.Name.substring(20,2)
	        $Minute = $File.Name.substring(22,2)
	        $Second = $File.Name.substring(24,2)
        
            $ascii = $false

            $DateTimePropertyItem = [String]::Format("{0}/{1}/{2} {3}:{4}:{5}", $Year, $Month, $Day, $Hour, $Minute, $Second)
        }

        if($File.Name.Contains('Screenshot_') -and $File.Name.Contains('.png')) {

            $Year = $File.Name.substring(11,4)
	        $Month = $File.Name.substring(16,2)
	        $Day = $File.Name.substring(19,2)
	        $Hour = $File.Name.substring(22,2)
	        $Minute = $File.Name.substring(25,2)
	        $Second = $File.Name.substring(28,2)
        
            $ascii = $false

            $DateTimePropertyItem = [String]::Format("{0}/{1}/{2} {3}:{4}:{5}", $Year, $Month, $Day, $Hour, $Minute, $Second)
        }

        if ($File.Name.IndexOf('-') -eq 4 -and $File.Name.substring(7,1).contains('-') -and $File.Name.substring(10,1).contains(' ') -and $File.Name.substring(13,1).contains('.') -and $File.Name.substring(16,1).contains('.') -and ($File.Name.Contains('.png') -or $File.Name.Contains('.gif'))) {
            
            $Year = $File.Name.substring(0,4)
	        $Month = $File.Name.substring(5,2)
	        $Day = $File.Name.substring(8,2)
	        $Hour = $File.Name.substring(11,2)
	        $Minute = $File.Name.substring(14,2)
	        $Second = $File.Name.substring(17,2)
        
            $ascii = $false

            $DateTimePropertyItem = [String]::Format("{0}/{1}/{2} {3}:{4}:{5}", $Year, $Month, $Day, $Hour, $Minute, $Second)
        }
    }

    if (!$DateTimePropertyItem) {
        $DateTimePropertyItem = $File.LastWriteTime.GetDateTimeFormats()[12]
        $ascii = $false
        $lastWriteTime = $true
    }
	
	$FileDetail.Dispose()

    if ($ascii -eq $true) {
        $Year = ConvertAsciiArrayToString $DateTimePropertyItem.value[0..3]
	    $Month = ConvertAsciiArrayToString $DateTimePropertyItem.value[5..6]
	    $Day = ConvertAsciiArrayToString $DateTimePropertyItem.value[8..9]
	    $Hour = ConvertAsciiArrayToString $DateTimePropertyItem.value[11..12]
	    $Minute = ConvertAsciiArrayToString $DateTimePropertyItem.value[14..15]
	    $Second = ConvertAsciiArrayToString $DateTimePropertyItem.value[17..18]
    } 
    
    if ($ascii -eq $false -and $lastWriteTime -eq $true) {
        $Year = $DateTimePropertyItem.substring(0,4)
	    $Month = $DateTimePropertyItem.substring(5,2)
	    $Day = $DateTimePropertyItem.substring(8,2)
	    $Hour = $DateTimePropertyItem.substring(11,2)
	    $Minute = $DateTimePropertyItem.substring(14,2)
	    $Second = $DateTimePropertyItem.substring(17,2)
    }

	
	
	$DateString = [String]::Format("{0}/{1}/{2} {3}:{4}:{5}", $Year, $Month, $Day, $Hour, $Minute, $Second)
	
	if ($null -ne ($DateString -as [DateTime])) {
		return [DateTime]::Parse($DateString).getDateTimeFormats()[8]
	} else {
		return $null
	}
}

function GetCreationDate($File) {
	switch ($File.Extension) { 
        ".jpg" { $CreationDate = GetDateTakenFromExifData($File) } 
        ".jpeg" { $CreationDate = GetDateTakenFromExifData($File) } 
        ".png" { $CreationDate = GetDateTakenFromExifData($File) } 
        ".gif" { $CreationDate = GetDateTakenFromExifData($File) } 
        default { $CreationDate = GetMediaCreatedDate($File) }
    }
	return $CreationDate
}

function BuildDesinationPath($Path, $Date) {
	return [String]::Format("{0}\{1}\{2}_{3}\{4}", $Path, $Date.Year, $Date.ToString("MM"), $Date.ToString("MMMM"), $Date.ToString("dd.MM.yyyy"))
}

$RandomGenerator = New-Object System.Random
function BuildNewFilePath($Path, $Date, $Extension) {
	return [String]::Format("{0}\{1}_{2}_{3}{4}", $Path, $Date.ToString("yyyy-MM-dd_HH-mm-ss"), $RandomGenerator.Next(100, 10000).ToString(), $RandomGenerator.Next(100, 10000).ToString(), $Extension)
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
	$CreationDate = GetCreationDate($File)
    $CreationDate = [DateTime]::Parse($CreationDate)
	if ($null -ne ($CreationDate -as [DateTime])) {
		$DestinationPath = BuildDesinationPath $DestinationRootPath $CreationDate
		CreateDirectory $DestinationPath
		$NewFilePath = BuildNewFilePath $DestinationPath $CreationDate $File.Extension
		
		Write-Host $File.FullName -> $NewFilePath
		if (!(Test-Path $NewFilePath)) {
			Move-Item -literalpath $File.FullName $NewFilePath
		} else {
			Write-Host "Unable to rename file. File already exists. "
			ConfirmContinueProcessing
		}
	} else {
		Write-Host "Unable to determine creation date of file. " $File.FullName
		ConfirmContinueProcessing
	}
} 

Write-Host "Done"