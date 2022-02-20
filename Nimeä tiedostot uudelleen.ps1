$SomeFolder= Get-ChildItem "E:\PathToFolder"

$File_data = Get-Content .\pvm.txt -Encoding UTF8
$File_dataName = Get-Content .\nimet.txt -Encoding UTF8

for($file=0; $file -lt $SomeFolder.Length; $file++){ 
    #Write-Host ($SomeFolder[$file])
    #endregionWrite-Host ($File_data[$file])

    #$SomeFolder[$file] | Rename-Item -NewName { $_.BaseName + '--' + $File_data[$file] + '--'+ $_.Extension}
    #$SomeFolder[$file] | Rename-Item -NewName { $_.Name -replace $File_data[$file],'' }

    #$SomeFolder[$file] | Rename-Item -NewName { $File_dataName[$file] + '_' + $File_data[$file] + '_'+ $_.Extension }
}

#Get-ChildItem *.mp4 | Rename-Item -NewName { $_.Name -replace 'Copy of','' }