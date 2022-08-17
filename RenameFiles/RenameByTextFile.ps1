$SomeFolder= Get-ChildItem "E:\Testi"

$File_data = Get-Content .\pvm.txt -Encoding UTF8

for($file=0; $file -lt $SomeFolder.Length; $file++){ 
    Write-Host ($SomeFolder[$file])
    Write-Host ($File_data[$file])

    $SomeFolder[$file] | Rename-Item -NewName { $_.BaseName + '--' + $File_data[$file] + '--'+ $_.Extension}
}

Write-Host "Done"