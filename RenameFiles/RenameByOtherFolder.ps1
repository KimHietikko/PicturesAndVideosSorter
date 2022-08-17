$SourceRootPath= Get-ChildItem "E:\Testi"
$TargetRootPath= Get-ChildItem "E:\Testi2"

for($file=0; $file -lt $SourceRootPath.Length; $file++){ 
    Write-Host ($SourceRootPath[$file])
    Write-Host ($TargetRootPath[$file])

    $TargetRootPath[$file] | Rename-Item -NewName { $SourceRootPath[$file].BaseName + $SourceRootPath[$file].Extension}
}

Write-Host "Done"