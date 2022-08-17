Get-ChildItem "E:\Testi" *.txt | Rename-Item -NewName { $_.Name -replace 'Copy of','' }

Write-Host "Done"