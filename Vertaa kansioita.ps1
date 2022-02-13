# Kansiot eivät saa olla tyhjiä. Muuten scripti ei toimi.
$fso = Get-ChildItem -Recurse -path "D:\Organisoidut"
$fsoBU = Get-ChildItem -Recurse -path "D:\Organisoitavat"
Compare-Object -ReferenceObject $fso -DifferenceObject $fsoBU