﻿$tdc="J:\Organisoidut kuvat ja videot"
$dirs = gci $tdc -directory -recurse | Where { (gci $_.fullName).count -eq 0 } | select -expandproperty FullName
$dirs | Foreach-Object { Remove-Item $_ }