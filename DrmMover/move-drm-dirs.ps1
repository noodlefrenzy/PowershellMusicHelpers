#
# Script.ps1
#
$rootDir = $args[0]
$destDir = $args[1]
$wmplayer = New-Object -ComObject "WMPlayer.OCX"
$dirs = ls $rootDir -recurse | ? { $_.Name -like "*.wma" -and [bool]::Parse($wmplayer.newMedia($_.FullName).getItemInfo("Is_Protected")) } | sort -unique -property DirectoryName
foreach ($dir in $dirs) {
	$sub = $dir.DirectoryName.Substring($rootDir.Length+1)
	$curDest = [System.IO.Path]::Combine($destDir, $sub)
	Write-Host "Copying $sub to $curDest"
	$newDir = md $curDest -Force
	$files = [System.IO.Path]::Combine($dir.DirectoryName, "*.*")
	Copy-Item $files $curDest
}
