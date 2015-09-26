#
# convert_wma_mp3.ps1
#
$ffmpeg = 'E:\dev\tools\ffmpeg\bin\ffmpeg.exe'
$wmplayer = New-Object -ComObject "WMPlayer.OCX"

$rootDir = $args[0]
$destDir = $args[1]
$dirs = ls $rootDir -recurse | ? { $_.Name -like "*.wma" } | sort -unique -property DirectoryName
foreach ($dir in $dirs) {
	$sub = $dir.DirectoryName.Substring($rootDir.Length+1)
	$curDest = [System.IO.Path]::Combine($destDir, $sub)
	$wmaFiles = ls $dir.DirectoryName -Filter '*.wma'
	$wmaFiles = $wmaFiles | Where-Object { ![bool]::Parse($wmplayer.newMedia($_.FullName).getItemInfo("Is_Protected")) }
	if ($wmaFiles.Length -gt 0) {
		Write-Host "Converting WMA files from $sub to $curDest"
		$newDir = md $curDest -Force
		foreach ($file in $wmaFiles) {
			$infile = $file.FullName
			$movedFile = [System.IO.Path]::Combine($curDest, $file.Name)
			$outfile = [System.IO.Path]::Combine($file.DirectoryName, $file.Name.Replace($file.Extension, '.mp3'))
			$arguments = "-i `"$infile`" -id3v2_version 3 -f mp3 `"$outfile`" -y"
			Invoke-Expression "$ffmpeg $arguments"
			move-item $infile $movedFile
		}
	}
}
