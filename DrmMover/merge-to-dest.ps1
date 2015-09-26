#
# merge_to_dest.ps1
#
param(
	[Parameter(Mandatory=$True,Position=0)]
	[string]$dest,
	
	[Parameter(mandatory=$True, position=1, ValueFromRemainingArguments=$true)]$srcs,
  
	[switch]$copyDRM,
	[switch]$overwrite,
	[switch]$convert,
	[switch]$dryrun
)

if ($srcs.Length -eq 0) {
	Write-Error "Must provide source directories"
	Exit 1
}
Write-Host "Merging ($srcs) into $dest (Options: " ("Copy DRM, ","")[!$copyDRM] ("OVR","")[!$overwrite] ("Convert","")[!$convert] ")"

$ffmpeg = 'E:\dev\tools\ffmpeg\bin\ffmpeg.exe'
$wmplayer = New-Object -ComObject "WMPlayer.OCX"

foreach ($rootDir in $srcs) {
	$dirs = ls $rootDir -recurse | sort -unique -property DirectoryName
	foreach ($dirInfo in $dirs) {
		$dir = $dirInfo.DirectoryName
		$sub = $dir
		if ($dir -ne $rootDir) {
			$sub = $dir.Substring($rootDir.Length+1)
		}
		$curDest = [System.IO.Path]::Combine($dest, $sub)
		$wmaFiles = ls $dir -Filter '*.wma'
		if (!$copyDRM) {
			$wmaFiles = $wmaFiles | Where-Object { ![bool]::Parse($wmplayer.newMedia($_.FullName).getItemInfo("Is_Protected")) }
		}
		$nonWmaFiles = ls $dir -Exclude '*.wma' | Where-Object { -not ($_.Attributes -band [System.IO.FileAttributes]::Directory) }
		if ($wmaFiles.Length -gt 0 -or $nonWmaFiles.Length -gt 0) {
			Write-Host "Merging $dir into $curDest"
			if (!$dryrun) { $newDir = md $curDest -Force }
			$allFiles = $wmaFiles + $nonWmaFiles
			foreach ($file in $allFiles) {
				$srcFile = $file.FullName
				if ($convert -and ($file.Extension -eq '.wma' -or $file.Extension -eq '.flac')) {
					$destFile = [System.IO.Path]::Combine($curDest, $file.Name.Replace($file.Extension, '.mp3'))
					if ($overwrite -or -not [System.IO.File]::Exists($destFile)) {
						Write-Host "Converting $srcFile to $destFile"
						$arguments = "-i `"$srcFile`" -id3v2_version 3 -f mp3 `"$destFile`" -y"
						if (!$dryrun) { Invoke-Expression "$ffmpeg $arguments" }
					}
				} else {
					$destFile = [System.IO.Path]::Combine($curDest, $file.Name)
					if ($overwrite -or -not [System.IO.File]::Exists($destFile)) {
						Write-Host "Copying $srcFile to $destFile"
						if (!$dryrun) { Copy-Item $srcFile $destFile }
					}
				}
			}
		}
	}
}

