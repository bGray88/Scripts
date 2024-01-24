$dryRun = $TRUE

$oldPrefix = 'G:\Disc'
$newPrefix = 'I:\Game\Disc\PC MS-DOS'
$searchPath = "C:\Users\Brandon\Downloads\_TEMP\"
$shell = new-object -com wscript.shell

if ( $dryRun ) 
{
   write-host "Executing dry run" -foregroundcolor green -backgroundcolor black
} 
else 
{
   write-host "Executing real run" -foregroundcolor red -backgroundcolor black
}

$files = dir $searchPath -filter *.lnk -recurse

foreach ( $filepath in $files ) 
{
	
	$lnk = $shell.createShortcut( $filepath.fullname )
	$oldTargetPath  = $lnk.TargetPath
	$oldArgsPath    = $lnk.Arguments
	$lnkRegex = [regex]::escape( $oldPrefix )
	
	if ( $dryRun ) 
	{
		write-host "Dry Run: $($oldTargetPath) => $($lnkRegex)" -foregroundcolor green -backgroundcolor black
	} 
	else 
	{
		write-host "Updating if match: $($oldTargetPath) => $($lnkRegex)" -foregroundcolor red -backgroundcolor black
	}
	
	write-host ( $oldTargetPath -match $lnkRegex ) -foregroundcolor yellow -backgroundcolor black

	if ( $oldTargetPath -match $lnkRegex ) 
	{
		$targetPath = $oldTargetPath -replace $lnkRegex, $newPrefix
		$targetArgs = $oldArgsPath -replace $lnkRegex, $newPrefix
		$startPath  = Split-Path $targetPath -Parent
		
		$dosboxRegex = [regex]::escape( 'Dosbox' )
		$dosboxPath  = Split-Path $startPath -Leaf
				
		if ( $dosboxPath -match $dosboxRegex )
		{
			write-host "Updating Dosbox Game: $($dosboxPath) == $($dosboxRegex)" -foregroundcolor red -backgroundcolor black
			$gamePath = Split-Path $startPath -Parent
		}
		else
		{
			write-host "Updating Standard Game: $($dosboxPath) != $($dosboxRegex)" -foregroundcolor red -backgroundcolor black
			$gamePath = $startPath
		}
		
		$fileName = Split-Path $gamePath -Leaf
		$iconPath = $gamePath + '\' + $fileName + ".ico"

		write-host "Found: " + $filepath.fullname -foregroundcolor yellow -backgroundcolor black
		write-host "  Replace: $($oldTargetPath) $($oldArgsPath)"
		write-host "  Target:  $($targetPath) $($targetArgs)"
		write-host "  Start:   $($startPath)"
		write-host "  Icon:    $($iconPath)"

		if ( !$dryRun ) 
		{
			$lnk.TargetPath       = $targetPath
			$lnk.Arguments        = $targetArgs
			$lnk.WorkingDirectory = $startPath
			$lnk.IconLocation 	  = $iconPath
			$lnk.Save()
		}
	}
}