$dryRun = $TRUE

$oldPrefix      = 'G:\Game\Digital\GOG\GalaxyClient\Games'
$newPrefix      = 'I:\Game\Digital\GOG\GOG Galaxy\Games'
$staticFilepath = 'G:\Digital\GOG\GalaxyClient\GalaxyClient.exe'
$searchPath     = "C:\Users\Brandon\Downloads\_TEMP\"
$shell          = new-object -com wscript.shell

if ( $dryRun ) 
{
   write-host "Executing dry run" -foregroundcolor green -backgroundcolor black
} 
else 
{
   write-host "Executing real run" -foregroundcolor red -backgroundcolor black
}

$lnkFiles = dir $searchPath -filter *.lnk -recurse
foreach ( $lnkFilepath in $lnkFiles ) 
{
	
	$lnk            = $shell.createShortcut( $lnkFilepath.fullname )
	$oldTargetPath  = $lnk.TargetPath
	$oldArgsPath    = $lnk.Arguments
	$lnkRegex       = [regex]::escape( $oldPrefix )
	$staticRegex      = [regex]::escape( $staticFilepath )
	
	if ( $dryRun ) 
	{
		write-host "Dry Run: $($oldTargetPath) => $($lnkRegex)" -foregroundcolor green -backgroundcolor black
	} 
	else 
	{
		write-host "Updating if match: $($oldTargetPath) => $($lnkRegex)" -foregroundcolor yellow -backgroundcolor black
	}
	
	write-host ( $oldTargetPath -match $lnkRegex ) -foregroundcolor yellow -backgroundcolor black

	if ( $oldTargetPath -match $lnkRegex ) 
	{
		$targetPath = $oldTargetPath -replace $lnkRegex, $newPrefix
		$targetArgs = $oldArgsPath -replace $lnkRegex, $newPrefix
		$startPath  = Split-Path $targetPath -Parent
		$exeName    = Split-Path $targetPath -Leaf
		
		$dosboxRegex = [regex]::escape( 'dosbox' )
		$scummRegex  = [regex]::escape( 'scummvm' )
		$runRegex    = [regex]::escape( 'run' )
		$systemRegex = [regex]::escape( 'system' )
		$dataRegex   = [regex]::escape( 'gamedata' )
		$emulName    = Split-Path $startPath -Leaf
				
		if ( $emulName -match $dosboxRegex -or $emulName -match $scummRegex -or $emulName -match $runRegex -or $emulName -match $systemRegex )
		{
			$gamePath = Split-Path $startPath -Parent
			$fileName = Split-Path $gamePath -Leaf
			write-host "Updating Emulating Game: $($emulName): $($fileName)" -foregroundcolor green -backgroundcolor black
		}
		else
		{
			$gamePath = $startPath
			$fileName = Split-Path $gamePath -Leaf
			write-host "Updating Standard Game: $($fileName)" -foregroundcolor green -backgroundcolor black
		}
		
		if (!(Test-Path -Path $gamePath)) 
		{
			$gameFiles = dir ( Split-Path $gamePath -Parent )
			foreach ( $folder in $gameFiles )
			{
				$folderRegex = [regex]::escape( $folder )
				if ( $fileName -match $folderRegex ) 
				{
					write-host "MATCH!!! $($filename)"
					$targetPath = ( Split-Path $startPath -Parent ) + '\' + $folder + '\' + $exeName
					$startPath  = ( Split-Path $startPath -Parent ) + '\' + $folder
					$gamePath   = ( Split-Path $gamePath -Parent ) + '\' + $folder
				}
			}
		}
		if (!(Test-Path -Path $gamePath)) 
		{
			$gameFiles = dir ( Split-Path $gamePath -Parent )
			foreach ( $folder in $gameFiles )
			{
				$fileRegex = [regex]::escape( $fileName )
				if ( $folder -match $fileRegex ) 
				{
					write-host "MATCH!!! $($filename)"
					$targetPath = ( Split-Path $startPath -Parent ) + '\' + $folder + '\' + $exeName
					$startPath  = ( Split-Path $startPath -Parent ) + '\' + $folder
					$gamePath   = ( Split-Path $gamePath -Parent ) + '\' + $folder
				}
			}
		}
		
		$gameFiles = dir $gamePath -filter *.ico
		foreach ( $file in $gameFiles )
		{
			if (!($file -match [regex]::escape( 'Support' )))
			{
				$iconPath = $gamePath + '\' + $file
			}
		}

		write-host "Found: " + $lnkFilepath.fullname -foregroundcolor yellow -backgroundcolor black
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
	elseif ( $oldTargetPath -match $staticRegex )
	{
		write-host "Bad Shortcut; Swap Necessary $($oldArgsPath)" -foregroundcolor red -backgroundcolor black
	}
	else
	{
		write-host "Game is up-to-date: $($oldTargetPath)" -foregroundcolor yellow -backgroundcolor black
	}
}