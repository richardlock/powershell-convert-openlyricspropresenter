# PowerShell script to import OpenLyrics xml files from a path, convert to ProPresenter format, and output as text files.
# Author: Richard Lock
# Date: 2021-03-02

# Variables
$ccliLicenceText = 'CCLI Licence No. xxxxx'
$inputFileFilter = '*.xml'
$inputPath = '.\input'
$logPath = ".\Convert-OpenLyricsProPresenter_$((Get-Date -Format s) -replace ':').txt"
$outputPath = '.\output'
$removeSongNumberLines = $true

# Import Logging module
Install-Module -Name Logging
Import-Module -Name Logging

# Configure logging
Add-LoggingTarget -Name Console
Add-LoggingTarget -Name File -Configuration @{
    Path = $logPath  
}
Write-Log -Level 'INFO' -Message 'Start log.'

# Function to update verse names to ProPresenter format
function Update-VerseName {
    param (
        [string]$verseName
    )
    switch -wildcard ($verseName) {
        '' { } # Work with blank verse order
        'b*' { $verseName -replace 'b','Bridge ' }
        'c*' { $verseName -replace 'c','Chorus ' }
        'e*' { $verseName -replace 'e[0-9]','Ending' }
        'i*' { $verseName -replace 'i[0-9]','Intro' }
        'p*' { $verseName -replace 'p[0-9]','PreChorus' }
        'v*' { $verseName -replace 'v','Verse ' }
        default { Write-Log -Level 'ERROR' -Message 'Unrecognised verse name.' }
    }  
}

# Get all xml files in the specified path
Get-ChildItem -Path $inputPath -File -Filter $inputFileFilter | ForEach-Object {
    Write-Log -Level 'INFO' -Message "Importing file '$_' ..."
    
    # Import file content as xml object
    try {
        [xml]$inputXml = Get-Content $_
    }
    catch {
        Write-Log -Level 'ERROR' -Message 'Error importing xml.'
    }

    $title = $inputXml.song.properties.titles.title
    # Select first title if more than one is specified
    if ($title.GetType().FullName -eq 'System.Object[]') {
        $title = $title[0]
    }

    $author = $inputXml.song.properties.authors.author -join ' | '
    
    $copyright = $inputXml.song.properties.copyright

    # Save space separated verse order as array
    $verseOrder = $inputXml.song.properties.verseOrder -split ' '
    
    $verses = $inputXml.song.lyrics.verse

    # Start with empty lyrics string
    $lyrics = ''

    # If no verse order is specified, get verse lyrics in order
    if ($verseOrder[0] -eq '') {
        foreach ($verse in $verses) {
            $lyrics += (Update-VerseName($verse.name)) + "`n"
            $lyrics += ($verse.lines.'#text' -join "`n") + "`n`n"
       }
    }
    # Else get verse lyrics according to specified verse order
    else {
        foreach ($verseName in $verseorder) {
            foreach ($verse in $verses) {
                if ($verseName -eq $verse.name) {
                    $lyrics += (Update-VerseName($verse.name)) + "`n"
                    $lyrics += ($verse.lines.'#text' -join "`n") + "`n`n"
                } 
            }
        }
    }

    # Remove lyric lines containing just a song number
    if ($removeSongNumberLines) {
        $lyrics = $lyrics -replace "(?ms)^[0-9]+$`n", ""
    }

    # Build output in ProPresenter import format
    $output = $title + "`n`n"
    $output += $lyrics
    $footer = "Copyright`n"
    $footer += $author + "`n"
    $footer += $copyright + "`n"
    $footer += $ccliLicenceText
    # Remove empty lines from footer
    $output += $footer -replace "`n`n", "`n"

    Write-Log -Level 'INFO' -Message "Writing output to file '$outputPath\$title.txt' ..."
    # Write output to file
    try {
        $output | Out-File -FilePath "$outputPath\$title.txt"
    }
    catch {
        Write-Log -Level 'ERROR' -Message 'Error writing output to file.'
    }
}

Wait-Logging
Write-Log -Level 'INFO' -Message 'Finish log.'