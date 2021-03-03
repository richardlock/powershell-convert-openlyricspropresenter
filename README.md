Convert-OpenLyricsProPresenter PowerShell script
================================================

PowerShell script to import OpenLyrics xml files from a path, convert to ProPresenter format, and output as text files.

Usage
-----

- Clone this repository or download the script.
- Update the variables as required.
- Export input songs in OpenLyrics xml format to the path specified in `$inputPath`.<br>E.g. OpenLP > File > Export > Song
- Run `Convert-OpenLyricsProPresenter.ps1` using PowerShell.
- The script will output songs in ProPresenter text format to the path specified in `$outputPath`.
- Import the text files using ProPresenter.<br>E.g. ProPresenter > File > Import > File...

Variables
---------

| Variable | Default value | Description |
| ---------|---------------|-------------|
| `$ccliLicenceText` | 'CCLI Licence No. xxxxx' | CCLI licence text to include at the end of the output.<br>Replace 'xxxxx' with your licence number, or use an empty string to exclude. |
| `$inputFileFilter` | '*.xml' | Filter to match files to import. |
| `$inputPath` | '.\input' | Path for files to import.<br>Create a directory '.\input' in the same directory as the script or specify another location. |
| `$logPath` | ".\Convert-OpenLyricsProPresenter_$((Get-Date -Format s) -replace ':').txt" |Full path to log file containing files processed and errors. By default, logs are written to file and the console. |
| `$outputPath` | '.\output' | Path for files to export.<br>Create a directory '.\output' in the same directory as the script or specify another location. |
| `$removeSongNumberLines` | $true | Whether to remove lyric lines containing just a song number. |

