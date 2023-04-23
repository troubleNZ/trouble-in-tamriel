# ModList Harvester - Export installed mods to CSV and markdown

# "Empire loves their damned lists" 
#                   - Ralof (4E 201)

# Author:   TroubleshooterNZ#0001
# Version:  v1.0.0
# Repo:     https://github.com/troubleNZ/trouble-in-tamriel
#---------------------------------------------------//---------------------------------------------
# PURPOSE:
# Query ModOrganizer Mods folder for list of directories containing 'meta.ini' , 
# read all results, and build a report with hyperlinks to the original source on NexusMods.
#---------------------------------------------------//---------------------------------------------

# Edit this / Config:

# Game Name as specified by NexusMods
$gameName   = "skyrimspecialedition"

# ModOrganizer's installed mods path:
$myMods     = 'V:\ModOrganizer\Skyrim Special Edition\mods'

# ModOrganizer Profile ini location (for load Order:)                       # todo #
#$myprofile = 'V:\ModOrganizer\Skyrim Special Edition\profiles\Default'


# File we are searching for
$iniFile    = "meta.ini"

# Markdown FileName to write to
$markdown   = 'modlist.MD'

$GenerateCSV        = $true
$GenerateMarkDown   = $true

# change to $true if you want it to print to the current powershell session.
$verbose            = $true

# ---------------------------------------------------//---------------------------------------------
# Do not edit below this line unless you know what you are doing
# ---------------------------------------------------//---------------------------------------------
# --------------------------------------------------------------------------------------------------

# Methods
function Get-IniContent ($filePath)
{
    $ini = @{}
    switch -regex -file $FilePath
    {
        "^\[(.+)\]" # Section
        {
            $section = $matches[1]
            $ini[$section] = @{}
            $CommentCount = 0
        }
        "^(;.*)$" # Comment
        {
            $value = $matches[1]
            $CommentCount = $CommentCount + 1
            $name = "Comment" + $CommentCount
            $ini[$section][$name] = $value
        }
        "(.+?)\s*=(.*)" # Key
        {
            $name,$value = $matches[1..2]
            $ini[$section][$name] = $value
        }
    }
    return $ini
}

function Out-IniFile($InputObject, $FilePath)
{
    $outFile = New-Item -ItemType file -Path $Filepath
    foreach ($i in $InputObject.keys)
    {
        if (!($($InputObject[$i].GetType().Name) -eq "Hashtable"))
        {
            #No Sections
            Add-Content -Path $outFile -Value "$i=$($InputObject[$i])"
        } else {
            #Sections
            Add-Content -Path $outFile -Value "[$i]"
            Foreach ($j in ($InputObject[$i].keys | Sort-Object))
            {
                if ($j -match "^Comment[\d]+") {
                    Add-Content -Path $outFile -Value "$($InputObject[$i][$j])"
                } else {
                    Add-Content -Path $outFile -Value "$j=$($InputObject[$i][$j])"
                }

            }
            Add-Content -Path $outFile -Value ""
        }
    }
}
# https://devblogs.microsoft.com/scripting/use-powershell-to-work-with-any-ini-file/


$timestamp = (Get-Date -Format o | ForEach-Object { $_ -replace ":", "." }).ToString()

$mdFilename = <#$timestamp + #> "_MO2MLMd_" + $markdown
#markdown table formatting
$mdheader = @("|modID|fileID|Name|Link|","|--|--|--|--|")

# url formatting
$aurl = @("https://www.nexusmods.com/","/mods/","?tab=files&file_id=","&nmm=1")

# measure how long it takes
$ostime = Get-Date -UFormat %s

# Gather meta.ini data from directories
$myMetas = Get-ChildItem -Path $myMods -Include $iniFile -Recurse -Depth 1

# Generate data for export
$data = @($myMetas | ForEach-Object { 
                Join-Path -Path $_.Directory -ChildPath ("\"+$iniFile) |`
                Get-IniContent $_})


[int]$max = $data.Count
if ([int]$data.Count -gt [int]$myMetas.Count) { $max = $myMetas.Count;}

$results = for ($i = 0; $i -lt $max; $i++)
{
    if ($verbose) {Write-Verbose "$($data[$i]),$($myMetas[$i].DirectoryName)"}
    [PSCustomObject]@{
        installedFiles = $data[$i].installedFiles
        General = $data[$i].General
        Directory = $myMetas[$i].Directory.Name
    }
}
$results

if ($GenerateCSV) {
    # csv output file headers
    $csvHeader = "modID,fileID,Name,URL"
    $OutputFile = 'myInstalledMods.csv'
    $csvHeader | Set-Content $OutputFile  # CLOBBERS 
                    
    # process the data and write to csv file
    $results | ForEach-Object {
            $a = $_.installedFiles["1\modid"]                               #$_["installedFiles"]["1\modid"]
            $b = $_.installedFiles["1\fileid"]                              #$_["installedFiles"]["1\fileid"]
            $c = $_.Directory                                               #$myMetas[$i].DirectoryName
            $link = $aurl[0]+$gameName+$aurl[1]+$a+$aurl[2]+$b+$aurl[3]
            #$_["installedFiles"].Values
            $concat = $a+","+$b+","+$c+","+$link
            $concat | Add-Content $OutputFile
    }
}
if ($GenerateMarkDown) {
    # markdown output file headers
    $mdheader[0] | Set-Content  $mdFilename     #CLOBBER
    $mdheader[1] | Add-Content $mdFilename

    # build for markdown
    $results | ForEach-Object {
        $a = $_.installedFiles["1\modid"]                                   #$_["installedFiles"]["1\modid"]
        $b = $_.installedFiles["1\fileid"]                                  #$_["installedFiles"]["1\fileid"]
        $c = $_.Directory                                                   #$myMetas[$i].DirectoryName
        $link = $aurl[0]+$a+$aurl[1]+$b+$aurl[2]
        $hyper = "[Link]($link)"
        $concat = "|"+$a+"|"+$b+"|"+$c+"|"+$hyper+"|"
        $concat | Add-Content $mdfilename
    }
}
((Get-Date -UFormat %s) - $ostime).ToString() +" seconds to generate." | Write-Host