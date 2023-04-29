# ModList Harvester - Export installed mods to CSV and markdown

# "Empire loves their damned lists" 
#                   - Ralof (4E 201)

# Author:   TroubleshooterNZ#0001
# Version:  v1.0.5
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
    Write-Progress -Activity "Query in Progress" -Status "($i/$max) Complete:" -PercentComplete (($i/$max)*100)
    
    if ($verbose) {Write-Verbose "$($data[$i]),$($myMetas[$i].DirectoryName)"}
    [PSCustomObject]@{
        installedFiles = $data[$i].installedFiles
        General = $data[$i].General
        Directory = $myMetas[$i].Directory.Name
    }
    Start-Sleep -Milliseconds 0
}

if ($results) {
     # csv output file headers
     $csvHeader = "modID,fileID,Name,URL"
     $OutputFile = 'myInstalledMods.csv'
     if ($GenerateCSV) {
     $csvHeader | Set-Content $OutputFile  # CLOBBERS 
    }
    if ($GenerateMarkDown) {
    # markdown output file headers
    $mdheader[0] | Set-Content  $mdFilename     #CLOBBER
    $mdheader[1] | Add-Content $mdFilename
    }
    # build results table
    $results | ForEach-Object {
        if ($null -ne $_.installedFiles["1\modid"] -and $null -ne $_.installedFiles["1\fileid"]){
            $a = $_.installedFiles["1\modid"]                                   #$_["installedFiles"]["1\modid"]
            $b = $_.installedFiles["1\fileid"]                                  #$_["installedFiles"]["1\fileid"]
            $c = $_.Directory.ToString()
            $link = $aurl[0]+$gameName+$aurl[1]+$a+$aurl[2]+$b+$aurl[3]
            $hyper = "[Download]($link)"
            $concat = "|"+$a+"|"+$b+"|"+$c+"|"+$hyper+"|"
            $concatcsv = $a+","+$b+","+$c+","+$link
        } elseif ($null -ne $_.General["modid"] ) {
            $a = $_.General["modid"]                                    #$_["installedFiles"]["1\modid"]
            $b = "0"                                                    #$_["installedFiles"]["1\fileid"]
            $c = $_.Directory.ToString()
            $link = $aurl[0]+$gameName+$aurl[1]+$a
            if ($null -ne $_.General["url"]) {
                $tempurl = $_.General["url"]
                $hyper = "[About]($tempurl)"
                $concat = "|"+$a+"|"+$b+"|"+$c+"|"+$hyper+"|"
            } elseif ($null -eq $_.General["url"] -Or $c -contains "_separator") {
                #$tempurl = "Category Separator"
                #$hyper = ""
                $concat = "|"+$a+"|"+$b+"|"+$c+"|Spacer|"
                $concatcsv = [string]$a+","+[string]$b+","+$c+","+"Category spacer"
            } else {
                $concat = "|"+$a+"|"+$b+"|"+$c+"|"+$link+"|"
                $concatcsv = [string]$a+","+[string]$b+","+$c+","+$link
            }
            
            $concatcsv = [string]$a+","+[string]$b+","+$c+",,"
        }
        if ($GenerateMarkDown) {
        $concat | Add-Content $mdfilename
        } 
        if ($GenerateMarkDown) {
        $concatcsv | Add-Content $OutputFile
        }
    }
}
((Get-Date -UFormat %s) - $ostime).ToString() +" seconds to generate." | Write-Host

