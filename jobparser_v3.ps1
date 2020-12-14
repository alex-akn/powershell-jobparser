
# Locates all .txt files in the current directory. Tries to parse 
# each ITEM block into a hashtable. Changes date format and filters
# job items by date. Saves resulting array of hashtables as CSV file.



# If False - one output file for each feed .txt file
# If True - one big output file
[bool] $combineFiles = $true

# Whether of not to output only the most recent day's jobs
# Set True to Yes
[bool] $filterByDate = $true

# Path to the output file
# Current directory by default
[string] $exportFile = "./parsedjobs.csv"

# Path to the directory with source files
# Current directory by default

$batches = "\\ad.xxx.net\WPS\NL\P\UD\200026\LC15MA\Home\My Documents_ds\batches"

$isDeployed = Test-Path -Path $batches
if($isDeployed){
    New-PSDrive -Name "Batches" -PSProvider "FileSystem" -Root $batches
    $workDir = "Batches:"
}
else {
    $workDir = "./"
}


# Take a hashtable and create a string representation of it. 
function Write-Line {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,
         ValueFromPipeline=$true)]
        [hashtable]$hash
    )
    Process
    {
        $descr = ""
        if($hash.Description)
        {
            $descr = '"' + $hash.Description + '"'
        }
        $line = "" + $hash.file + "," + $hash.ItemNr + "," + $hash.JobName + "," + $hash.Status + "," + $hash.Started + "," + $hash.OnDate + "," + $hash.OnDate_DATETIME + "," + $hash.LastRan + "," + $hash.LastRanOnDate + "," + $hash.LastRanOnDate_DATETIME + "," + $hash.ElapsedTime + "," + $descr
        return $line
        #Add-Content -Path $exportFile -Value $line
    }    
}


# Read each line of a text file until "Item #" is found.
# Combine 8 next lines into one string. Repeat.
function Read-File
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,
         ValueFromPipeline=$true)]
        [string] $filename
    )
    Process
    {
        $file = Join-Path -Path $workDir -ChildPath $filename
        Write-Host ("Reading $filename...")
        $content = Get-Content $file
        $jobs = @()
        $group = ""
        $isBlock = $false
        $current = 0
        foreach ($line in $content) {
            if($line -like "*Item #*")
            {
                $isBlock = $true
            }
            if ($isBlock)
            {
                $group = $group + $line
                $current = $current + 1
            }
            if ($current -eq 9)
            {
                $isBlock = $false
                $current = 0                
                $jobs = $jobs + $group
                $group = ""
            }
        }
        Write-Host ("done")
        return $jobs
    }    
}


# Take date and time strings in the old format, combine and parse
# them to extract year, month, day, hour and minute. Returns 
# date and datetime formatted properly
function Format-Date
{
    param (
        [string]$date,
        [string]$time
    )
    Process
    {
        $dateTimeStr = $date + " " + $time
        [Hashtable] $dateTimeHash = @{}
        if($dateTimeStr -match $r)
        {    
            $dateTimeHash = $Matches
        }
        $dateTimeFormatted = ""
        if($dateTimeHash.Count){
            $dateTimeFormatted = Get-LongDate($dateTimeHash)
        }
        $properDate = Get-ShortDate($dateTimeHash)
        return @{hash=$dateTimeHash; str=$dateTimeFormatted; properdate=$properDate}
    }
}


# Add new OnDate_DateTime and LastRan_DateTime to hashtable
# that represents a job item
function Add-Datatime {    
    param (
        [Hashtable]$hash
    )
    Process
    {
        $strartedDateTime = Format-Date($hash.OnDate, $hash.Started)
        $lastRanDateTime = Format-Date($hash.LastRanOnDate, $hash.LastRan)
        $hash.OnDate = $strartedDateTime.properdate
        $hash.LastRanOnDate = $lastRanDateTime.properdate
        $hash = $hash + @{OnDate_DATETIME=$strartedDateTime.str; LastRanOnDate_DATETIME=$lastRanDateTime.str}
        return $hash
    }   
}


# Return date in "yyyy-mm-dd hh:mm:ss.fff" format
function Get-LongDate {
    param (
        [hashtable]$d
    )
    $result = ""
    if($d.Count)
    {        
        $result = $d.y + "-" + $d.m + "-" + $d.d + " " + $d.h + ":" + $d.mm + ":00.000"
    }   
    return $result
}


# Date in "yyyy-mm-dd" format
function Get-ShortDate {
    param (
        [hashtable]$d
    )
    $result = ""
    if($d.Count)
    {        
        $result = $d.y + "-" + $d.m + "-" + $d.d
    }   
    return $result
}


# Take an array of txt fragments from a source file, 
# parse them into hashtables, filter them and add new
# datetime fields
function Get-Jobs
{
    Param(
        [string[]]$source
    )
    Process
    {
        Write-Host ("Parsing jobs...")
        $jobs = @()
        $j = 0 #today jobs
        
        for ($i = 0; $i -le ($source.length - 1); $i += 1)
        {
            if($source[$i] -match $pattern_text)
            {    
                $job = $Matches
                if(($j -eq 0) -and !$currentDate)
                {
                    $currentDate = $job.OnDate
                }
                if($filterByDate -and ($job.OnDate -notmatch $currentDate))
                {
                    continue
                }   
                $job.Remove(0)
                $job = Add-Datatime($job)
                $jobs = $jobs + $job
                $j = $j + 1
            }
        }
        Write-Host ("done")
        return $jobs        
    }    
}


# Split results into chunks of $chunkSize
# Save chunks one by one into an $exportFile
function Save-Jobs {
    param (
        [hashtable[]] $jobs
    )
    Process
    {
        Write-Host("Saving results...")
        Set-Content -Path $exportFile -Value $header
        $rows = $jobs | Write-Line
        
        $chunkSize = 300  # 300 lines
        [Int32] $chunks = $rows.Count / $chunkSize
        $start = 0
        $end = 0
        for ($i = 0; $i -le ($chunks); $i += 1)
        {
            $end = ($i + 1) * $chunkSize - 1
            if($end -gt ($rows.Count - 1))
            {
                $end = $rows.Count - 1
            }
            $content = $rows[$start..$end] -join "`n"
            Add-Content -Path $exportFile -Value $content
            $start = $start + $chunkSize
        }
        Write-Host ("done")        
    }
}


function ParseDate {
    param (
        [string]$date
    )
    Process
    {
        $reg_ex = "((\d{4})[/-]+)?(\d{1,2})[/-]+(\d{1,2})([/-]+(\d{4}))?"
        if($date -match $reg_ex)
        {
            $dh = $Matches
            if($dh.2){
                return "" + $dh.4 + "-" + $dh.3 + "-" + $dh.2
            }
            if ($dh.6){
                return "" + $dh.3 + "-" + $dh.4 + "-" + $dh.6
            }
        }
        Write-Host "Acceptable date formats are YYYY-MM-DD and DD-MM-YYYY" -ForegroundColor red -BackgroundColor white
        return ""
        
    }
    
}


# Regular expression to parse jobs
$pattern_text = "^\s+Item #: (?<ItemNr>\d+)" +
                "\s+Job name: (?<JobName>\S*)" +
                "\s+Status: (?<Status>\D+\S+)" +
                "\s+Started: (?<Started>\d+:\d+)?" +
                "\s+On date: (?<OnDate>\d+-\d+-\d+)?" +
                "\s+Last ran: (?<LastRan>\d+:\d+)?" +
                "\s+On date: (?<LastRanOnDate>\d+-\d+-\d+)?" +
                "\s+Elapsed time: (?<ElapsedTime>\d+:\d+:\d+)?" +
                "\s+Description: (?<Description>.*)$"
# Regular expression to parse Date and Time
$r = "(?<d>\d+)-(?<m>\d+)-(?<y>\d+) (?<h>\d+):(?<mm>\d+)"
# Head of the export file
$header = "File,ItemNr,JobName,Status,Started,OnDate,OnDate_DATETIME,LastRan,LastRanOnDate,LastRanOnDate_DATETIME,ElapsedTime,Description"



#[string] $currentDate = Get-Date -Format "dd-MM-yyyy"
$currentDate = ""

for ($i = 0; $i -le ($args.length - 1); $i += 1)
{
    switch -Exact ($args[$i]) {
        "split" {
            $combineFiles = $false
            Break
        }
        'all' {
            $filterByDate = $false
            Break
        }
        "ondate" {
            try {
                $ondate = $args[$i + 1]
                $currentDate = ParseDate($ondate)
            }
            catch {                
            }
            Break
        }
    }
}


do 
{
    $files = Get-ChildItem -Path $workDir -Depth 1 -Name -Include *.txt
    if(-not $files.Count)
    {
        Write-Host ("No .txt files found.")
        $workDir = Read-Host "Please, point to C:\Path\To\File"
        #throw "No .txt files found"
    }
} while (-not $files.Count)


if($combineFiles)
{
    $jobs = @()
    foreach($file in $files)
    {
        $data = Read-File($file)
        $jobs_1 = Get-Jobs($data)
        $jobs_1 | ForEach-Object -Process {$_["file"]=$file}
        $jobs = $jobs + $jobs_1
    }
    if($jobs.Count)
    {
        Save-Jobs($jobs)
    }
    else {
        Write-Host ("No jobs found")
    }
}
else
{
    $i = 0
    $exportChunks = $exportFile -split "\."
    $preservedPart = $exportChunks[-2]
    foreach($file in $files)
    {
        $appendix = "_" + $i
        if($file -match "\w+(_\w+)\.txt")
        {
            $appendix = $Matches[1]
            $appendix = $appendix
        }
        $exportChunks[-2] = $preservedPart + $appendix
        $exportFile = $exportChunks -join "."
        $data = Read-File($file)
        $jobs = Get-Jobs($data)
        if($jobs.Count)
        {
            Save-Jobs($jobs)
        }
        else {
            Write-Host ("No jobs found")
        }
        $i = $i + 1
    }
}


