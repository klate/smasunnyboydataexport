# Path to Sunny Explorer EXE
$sunnyExplorerExe = ""
# Path to plant configuration
$plantFile = ""
# global export path
$globalexportPath = ""
# Password to access inverter
$userPassword = ""

"starting sma export"

# number formatting func
function Get-TwoDigitNumString ([int] $num){
    if ($num -lt 10) {
        return "0$num"
    }
    return "$num"
}

# folder creation
function New-NonExistingDirs ([string[]] $dirs){
    foreach ($dir in $dirs) {
        if (![System.IO.Directory]::Exists($dir))
        {
            [System.IO.Directory]::CreateDirectory($dir)
        }
    }   
}

# sma export process 
function Get-SmaData ([string[]] $processargs){
    foreach ($parg in $processargs) {       
        $process = New-Object System.Diagnostics.Process
        "calling sma with $parg"
        try
        {
            $process.StartInfo.Filename = $sunnyExplorerExe
            $process.StartInfo.Arguments = $parg
            $process.StartInfo.UseShellExecute = $false
            $process.Start()
            $process.WaitForExit()
        }
        finally
        {
            $process.Dispose()
        }
    }
}

$now = Get-Date
$lastMonth = $now.AddMonths(-1)
[string] $month = Get-TwoDigitNumString($lastMonth.Month)

# export directory dirs
$baseMonthExportPath = Join-Path $globalexportPath $lastMonth.Year
$baseMonthExportPath = Join-Path $baseMonthExportPath $month
$daily5minexportPath = Join-Path $baseMonthExportPath "daily5minInterval"

# create missing dirs
New-NonExistingDirs $baseMonthExportPath, $daily5minexportPath

# sma date params
$beginningOfLastMonth = $lastMonth.ToString("yyyyMM") + "01"
$daysInLastMonth = [DateTime]::DaysInMonth($lastMonth.Year, $lastMonth.Month)
$endOfLastMonth = $lastMonth.ToString("yyyyMM") + $daysInLastMonth.ToString()

# get the data
Get-SmaData (
    # 5-minute interval data
    [string]::Format("`"{0}`" -userlevel user -password {1} -exportdir `"{2}`" -exportrange {3}-{4} -export energy5min", $plantFile, $userPassword, $daily5minexportPath, $beginningOfLastMonth, $endOfLastMonth),
    # Daily data
    [string]::Format("`"{0}`" -userlevel user -password {1} -exportdir `"{2}`" -exportrange {3}-{4} -export energydaily", $plantFile, $userPassword, $baseMonthExportPath, $beginningOfLastMonth, $endOfLastMonth),
    #Events 
    [string]::Format("`"{0}`" -userlevel user -password {1} -exportdir `"{2}`" -exportrange {3}-{4} -export events", $plantFile, $userPassword, $baseMonthExportPath, $beginningOfLastMonth, $endOfLastMonth)
)
"finished sma export"
