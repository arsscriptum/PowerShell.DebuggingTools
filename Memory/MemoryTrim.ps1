<#
#̷𝓍    𝓐𝓡𝓢 𝓢𝓒𝓡𝓘𝓟𝓣𝓤𝓜
#̷𝓍    Platform Invoke (P/Invoke) for 🇵​​​​​🇴​​​​​🇼​​​​​🇪​​​​​🇷​​​​​🇸​​​​​🇭​​​​​🇪​​​​​🇱​​​​​🇱​​​​​ 
#̷𝓍    🇧​​​​​🇾​​​​​ 🇬​​​​​🇺​​​​​🇮​​​​​🇱​​​​​🇱​​​​​🇦​​​​​🇺​​​​​🇲​​​​​🇪​​​​​🇵​​​​​🇱​​​​​🇦​​​​​🇳​​​​​🇹​​​​​🇪​​​​​.🇶​​​​​🇨​​​​​@🇬​​​​​🇲​​​​​🇦​​​​​🇮​​​​​🇱​​​​​.🇨​​​​​🇴​​​​​🇲​​​​​
#>


function Convert-Bytes {
    # Converts input number (bytes) into human readable format, up to Petabytes.
    [Alias('2b', 'convertbytes')]
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true,Position=0)]
        [float]$Size,
        [ValidateSet("Auto", "KB", "MB", "GB", "TB")]
        [string]$Format = "Auto",
        [ValidateRange(0, 8)]
        [int]$Decimals = 2,        
        [Parameter(Mandatory=$false, HelpMessage="Value only, no format")]
        [switch]$ValueOnly,
        [Parameter(Mandatory=$false, HelpMessage="Value in Kb")]
        [switch]$KB,
        [Parameter(Mandatory=$false, HelpMessage="Value in MegaBytes")]
        [switch]$MB,
        [Parameter(Mandatory=$false, HelpMessage="Value in GigaBytes")]
        [switch]$GB,
        [Parameter(Mandatory=$false, HelpMessage="Value in TeraBytes")]
        [switch]$TB
    )
    Begin{
        $sizes = 'KB','MB','GB','TB','PB'
    }
    Process {

    $FormatStr = ''

        if($KB){
            $Format = "KB"
        }elseif($MB){
            $Format = "MB"
        }elseif($GB){
            $Format = "GB"    
        }elseif($TB){
            $Format = "TB"
        }

        if($PSBoundParameters.ContainsKey('ValueOnly') -eq $False){
            $FormatStr = " $Format"
        }
        if(($Null -eq $Size) -Or ($Size -eq 0)){
            return "0 Bytes"
        }
        switch ($format) {
            "KB" { "{0:n$Decimals}$FormatStr" -f ($Size / 1KB) }
            "MB" { "{0:n$Decimals}$FormatStr" -f ($Size / 1MB) }
            "GB" { "{0:n$Decimals}$FormatStr" -f ($Size / 1GB) }
            "TB" { "{0:n$Decimals}$FormatStr" -f ($Size / 1TB) }
            "Auto" {
                # New for loop
                for($x = 0;$x -lt $sizes.count; $x++){
                    if ($Size -lt [int64]"1$($sizes[$x])"){
                        if ($x -eq 0){
                            return "$Size B"
                        } else {
                            $num = $Size / [int64]"1$($sizes[$x-1])"
                            $num = "{0:N2}" -f $num
                            return "$num $($sizes[$x-1])"
                        }
                    }
                }               
            }
        }
    }
}

function Register-MemoryTools{
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $false)]
        [Switch]$Force
    )

    $CsSource = (Join-Path $PSScriptRoot "MemoryTools.cs")  
    
    if (!("MemoryTools.Win32" -as [type])) {
        Write-Verbose "Registering $CsSource... " 
        Add-Type -Path "$CsSource"
    }else{
        Write-Verbose "MemoryTools.Win32 already registered: $CsSource... " 
    }
}


function Get-AvailableMBytes {
    $cname = (Get-Counter -ListSet Memory).Paths[28]

    [long]$availableMemory = (Get-Counter $cname).CounterSamples.CookedValue * 1MB
    return $availableMemory
}

function Get-MemoryUsageStat{
    [CmdletBinding(SupportsShouldProcess)]
    param ()
        $last = Get-Variable -Name "last_memory_usage_byte" -ValueOnly
        $memusagebyte = [System.GC]::GetTotalMemory('forcefullcollection')
        $memusageMB = $memusagebyte / 1MB
        $diffbytes = $memusagebyte - $last
        $difftext = ''
        $sign = ''
        $colordiff = 'Green'
        if ( $script:last_memory_usage_byte -ne 0 ){
            if ( $diffbytes -ge 0 )
            {
              $sign = '+'
              $colordiff = 'Red'
            }
            $difftext = "Diff $sign$diffbytes"
        }
        $StrData = 'Memory usage: {0:n1} MB ({1:n0} Bytes). ' -f  $memusageMB, $memusagebyte
        Write-Host $StrData -n -f DarkCyan
        Write-Host $difftext -f $colordiff
      

        # save last value in script global variable
        Set-Variable -Name "last_memory_usage_byte" -Value $memusagebyte -Option AllScope -Scope Global -Visibility Public -Force -ErrorAction Ignore
}
 
function Invoke-TrimProcessMemory {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory = $True, Position =0)]
        [ValidateNotNullOrEmpty()]$ProcessName,
        [Parameter(Mandatory = $true)]
        [int]$MaxWorkingSet,
        [Parameter(Mandatory = $false)]        
        [switch]$HardMax,
        [Parameter(Mandatory = $false)]
        [switch]$HardMin,
        [Parameter(Mandatory = $false)]
        [switch]$ReportStat
    )

    Register-MemoryTools -Force


    [long]$saved = 0
    [int]$trimmed = 0
    [int]$adjusted = 0
    [int]$flags = 0
    [int]$minWorkingSet = -1
    if( $minWorkingSet -gt 0 )
    {
        if( $HardMin )
        {
            $flags = $flags -bor 1
        }
        else  ## soft
        {
            $flags = $flags -bor 2
        }
    }

    if( $MaxWorkingSet -gt 0 )
    {
        if( $HardMax )
        {
            $flags = $flags -bor 4
        }
        else  ## soft
        {
            $flags = $flags -bor 8
        }
        if( $minWorkingSet -le 0 )
        {
            $minWorkingSet = 1 ## if a maximum is specified then we must specify a minimum too - this will default to the minimum
        }
    }

    $ErrorActionPreference = 'Ignore'
    try {

        [long]$totalMemory = ( Get-CimInstance -Class Win32_ComputerSystem -Property TotalPhysicalMemory ).TotalPhysicalMemory
        [system.collections.arraylist]$reports = [system.collections.arraylist]::new()
        
        [long]$availableMemory = Get-AvailableMBytes
        if($ReportStat){
            [int]$left = ( $availableMemory / $totalMemory ) * 100
            $logstr =  "[Before] Available memory is {0}MB out of {1}MB total ({2}%)" -f ( $availableMemory / 1MB ) , [math]::Floor( $totalMemory / 1MB ) , [math]::Round( $left )
            Write-Host $logstr -f DarkYellow
            Get-MemoryUsageStat
        }

        [array]$processes = Get-Process "$ProcessName"

        ForEach($process in $processes){
            [int]$thisMinimumWorkingSet = -1 
            [int]$thisMaximumWorkingSet = -1 
            [int]$thisFlags = -1 ## Grammar alert! :-)
            ## https://msdn.microsoft.com/en-us/library/windows/desktop/ms683227(v=vs.85).aspx
            [bool]$result = [MemoryTools.Win32.Memory]::GetProcessWorkingSetSizeEx( $p.Handle, [ref]$thisMinimumWorkingSet,[ref]$thisMaximumWorkingSet,[ref]$thisFlags);
            $LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
            if( $result ){
                ## convert flags value - if not hard then will be soft so no point reporting that separately IMHO
                [bool]$HardMinimumWorkingSet = $thisFlags -band 1 ## QUOTA_LIMITS_HARDWS_MIN_ENABLE
                [bool]$HardMaximumWorkingSet = $thisFlags -band 4 ## QUOTA_LIMITS_HARDWS_MAX_ENABLE
                [pscustomobject]$obj = [pscustomobject][ordered]@{ 'Name' = $process.Name ; 'PID' = $process.Id ; 'Handle Count' = $process.HandleCount ; 'Start Time' = $process.StartTime ;
                                'Hard Minimum Working Set Limit' = $HardMinimumWorkingSet ; 'Hard Maximum Working Set Limit' = $HardMaximumWorkingSet ;
                                'Working Set (MB)' = $process.WorkingSet64 / 1MB ;'Peak Working Set (MB)' = $process.PeakWorkingSet64 / 1MB ;
                                'Commit Size (MB)' = $process.PagedMemorySize / 1MB; 
                                'Paged Pool Memory Size (KB)' = $process.PagedSystemMemorySize64 / 1KB ; 'Non-paged Pool Memory Size (KB)' = $process.NonpagedSystemMemorySize64 / 1KB ;
                                'Minimum Working Set (KB)' = $thisMinimumWorkingSet / 1KB ; 'Maximum Working Set (KB)' = $thisMaximumWorkingSet / 1KB
                                'Hard Minimum Working Set' = $HardMinimumWorkingSet ; 'Hard Maximum Working Set' = $HardMaximumWorkingSet 
                                'Virtual Memory Size (GB)' = $process.VirtualMemorySize64 / 1GB; 'Peak Virtual Memory Size (GB)' = $process.PeakVirtualMemorySize64 / 1GB; }
                
                [void]$reports.Add($obj)
            }else{                   
                Write-Warning ( "Failed to get working set info for {0} pid {1} - {2} " -f $process.Name , $process.Id , $LastError)
            }


            [bool]$result = [MemoryTools.Win32.Memory]::SetProcessWorkingSetSizeEx( $process.Handle,$minWorkingSet,$MaxWorkingSet,$flags);$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
            if( ! $result ){
                $errstr = "Failed to trim {0} pid {1} - {2} " -f $process.Name , $process.Id , $LastError
                throw $errstr
            }
            $now = Get-Process -Id $process.Id -ErrorAction SilentlyContinue
            if( $now ){
                $saved += $process.WS - $now.WS
                $trimmed++
            }
            $reports |  ConvertTo-Json | Set-Content "$PSScriptRoot\Report.json"
        }

        if($ReportStat){
            [long]$availableMemoryAfter = Get-AvailableMBytes
            $a = $availableMemory / 1MB 
            $b = $availableMemoryAfter  / 1MB 
            [int]$left = ( $availableMemoryAfter / $totalMemory ) * 100
            $trimmedstr = "Trimmed {0}MB from {1} processes giving {2}MB extra available" -f [math]::Round( $saved / 1MB , 1 ) , $trimmed , ( $b - $a)
            $logstr =     "[After] Available memory is {0}MB out of {1}MB total ({2}%)" -f ( $availableMemoryAfter / 1MB ) , [math]::Floor( $totalMemory / 1MB ) , [math]::Round( $left )
            Write-Host $trimmedstr -f DarkYellow
            Write-Host $logstr -f DarkGreen
            Get-MemoryUsageStat
        }
    }
    catch {
        Write-Error "$_"
    }
}


function Get-ProcessMemoryUsageDetails{
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory = $True)]
        [string]$ProcessName   
    )
    [bool]$ShowCmdline=$True

    $ErrorActionPreference = 'Ignore'
    try {
        [array]$processes = Get-Process "$ProcessName"

        ForEach($process in $processes){
            $process | Select-Object @{Name="ProcessId";Expression={$_.Id}}, @{Name="Mem Usage(MB)";Expression={[math]::round($_.ws / 1mb)}},@{Name="WorkingSet (Kb)";Expression={[math]::round($_.ws / 1kb)}}
        }
    }catch {
        Write-Error "$_"
    }
}

function Get-ProcessMemoryUsage
{
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]$ProcessName
    )
    $ErrorActionPreference = 'Ignore'
    try {
        [array]$prcs = Get-Process "$ProcessName"

        if($prcs -eq $Null) { throw "No Such Process" }
        #"Write-Host "===============================================================================" -f DarkRed
        #Write-Host "MEMORY USAGE FOR $ProcessName" -f DarkYellow;
        #$Data = $Process | Group-Object -Property ProcessName | Format-Table Name, Count, @{n='Mem (KB)';e={'{0:N0}' -f (($_.Group|Measure-Object WorkingSet -Sum).Sum / 1KB)};a='right'} -AutoSize
        $NumProcess =$prcs.Length
        $MemoryBytes = $prcs | Group-Object -Property ProcessName | % {  (($_.Group|Measure-Object WorkingSet -Sum).Sum) }

        $StrMem = Convert-Bytes $MemoryBytes
        [pscustomobject]$ret = [PSCustomObject]@{
            Name = $ProcessName
            Count = $NumProcess
            Memory = $StrMem
        }
        return $ret
    }
    catch {
        Write-Host '[ProcessMemoryUsage] ' -n -f DarkRed
        Write-Host "$_" -f DarkYellow
    }
}
 