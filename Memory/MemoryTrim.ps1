<#
#̷𝓍    𝓐𝓡𝓢 𝓢𝓒𝓡𝓘𝓟𝓣𝓤𝓜
#̷𝓍    Platform Invoke (P/Invoke) for 🇵​​​​​🇴​​​​​🇼​​​​​🇪​​​​​🇷​​​​​🇸​​​​​🇭​​​​​🇪​​​​​🇱​​​​​🇱​​​​​ 
#̷𝓍    🇧​​​​​🇾​​​​​ 🇬​​​​​🇺​​​​​🇮​​​​​🇱​​​​​🇱​​​​​🇦​​​​​🇺​​​​​🇲​​​​​🇪​​​​​🇵​​​​​🇱​​​​​🇦​​​​​🇳​​​​​🇹​​​​​🇪​​​​​.🇶​​​​​🇨​​​​​@🇬​​​​​🇲​​​​​🇦​​​​​🇮​​​​​🇱​​​​​.🇨​​​​​🇴​​​​​🇲​​​​​
#>



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

 
function Invoke-TrimProcessMemory {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory = $True, Position =0)]
        [ValidateNotNullOrEmpty()]$ProcessName,
        [Parameter(Mandatory = $true)]
        [int]$MaxWorkingSet,
        [Parameter(Mandatory = $false)]        
        [switch]$hardMax ,
        [Parameter(Mandatory = $false)]
        [switch]$hardMin
    )

    Register-MemoryTools -Verbose -Force


    [long]$saved = 0
    [int]$trimmed = 0
    [int]$adjusted = 0
    [int]$flags = 0
    [int]$minWorkingSet = -1
    if( $minWorkingSet -gt 0 )
    {
        if( $hardMin )
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
        if( $hardMax )
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
        [int]$left = ( $availableMemory / $totalMemory ) * 100
        $logstr =  "[Before] Available memory is {0}MB out of {1}MB total ({2}%)" -f ( $availableMemory / 1MB ) , [math]::Floor( $totalMemory / 1MB ) , [math]::Round( $left )
        Write-Host $logstr -f DarkYellow

        [array]$processes = Get-Process "$ProcessName"

        ForEach($process in $processes){
            

            #$process | Select-Object Name, @{Name="ProcessId";Expression={$_.Id}}, @{Name="Mem Usage(MB)";Expression={[math]::round($_.ws / 1mb)}}, @{Name="WorkingSet (Kb)";Expression={[math]::round($_.WorkingSet / 1kb)}}
            [int]$thisMinimumWorkingSet = -1 
            [int]$thisMaximumWorkingSet = -1 
            [int]$thisFlags = -1 ## Grammar alert! :-)
            ## https://msdn.microsoft.com/en-us/library/windows/desktop/ms683227(v=vs.85).aspx
            [bool]$result = [MemoryTools.Win32.Memory]::GetProcessWorkingSetSizeEx( $p.Handle, [ref]$thisMinimumWorkingSet,[ref]$thisMaximumWorkingSet,[ref]$thisFlags);
            $LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
            if( $result ){
                ## convert flags value - if not hard then will be soft so no point reporting that separately IMHO
                [bool]$hardMinimumWorkingSet = $thisFlags -band 1 ## QUOTA_LIMITS_HARDWS_MIN_ENABLE
                [bool]$hardMaximumWorkingSet = $thisFlags -band 4 ## QUOTA_LIMITS_HARDWS_MAX_ENABLE
                [pscustomobject]$obj = [pscustomobject][ordered]@{ 'Name' = $process.Name ; 'PID' = $process.Id ; 'Handle Count' = $process.HandleCount ; 'Start Time' = $process.StartTime ;
                                'Hard Minimum Working Set Limit' = $hardMinimumWorkingSet ; 'Hard Maximum Working Set Limit' = $hardMaximumWorkingSet ;
                                'Working Set (MB)' = $process.WorkingSet64 / 1MB ;'Peak Working Set (MB)' = $process.PeakWorkingSet64 / 1MB ;
                                'Commit Size (MB)' = $process.PagedMemorySize / 1MB; 
                                'Paged Pool Memory Size (KB)' = $process.PagedSystemMemorySize64 / 1KB ; 'Non-paged Pool Memory Size (KB)' = $process.NonpagedSystemMemorySize64 / 1KB ;
                                'Minimum Working Set (KB)' = $thisMinimumWorkingSet / 1KB ; 'Maximum Working Set (KB)' = $thisMaximumWorkingSet / 1KB
                                'Hard Minimum Working Set' = $hardMinimumWorkingSet ; 'Hard Maximum Working Set' = $hardMaximumWorkingSet 
                                'Virtual Memory Size (GB)' = $process.VirtualMemorySize64 / 1GB; 'Peak Virtual Memory Size (GB)' = $process.PeakVirtualMemorySize64 / 1GB; }
                
                [void]$reports.Add($obj)
            }else{                   
                Write-Warning ( "Failed to get working set info for {0} pid {1} - {2} " -f $process.Name , $process.Id , $LastError)
            }


            $minWorkingSet = 1

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

        if( $True ){
            [long]$availableMemoryAfter = Get-AvailableMBytes
            $a = $availableMemory / 1MB 
            $b = $availableMemoryAfter  / 1MB 
            [int]$left = ( $availableMemoryAfter / $totalMemory ) * 100
            $trimmedstr = "Trimmed {0}MB from {1} processes giving {2}MB extra available" -f [math]::Round( $saved / 1MB , 1 ) , $trimmed , ( $b - $a)
            $logstr =     "[After] Available memory is {0}MB out of {1}MB total ({2}%)" -f ( $availableMemoryAfter / 1MB ) , [math]::Floor( $totalMemory / 1MB ) , [math]::Round( $left )
            Write-Host $trimmedstr -f DarkYellow
            Write-Host $logstr -f DarkGreen
        }
    }
    catch {
        Write-Error "$_"
    }
}

