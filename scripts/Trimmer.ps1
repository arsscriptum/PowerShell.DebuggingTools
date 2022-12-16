#requires -version 3

<#
    Trim working sets of processes or set maximum working set sizes

    Guy Leech, 2018

    Modification History

    10/03/18  GL  Optimised code

    12/03/18  GL  Added reporting

    16/03/18  GL Fixed bug where -available was being passed as False when not specified
                 Workaround for invocation external to PowerShell for arrays being flattened
                 Made process ids parameter an array
                 Added process id and name filter to Get-Process cmdlet call for efficiency
                 Added include and exclude options for user names

    30/03/18  GL Added ability to wait for specific list of processes to start before continuing.
                 Added looping

    23/04/18  GL Added exiting from loop if pids specified no longer exist
                 Added -forceIt as equivalent to -confirm:$false for use via scheduled tasks

    12/02/22  GL Added output of whether hard workin set limits and -nogridview
#>

<#
.SYNOPSIS

Manipulate the working sets (memory usage) of processes or report their current memory usage and working set limits and types (hard or soft)

.DESCRIPTION

Can reduce the memory footprints of running processes to make more memory available and stop processes that leak memory from leaking

.PARAMETER Processes

A comma separated list of process names to use (without the .exe extension). By default all processes will be trimmed if the script has access to them.

.PARAMETER IncludeUsers

A comma separated of qualified names of process owners to include. Must be run as an admin for this to work. Specify domain or other qualifier, e.g. "NT AUTHORITY\SYSTEM'

.PARAMETER ExcludeUsers

A comma separated of qualified names of process owners to exclude. Must be run as an admin for this to work. Specify domain or other qualifier, e.g. "NT AUTHORITY\NETWORK SERVICE' or DOMAIN\Chris.Harvey

.PARAMETER Exclude

A comma separated list of process names to ignore (without the .exe extension).

.PARAMETER Above

Only trim the working set if the process' working set is currently above this value. Qualify with MB or GB as required. Default is to trim all processes

.PARAMETER WaitFor

A comma separated list of processes to wait for unless -alreadyStarted is specified and one of the processes is already running unless it is not in the current session and -thisSession specified

.PARAMETER AlreadyStarted

Used when -WaitFor specified such that waiting will not occur if any of the processes specified via -WaitFor are already running although only in the current session if -thisSession is specified

.PARAMETER PollPeriod

The time in seconds between checks for new processes that match the -WaitFor process list.

.PARAMETER MinWorkingSet

Set the minimum working set size to this value. Qualify with MB or GB as required. Default is to not set a minimum value.

.PARAMETER MaxWorkingSet

Set the maximum working set size to this value. Qualify with MB or GB as required. Default is to not set a maximum value.

.PARAMETER HardMin

When MinWorkingSet is specified, the limit will be enforced so the working set is never allowed to be less that the value. Default is a soft limit which is not enforced.

.PARAMETER HardMax

When MaxWorkingSet is specified, the limit will be enforced so the working set is never allowed to exceed the value. Default is a soft limit which can be exceeded.

.PARAMETER Loop

Loop infinitely

.PARAMETER forceIt

DO not prompt for confirmation before adjusting CPU priority

.PARAMETER Report

Produce a report of the current working set usage and limit types for processes in the selection. Will output to a grid view unless -outputFile is specified.

.PARAMETER OutputFile

Ue with -report to write the results to a csv format file. If the file already exists the operation will fail.

.PARAMETER ProcessIds

Only trim the specific process ids

.PARAMETER ThisSession

Will only trim working sets of processes in the same session as the sript is running in. The default is to trim in all sessions.

.PARAMETER SessionIds

Only trim processes running in the specified sessions which is a comma separated list of session ids. The default is to trim in all sessions.

.PARAMETER NotSessionId

Only trim processes not running in the specified sessions which is a comma separated list of session ids. The default is to trim in all sessions.

.PARAMETER Available

Specify as a percentage or an absolute value. Will only trim if the available memory is below the parameter specified. The default is to always trim.

.PARAMETER Savings

This will show a summary of the trimming at the end of processing. Note that working sets can grow once trimmed so the amount trimmed may be higher than the actual increase in available memory.

.PARAMETER Disconnected

This will only trim memory in sessions which are disconnected. The default is to target all sessions.

.PARAMETER Idle

If no user input has been received in the last x seconds, whre x is the parameter passed, then the session is considered idle and processes will be trimmed.

.PARAMETER nogridview

Put the results (use -report) onto the pipeline, not in a grid view

.PARAMETER Background

Only trim processes which are not the process responsible for the foreground window. Implies -ThisSession since cannot check windows in other sessions.

.PARAMETER Install

Create two scheduled tasks, one which will trim that user's session on disconnect or screen lock and the other runs at the frequency specified in seconds divided by two, checks if the user is idle for the specified number of seconds and trims if they are.
So if a parameter of 600 is passed, the task will run every 5 minutes and if the user has made no mouse or keyboard input for 10 minutes then their processes are trimmed.

.PARAMETER Uninstall

Removes the two scheduled tasks previously created for the user running the script.

.EXAMPLE

& .\Trimmer.ps1

This will trim all processes in all sessions to which the account running the script has access.

.EXAMPLE

& .\Trimmer.ps1 -ThisSession -Above 50MB

Only trim processes in the same session as the script and whose working set exceeds 50MB.

.EXAMPLE

& .\Trimmer.ps1 -MaxWorkingSet 100MB -HardMax -Above 50MB -Processes LeakyApp

Only trim processes called LeakyApp in any session whose working set exceeds 50MB and set the maximum working set size to 100MB which cannot be exceeded.
Will only apply to instances of LeakyApp which have already started, instances started after the script is run will not be subject to the restriction. 

.EXAMPLE

& .\Trimmer.ps1 -MaxWorkingSet 10MB -Processes Chrome

Trim Chrome processes to 10MB, rather than completely emptying their working set. If processes rapidly regain working sets after being trimmed, this can
cause page file thrashing so reducing the working set but not completely emptying them can still save memory but reduce the risk of page file thrashing.
Picking the figure to use for the working set is trial and error but typically one would use the value that it settles to a few minutes after trimming.

.EXAMPLE

& .\Trimmer.ps1 -Install 600 -Logoff

Create two scheduled tasks for this user which only run when the user is logged on. The first runs at session lock or disconnect and trims all processes in that session.
The second task runs every 300 seconds and if no mouse or keyboard input has been received in the last 600 seconds then all background processes in that session will be trimmed.
At logoff, the scheduled tasks will be removed.

.EXAMPLE

& .\Trimmer.ps1 Uninstall

Delete the two scheduled tasks for this user

.NOTES

If you trim too much and/or too frequently, you run the risk of reducing performance by overusing the page file.

Supports the "-whatif" parameter so you can see what processes it will trim without actually performing the trim.

If emptying the working set does cause too much paging, try using the -MaxWorkingSet parameter to apply a soft limit which will cause the process
to be trimmed down to that value but it can then grow larger if required.

Uses Windows API SetProcessWorkingSetSizeEx() - https://msdn.microsoft.com/en-us/library/windows/desktop/ms686237(v=vs.85).aspx
#>

[cmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')]

Param
(
    [string]$logFile ,
    [int]$install ,
    [int]$idle ,  ## seconds
    [switch]$uninstall ,
    [string[]]$processes ,
    [string[]]$exclude , 
    [string[]]$includeUsers ,
    [string[]]$excludeUsers ,
    [string[]]$waitFor ,
    [switch]$alreadyStarted ,
    [switch]$report ,
    [switch]$nogridview ,
    [string]$outputFile ,
    [int]$above = 10MB ,
    [int]$minWorkingSet = -1 ,
    [int]$maxWorkingSet = -1 ,
    [int]$pollPeriod = 5 ,
    [switch]$hardMax ,
    [switch]$hardMin ,
    [switch]$newOnly ,
    [switch]$thisSession ,
    [string[]]$processIds ,
    [string[]]$sessionIds ,
    [string[]]$notSessionIds ,
    [string]$available ,
    [switch]$loop ,
    [switch]$savings ,
    [switch]$disconnected ,
    [switch]$background ,
    [switch]$scheduled ,
    [switch]$logoff ,
    [switch]$forceIt ,
    [string]$taskFolder = '\MemoryTrimming' 
)

[int]$minimumIdlePeriod = 120 ## where minimum reptition of a scheduled task must be at least 1 minute, thus idle time must be at least double that (https://msdn.microsoft.com/en-us/library/windows/desktop/aa382993(v=vs.85).aspx)

## Borrowed from http://stackoverflow.com/a/15846912 and adapted
Add-Type @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PInvoke.Win32
{
  
    public static class Memory
    {
        [DllImport("kernel32.dll", SetLastError=true)]
        public static extern bool SetProcessWorkingSetSizeEx( IntPtr proc, int min, int max , int flags );
        [DllImport("kernel32.dll", SetLastError=true)]
        public static extern bool GetProcessWorkingSetSizeEx( IntPtr hProcess, ref int min, ref int max , ref int flags );
    }
    public static class UserInput
    {  
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        
        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
 
        [DllImport("user32.dll", SetLastError=false)]
        private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        [StructLayout(LayoutKind.Sequential)]
        private struct LASTINPUTINFO
        {
            public uint cbSize;
            public int dwTime;
        }
        public static DateTime LastInput
        {
            get
            {
                DateTime bootTime = DateTime.UtcNow.AddMilliseconds(-Environment.TickCount);
                DateTime lastInput = bootTime.AddMilliseconds(LastInputTicks);
                return lastInput;
            }
        }
        public static TimeSpan IdleTime
        {
            get
            {
                return DateTime.UtcNow.Subtract(LastInput);
            }
        }
        public static int LastInputTicks
        {
            get
            {
                LASTINPUTINFO lii = new LASTINPUTINFO();
                lii.cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO));
                GetLastInputInfo(ref lii);
                return lii.dwTime;
            }
        }
    }
}
'@

[datetime]$monitoringStartTime = Get-Date
[int]$thisSessionId = (Get-Process -Id $pid).SessionId

## workaround for scheduled task not liking -confirm:$false being passed
if( $forceIt )
{
     $ConfirmPreference = 'None'
}

do
{
    if( $waitFor -and $waitFor.Count )
    {
        [bool]$found = $false
        $thisProcess = $null

        [datetime]$startedAfter = Get-Date
        if( $alreadyStarted ) ## we are not waiting for new instances so existing ones qualify too
        {
            $startedAfter = Get-Date -Date '01/01/1970' ## saves having to grab LastBootupTime
        }

        while( ! $found )
        {
            Write-Verbose "$(Get-Date): waiting for one of $($waitFor -join ',') to launch (only in session $thisSessionId is $thisSession)"
            ## wait for one of a set of specific processes to start - useful when you need to apply a hard working set limit to a known leaky process
            Get-Process -Name $waitFor -ErrorAction SilentlyContinue | Where-Object { $_.StartTime -gt $startedAfter } | ForEach-Object `
            {
                if( ! $found )
                {
                    $thisProcess = $_
                    ## we don't support all filtering options here
                    if( $thisSession )
                    {
                        $found = ( $thisSessionId -eq $thisProcess.SessionId )
                    }
                    else
                    {
                        $found = $true
                    }
                }
            }
            if( ! $found )
            {
                Start-Sleep -Seconds $pollPeriod
            }
        }
        Write-Verbose "$(Get-Date) : process $($thisProcess.Name) id $($thisProcess.Id) started at $($thisProcess.StartTime)"
    }
    
    if( $idle -gt 0 )
    {
        $idleTime = [PInvoke.Win32.UserInput]::IdleTime.TotalSeconds
        Write-Verbose "Idle time is $idleTime seconds"
        if( $idleTime -lt $idle )
        {
            Write-Verbose "Idle time is only $idleTime seconds, less than $idle"
            if( ! $scheduled -or ( $scheduled -and ! $background ) )
            {
                if( ! [string]::IsNullOrEmpty( $logFile ) )
                {
                    Stop-Transcript
                }
                return
            }
            else
            {
                Write-Verbose "Not idle but we are a scheduled task and trimming background processes so continue"
            }
        }
    }

    [long]$ActiveHandle = $null
    $activePid = [IntPtr]::Zero

    if( $background )
    {
        [long]$ActiveHandle = [PInvoke.Win32.UserInput]::GetForeGroundWindow( )
        if( ! $ActiveHandle )
        {
            Write-Error "Unable to find foreground window"
            return 1
        }
        else
        {
            $activeThreadId = [PInvoke.Win32.UserInput]::GetWindowThreadProcessId( $ActiveHandle , [ref] $activePid )
            if( $activePid -ne [IntPtr]::Zero )
            {
                Write-Verbose ( "Foreground window is pid {0} {1}" -f $activePid , (Get-Process -Id $activePid).Name )
            }
            else
            {
                Write-Error "Unable to get handle on process for foreground window $ActiveHandle"
                return 1
            }
        }
        $thisSession = $true ## can only check windows in this session
    }

    [int]$flags = 0
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

    if( $maxWorkingSet -gt 0 )
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

    [long]$availableMemory = (Get-Counter '\Memory\Available MBytes').CounterSamples[0].CookedValue * 1MB

    if( ! [string]::IsNullOrEmpty( $available ) )
    {
        ## Need to find out memory available and total
        [long]$totalMemory = ( Get-CimInstance -Class Win32_ComputerSystem -Property TotalPhysicalMemory ).TotalPhysicalMemory
        [int]$left = ( $availableMemory / $totalMemory ) * 100
        Write-Verbose ( "Available memory is {0}MB out of {1}MB total ({2}%)" -f ( $availableMemory / 1MB ) , [math]::Floor( $totalMemory / 1MB ) , [math]::Round( $left ) )

        [bool]$proceed = $false
        ## See if we are dealing with absolute or percentage
        if( $available[-1] -eq '%' )
        {
            [int]$percentage = $available -replace '%$'
            $proceed = $left -lt $percentage 
        }
        else ## absolute
        {
            [long]$threshold = Invoke-Expression $available
            $proceed = $availableMemory -lt $threshold
        }

        if( ! $proceed )
        {
            Write-Verbose "Not trimming as memory available is above specified threshold"
            if( ! [string]::IsNullOrEmpty( $logFile ) )
            {
                Stop-Transcript
            }
            Exit 0
        }
    }

    [long]$saved = 0
    [int]$trimmed = 0

    $params = @{}

    [int[]]$sessionsToTarget = @()
    $results = New-Object -TypeName System.Collections.ArrayList

    if( $disconnected ) 
    {
        ## no native session support so parse output of quser.exe
        ## Columns are 'USERNAME              SESSIONNAME        ID  STATE   IDLE TIME  LOGON TIME' but SESSIONNAME is empty for disconnected so all shifted left by one column (yuck!)

        $sessionsToTarget = @( (quser) -replace '\s{2,}', ',' | ConvertFrom-Csv | ForEach-Object ` 
        {
            $session = $_
            if( $session.Id -like "Disc*" )
            {
                $session.SessionName -as [int]
                Write-Verbose ( "Session {0} is disconnected for user {1} logon {2} idle {3}" -f $session.SessionName , $session.Username , $session.'Idle Time' , $session.State )
            }
        } )
    }

    ## Reform arrays as they will not be passed correctly if command not invoked natively in PowerShell, e.g. via cmd or scheduled task
    if( $processes )
    {
        if( $processes.Count -eq 1 -and $processes[0].IndexOf(',') -ge 0 )
        {
            $processes = $processes -split ','
        }
        $params.Add( 'Name' , $processes )
    }

    if( $processIds )
    {
        if( $processIds.Count -eq 1 -and $processIds[0].IndexOf(',') -ge 0 )
        {
            $processIds = $processIds -split ','
        }
        $params.Add( 'Id' , $processIds )
    }

    if( $includeUsers -or $excludeUsers )
    {
        $params.Add( 'IncludeUserName' , $true ) ## Needs admin rights
        if( $includeUsers.Count -eq 1 -and $includeUsers[0].IndexOf(',') -ge 0 )
        {
            $includeUsers = $includeUsers -split ','
        }
        if( $excludeUsers.Count -eq 1 -and $excludeUsers[0].IndexOf(',') -ge 0 )
        {
            $excludeUsers = $excludeUsers -split ','
        }
    }

    if( $exclude -and $exclude.Count -eq 1 -and $exclude[0].IndexOf(',') -ge 0 )
    {
        $exclude = $exclude -split ','
    }

    if( $sessionIds -and $sessionIds.Count -eq 1 -and $sessionIds[0].IndexOf(',') -ge 0 )
    {
        $sessionIds = $sessionIds -split ','
    }

    if( $notSessionIds -and $notSessionIds.Count -eq 1 -and $notSessionIds[0].IndexOf(',') -ge 0 )
    {
        $notSessionIds = $notSessionIds -split ','
    }

    [int]$adjusted = 0

    Get-Process @params -ErrorAction SilentlyContinue | ForEach-Object `
    {      
        $process = $_
        [bool]$doIt = $true
        if( $excludeUsers -and $excludeUsers.Count -And $excludeUsers -contains $process.UserName )
        {
            Write-Verbose ( "`tSkipping {0} pid {1} for user {2} as specifically excluded" -f $process.Name , $process.Id , $process.UserName )
            $doIt = $false
        }
        elseif( $doIt -and $includeUsers -and $includeUsers.Count -And $includeUsers -notcontains $process.UserName )
        {
            Write-Verbose ( "`tSkipping {0} pid {1} for user {2} as not included" -f $process.Name , $process.Id , $process.UserName )
            $doIt = $false
        }
        elseif( $doIt -and $exclude -and $exclude.Count -And $exclude -contains $process.Name )
        {
            Write-Verbose ( "`tSkipping {0} pid {1} as specifically excluded" -f $process.Name , $process.Id )
            $doIt = $false
        }
        elseif( $doIt -and $thisSession -And $process.SessionId -ne $thisSessionId )
        {
            Write-Verbose ( "`tSkipping {0} pid {1} as session {2} not {3}" -f $process.Name , $process.Id , $process.SessionId , $thisSessionId )
            $doIt = $false
        }
        elseif( $doIt -and $process.Id -eq $activePid -And $idle -eq 0 ) ## if idle then we'll trim anyway as not being used (will have quit already if not idle if idle parameter specified)
        {
            Write-Verbose ( "`tSkipping {0} pid {1} as it is the foreground window process" -f $process.Name , $process.Id )
            $doIt = $false
        }
        elseif( $doIt -and $sessionIds -and $sessionIds.Count -gt 0 -And $sessionIds -notcontains $process.SessionId.ToString() )
        {
            Write-Verbose ( "`tSkipping {0} pid {1} as session {2} not in list" -f $process.Name , $process.Id , $process.SessionId )
            $doIt = $false
        }
        elseif( $notsessionIds -and $notSessionIds.Count -gt 0 -And $notSessionIds -contains $process.SessionId.ToString() )
        {
            Write-Verbose ( "`tSkipping {0} pid {1} as session {2} is specifically excluded" -f $process.Name , $process.Id , $process.SessionId )
            $doIt = $false
        }
        elseif( $doIt -and $sessionsToTarget.Count -gt 0 -And $sessionsToTarget -notcontains $process.SessionId )
        {
            Write-Verbose ( "`tSkipping {0} pid {1} as session {2} which is not disconnected" -f $process.Name , $process.Id , $process.SessionId )
            $doIt = $false
        }
        elseif( $doIt -and $above -gt 0 -And $process.WS -le $above )
        {
            Write-Verbose ( "`tSkipping {0} pid {1} as working set only {2} MB" -f $process.Name , $process.Id , [Math]::Round( $process.WS / 1MB , 1 ) )
            $doIt = $false
        }
        elseif( $doIt -and $newOnly -and $process.StartTime -lt $monitoringStartTime )
        {
            Write-Verbose ( "`tSkipping {0} pid {1} as start time {2} prior to {3}" -f $process.Name , $process.Id , $process.StartTime , $monitoringStartTime )
            $doit = $false
        }

        if( $doIt )
        {
            $action = "Process {0} pid {1} session {2} working set {3} MB" -f $process.Name , $process.Id , $process.SessionId , [Math]::Floor( $process.WS / 1MB )

            if( $process.Handle )
            {
                if( $report )
                {
                    [int]$thisMinimumWorkingSet = -1 
                    [int]$thisMaximumWorkingSet = -1 
                    [int]$thisFlags = -1 ## Grammar alert! :-)
                    ## https://msdn.microsoft.com/en-us/library/windows/desktop/ms683227(v=vs.85).aspx
                    [bool]$result = [PInvoke.Win32.Memory]::GetProcessWorkingSetSizeEx( $process.Handle, [ref]$thisMinimumWorkingSet,[ref]$thisMaximumWorkingSet,[ref]$thisFlags);$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                    if( $result )
                    {
                        ## convert flags value - if not hard then will be soft so no point reporting that separately IMHO
                        [bool]$hardMinimumWorkingSet = $thisFlags -band 1 ## QUOTA_LIMITS_HARDWS_MIN_ENABLE
                        [bool]$hardMaximumWorkingSet = $thisFlags -band 4 ## QUOTA_LIMITS_HARDWS_MAX_ENABLE
                        $null = $results.Add( ([pscustomobject][ordered]@{ 'Name' = $process.Name ; 'PID' = $process.Id ; 'Handle Count' = $process.HandleCount ; 'Start Time' = $process.StartTime ;
                                'Hard Minimum Working Set Limit' = $hardMinimumWorkingSet ; 'Hard Maximum Working Set Limit' = $hardMaximumWorkingSet ;
                                'Working Set (MB)' = $process.WorkingSet64 / 1MB ;'Peak Working Set (MB)' = $process.PeakWorkingSet64 / 1MB ;
                                'Commit Size (MB)' = $process.PagedMemorySize / 1MB; 
                                'Paged Pool Memory Size (KB)' = $process.PagedSystemMemorySize64 / 1KB ; 'Non-paged Pool Memory Size (KB)' = $process.NonpagedSystemMemorySize64 / 1KB ;
                                'Minimum Working Set (KB)' = $thisMinimumWorkingSet / 1KB ; 'Maximum Working Set (KB)' = $thisMaximumWorkingSet / 1KB
                                'Hard Minimum Working Set' = $hardMinimumWorkingSet ; 'Hard Maximum Working Set' = $hardMaximumWorkingSet 
                                'Virtual Memory Size (GB)' = $process.VirtualMemorySize64 / 1GB; 'Peak Virtual Memory Size (GB)' = $process.PeakVirtualMemorySize64 / 1GB; }) )
                    }
                    else
                    {                   
                        Write-Warning ( "Failed to get working set info for {0} pid {1} - {2} " -f $process.Name , $process.Id , $LastError)
                    }
                }
                elseif( $pscmdlet.ShouldProcess( $action , 'Trim' ) ) ## Handle may be null if we don't have sufficient privileges to that process
                {
                    ## see https://msdn.microsoft.com/en-us/library/windows/desktop/ms686237(v=vs.85).aspx
                    [bool]$result = [PInvoke.Win32.Memory]::SetProcessWorkingSetSizeEx( $process.Handle,$minWorkingSet,$maxWorkingSet,$flags);$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

                    $adjusted++
                    if( ! $result )
                    {
                        Write-Warning ( "Failed to trim {0} pid {1} - {2} " -f $process.Name , $process.Id , $LastError)
                    }
                    elseif( $savings )
                    {
                        $now = Get-Process -Id $process.Id -ErrorAction SilentlyContinue
                        if( $now )
                        {
                            $saved += $process.WS - $now.WS
                            $trimmed++
                        }
                    }
                }
            }
            else
            {
                Write-Warning ( "No handle on process {0} pid {1} working set {2} MB so cannot access working set" -f $process.Name , $process.Id , [Math]::Floor( $process.WS / 1MB ) )
            }
        }
    }

    if( $report )
    {
        if( [string]::IsNullOrEmpty( $outputFile ) )
        {
            if( -Not $nogridview )
            {
                $selected = $results | Sort-Object Name | Out-GridView -PassThru -Title "Memory information from $($results.Count) processes at $(Get-Date -Format U)"
                if( $selected )
                {
                    $selected | clip.exe
                }
            }
            else
            {
                $results
            }
        }
        else
        {
            $results | Sort-Object Name | Export-Csv -Path $outputFile -NoTypeInformation -NoClobber
        }
    }

    if( $savings )
    {
        [long]$availableMemoryAfter = (Get-Counter '\Memory\Available MBytes').CounterSamples[0].CookedValue
        Write-Output ( "Trimmed {0}MB from {1} processes giving {2}MB extra available" -f [math]::Round( $saved / 1MB , 1 ) , $trimmed , ( $availableMemoryAfter - ( $availableMemory / 1MB ) ) )
    }
    if( $loop )
    {
        if( $processIds -and $processIds.Count -and ! $adjusted )
        {
            Write-Warning "None of the specified pids $($processIds -join ', ') were found or were not included or were excluded so exiting loop"
            $loop = $false
        }
        else
        {
            Write-Verbose "$(Get-Date) : sleeping for $pollPeriod seconds before looping"
            Start-Sleep -Seconds $pollPeriod
        }
    }
} while( $loop )

if( ! [string]::IsNullOrEmpty( $logFile ) )
{
    Stop-Transcript
}
