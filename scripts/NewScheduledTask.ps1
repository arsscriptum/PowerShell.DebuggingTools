


Function Schedule-Task
{
    Param
    (
        [string]$taskFolder ,
        [string]$taskname , 
        [string]$script ,  ## if null then we are deleting
        [int]$idle ,
        [switch]$background ,
        [int]$above ,
        [switch]$savings ,
        [string]$available = $null,
        [string[]]$processes ,
        [string[]]$exclude ,
        [string]$logFile = $null
    )

    Write-Verbose "Schedule-Task( $taskFolder , $taskName , $script )"

    ## https://www.experts-exchange.com/articles/11591/VBScript-and-Task-Scheduler-2-0-Creating-Scheduled-Tasks.html

    Set-Variable TASK_LOGON_INTERACTIVE_TOKEN      3   #-Option Constant
    Set-Variable TASK_RUNLEVEL_LUA                 0   #-Option Constant
    Set-Variable TASK_TRIGGER_EVENT                0   #-Option Constant
    Set-Variable TASK_TRIGGER_TIME                 1
    Set-Variable TASK_TRIGGER_DAILY                2
    Set-Variable TASK_TRIGGER_IDLE                 6
    Set-Variable TASK_TRIGGER_SESSION_STATE_CHANGE 11  #-Option Constant
    Set-Variable TASK_STATE_SESSION_LOCK           7   #-Option Constant
    Set-Variable TASK_STATE_REMOTE_DISCONNECT      4
    Set-Variable TASK_ACTION_EXEC                  0   #-Option Constant
    Set-Variable TASK_CREATE_OR_UPDATE             6   #-Option Constant

    $objTaskService  = New-Object -ComObject "Schedule.Service" ##-Strict
    $objTaskService.Connect()

    $objRootFolder = $objTaskService.GetFolder("\")

    $objTaskFolders = $objRootFolder.GetFolders(0)

    [bool]$blnFoundTask = $false

    ForEach( $objTaskFolder In $objTaskFolders )
    {
        If( $objTaskFolder.Path -eq $taskFolder )
        {
            $blnFoundTask = $True
            break
        }
    }

    if( [string]::IsNullOrEmpty( $script ) )
    {
        ## Find task and delete
        if( $blnFoundTask )
        {
            [bool]$deleted = $false
            $objTaskFolder.GetTasks(0) | ?{ $_.Name -eq $taskname } | %{ $objTaskFolder.DeleteTask( $_.Name , 0 ) ; $deleted = $true }
            if( ! $deleted )
            {
                Write-Warning "Failed to find task `"$taskname`" so cannot remove it"
            }
        }
        else
        {
            Write-Warning "Unable to find task folder $taskFolder so cannot remove scheduled tasks"
        }
        return
    }
    elseif( ! $blnFoundTask )
    {
        $objTaskFolder = $objRootFolder.CreateFolder($taskFolder)
    }

    $objNewTaskDefinition = $objTaskService.NewTask(0) 

    $objNewTaskDefinition.Data = 'This is Guys task from PoSH'

    $objNewTaskDefinition.RegistrationInfo.Author = $objTaskService.ConnectedDomain  + "\" + $objTaskService.ConnectedUser
    $objNewTaskDefinition.RegistrationInfo.Date = ([datetime]::Now).ToString("yyyy-MM-dd'T'HH:mm:ss")
    $objNewTaskDefinition.RegistrationInfo.Description = 'Trim process memory'
    $objNewTaskDefinition.RegistrationInfo.Documentation = 'RTFM'
    $objNewTaskDefinition.RegistrationInfo.Source = 'PowerShell'
    $objNewTaskDefinition.RegistrationInfo.URI = 'http://guyrleech.wordpress.com'
    $objNewTaskDefinition.RegistrationInfo.Version = '1.0'

    $objNewTaskDefinition.Principal.Id = 'My ID'
    $objNewTaskDefinition.Principal.DisplayName = 'Principal Description'
    $objNewTaskDefinition.Principal.UserId = $objTaskService.ConnectedDomain  + "\" + $objTaskService.ConnectedUser
    $objNewTaskDefinition.Principal.LogonType = $TASK_LOGON_INTERACTIVE_TOKEN
    $objNewTaskDefinition.Principal.RunLevel = $TASK_RUNLEVEL_LUA

    $objTaskTriggers = $objNewTaskDefinition.Triggers
    
    $objTaskAction = $objNewTaskDefinition.Actions.Create($TASK_ACTION_EXEC)
    $objTaskAction.Id = 'Execute Action'
    ## powershell.exe even with windowstyle hidden still shows a window so we start via vbs in order for it to be truly hidden
$vbsscriptbody = @"
Dim objShell,strArgs , i
for i = 0 to WScript.Arguments.length - 1
    strArgs = strArgs & WScript.Arguments(i) & " "
next
Set objShell = CreateObject("WScript.Shell")
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File ""$script"" -ThisSession -scheduled " & strArgs , 0

"@
   
    [string]$vbsscript = $script -replace '\.ps1$' , '.vbs'

    if( Test-Path $vbsscript )
    {
        [string]$content = ""
        $existingScript = Get-Content $vbsscript | %{ $content += $_ + "`r`n" }

        if( $content -ne $vbsscriptbody )
        {
            Write-Error "vbs script `"$vbsscript`" already exists but is different to the file we need to write"
        }
    }
    else
    {
        [io.file]::WriteAllText( $vbsscript , $vbsscriptbody ) ## ensure no newline as breaks comparison
        if( ! $? -or ! ( Test-Path $vbsscript ) )
        {
            Write-Error "Error creating vbs script `"$vbsscript`""
        }
    }
    
    $objTaskAction.WorkingDirectory = $env:TEMP
    $objTaskAction.Path = 'wscript.exe'
    $objTaskAction.Arguments = "//nologo `"$vbsscript`""

    if( $idle -gt 0 )
    {
        $objTaskAction.Arguments += " -Idle $install"
    }
    if( $background )
    {
        $objTaskAction.Arguments += ' -background'
    }
    if( $above -ge 0 )
    {
        $objTaskAction.Arguments += " -above $above"
    }
    if( $savings )
    {
        $objTaskAction.Arguments += " -savings"
    }
    if( ! [string]::IsNullOrEmpty( $available ) )
    {
        $objTaskAction.Arguments += " -available $available"
    }
    if( ! [string]::IsNullOrEmpty( $logFile ) )
    {
        $objTaskAction.Arguments += " -logfile `"$logfile`""
    }
    if( $processes -and $processes.Count )
    {
        $objTaskAction.Arguments += " -processes $processes"
    }
    if( $exclude -and $exclude.Count )
    {
        $objTaskAction.Arguments += " -exclude $exclude"
    }  
    if( $VerbosePreference -eq 'Continue' )
    {
        $objTaskAction.Arguments += " -verbose"
    }

    ## http://msdn.microsoft.com/en-us/library/windows/desktop/aa383480%28v=vs.85%29.aspx
    $objNewTaskDefinition.Settings.Enabled = $true
    $objNewTaskDefinition.Settings.Compatibility = 2 ## Win7/WS08R2
    $objNewTaskDefinition.Settings.Priority = 5 ## 0 High - 10 Low
    $objNewTaskDefinition.Settings.Hidden = $false
    
    ## Can't use idle trigger as means more than just no input from user so we run a standard, repeating scheduled task and check for no input in the script itself
    if( $idle -gt 0 )
    {
        $objTaskTrigger = $objTaskTriggers.Create($TASK_TRIGGER_DAILY)
        $objTaskTrigger.Enabled = $true
        $objTaskTrigger.DaysInterval = 1
        $objTaskTrigger.Repetition.Duration = 'P1D'
        $objTaskTrigger.Repetition.Interval = 'PT' + [math]::Round( $idle / 2 ) + 'S'
        $objTaskTrigger.Repetition.StopAtDurationEnd = $true
        $objTaskTrigger.StartBoundary = ([datetime]::Now).ToString('yyyy-MM-dd''T''HH:mm:ss')
    }
    else
    {
        $objTaskTrigger = $objTaskTriggers.Create($TASK_TRIGGER_SESSION_STATE_CHANGE)
        $objTaskTrigger.Enabled = $true
        $objTaskTrigger.Id = 'Session state change lock'
        $objTaskTrigger.StateChange = $TASK_STATE_SESSION_LOCK

        ## Format For Days = P#D where # is the number of days
        ## Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
        $objTaskTrigger.ExecutionTimeLimit = 'PT5M'
        $objTaskTrigger.Delay = 'PT5S'
        $objTaskTrigger.UserId = $objTaskService.ConnectedDomain  + '\' + $objTaskService.ConnectedUser

        ## http://msdn.microsoft.com/en-us/library/windows/desktop/aa382144%28v=vs.85%29.aspx

        $objTaskTrigger = $objTaskTriggers.Create($TASK_TRIGGER_SESSION_STATE_CHANGE)
        $objTaskTrigger.Enabled = $true
        $objTaskTrigger.Id = 'Session state change disconnect'
        $objTaskTrigger.StateChange = $TASK_STATE_REMOTE_DISCONNECT

        ## Format For Days = P#D where # is the number of days
        ## Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
        $objTaskTrigger.ExecutionTimeLimit = "PT5M"
        $objTaskTrigger.Delay = "PT5S"
        $objTaskTrigger.UserId = $objTaskService.ConnectedDomain  + '\' + $objTaskService.ConnectedUser
    }

    $objNewTaskDefinition.Settings.DisallowStartIfOnBatteries = $false
    $objNewTaskDefinition.Settings.AllowDemandStart = $true
    $objNewTaskDefinition.Settings.StartWhenAvailable = $true
    $objNewTaskDefinition.Settings.RestartInterval = 'PT10M'
    $objNewTaskDefinition.Settings.RestartCount = 2
    $objNewTaskDefinition.Settings.ExecutionTimeLimit = "PT1H"
    $objNewTaskDefinition.Settings.AllowHardTerminate = $true
    ## 0 = Run a second instance now (Parallel)
    ## 1 = Put the new instance in line behind the current running instance (Add To Queue)
    ## 2 = Ignore the new request
    $objNewTaskDefinition.Settings.MultipleInstances = 2

    try
    {
        $task = $objTaskFolder.RegisterTaskDefinition( $taskname , $objNewTaskDefinition , $TASK_CREATE_OR_UPDATE , $null , $null , $TASK_LOGON_INTERACTIVE_TOKEN )
    }
    catch
    {
        $task = $null
    }

    if( ! $task )
    {
        Write-Error ( "Failed to create scheduled task: {0}" -f $error[0] )
    }
}

if( ! [string]::IsNullOrEmpty( $logFile ) )
{
    Start-Transcript $logFile -Append
}

if( $install -gt 0 -or $uninstall )
{
    Write-Verbose ( "{0} requested" -f $( if( $uninstall ) { "Uninstall" } else {"Install"} )  )

    if( $uninstall -and $install -gt 0 )
    {
        Write-Error "Cannot specify -install and -uninstall together"
        return 1
    }
    elseif( $report )
    {
        Write-Error "Cannot specify -install or -uninstall with -report"
    }
    elseif( $uninstall )
    {
        $scriptName = $null
    }
    elseif( $install -lt $minimumIdlePeriod ) ## minimum repetition is 1 minute see https://msdn.microsoft.com/en-us/library/windows/desktop/aa382993(v=vs.85).aspx
    {
        Write-Error "Idle time is too low - minimum idle time is $minimumIdlePeriod seconds"
        return
    }
    else
    {
        $scriptName = & { $myInvocation.ScriptName }
    }

    [hashtable]$taskArguments =
    @{
        Taskfolder = $taskFolder 
        Script = $scriptName
        Above = $above 
        Savings = $savings
        Exclude = $exclude
        Processes = $processes
        Logfile = $logFile
        Available = $available
    }

    Schedule-Task -taskName "Trim on lock and disconnect for $($env:username)" @taskArguments
    Schedule-Task -taskName "Trim idle for $($env:username)" @taskArguments -idle $install -background $background

    ## if we have been asked to hook logoff, to uninstall, then we create a hidden window so we can capture events
    if( $logoff )
    {
        Add-Type –AssemblyName System.Windows.Forms 

        $form = New-Object Windows.Forms.Form
        $form.Size = New-Object System.Drawing.Size(0,0)
        $form.Location = New-Object System.Drawing.Point(-5000,-5000)
        $form.FormBorderStyle = 'FixedToolWindow'
        $form.StartPosition = 'manual'
        $form.ShowInTaskbar = $false
        $form.WindowState = 'Normal'
        $form.Visible = $false
        $form.AutoSize = $false

        $form.add_FormClosing(
            {
               Write-Verbose "$(Get-Date) dialog closing" 
               Schedule-Task -taskName "Trim on lock and disconnect for $($env:username)" -taskFolder $taskFolder -Script $null 
               Schedule-Task -taskName "Trim idle for $($env:username)" -taskFolder $taskFolder -Script $null
            })

        $form.add_Load({ $form.Opacity = 0 })

        $form.add_Shown({ $form.Opacity = 100 })

        Write-Verbose "About to show dialog for logoff intercept - script will not exit until logoff"
        [void]$form.ShowDialog()
        ## We will only get here when the hidden dialogue exits which should only be logoff
    }

    if( ! [string]::IsNullOrEmpty( $logFile ) )
    {
        Stop-Transcript
    }

    Exit 0
}
