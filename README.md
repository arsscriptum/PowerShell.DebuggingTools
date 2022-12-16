# How to Capture Debugging Information for Memory Leak in Windows

## Objective
This article describes how to capture IMA debugging information for an IMA memory leak in XenApp on Windows 2008 R2. This procedure can be used for any process.
This document also discusses a sample scenario of recording memory leak using PowerShell for the firefox.exe process.

## Requirements
To complete the procedure, access the following components:

- Windows Debugging Tools
- ProcDump

The following components are optional:

- PowerShell
- PowerShell ISE


## Instructions
Recording IMA Memory Leak for XenApp on Windows 2008 R2 Operating System
To record the IMA Memory leak, complete the following procedure:

1) Download the 32-bit version of the Windows Debugging Tools from the Microsoft website, see here. 

2) Install the 32-bit version 6.11.1.404 of the tool.
Note: The older version is required to avoid a 64/32-bit mismatch in gflags.exe and umdh.exe, which are running on a 32-bit process. In this procedure, ImaSrv.exe is the process.

3) Create the following folders to save symbols, log files, and dump files on the computer:
C:\Symbols\CTX
C:\Symbols\MS
C:\Logs
C:\Dumps

4) Create a new environment variable for the Symbols folder with the following details:

Variable  name: _NT_SYMBOL_PATH_

Variable value: SRV*c:\Symbols \MS*http://msdl.microsoft.com/download/symbols;SRV*c:\Symbols\CTX*http://ctxsym.citrix.com/symbols

5) Run the following command from the command prompt to enable User Mode Stack Trace Database:
gflags -i ImaSrv.exe +ust
Note: You can also enable the User Mode Stack Trace Database through the GFlags GUI.
To enable the User Mode Stack Trace Database through the GUI, complete the following procedure:

Start Gflags.exe.
In the GUI, activate the Image File tab and specify the required details, as shown in the following screen shot.

6) Restart the IMA Service.

7) Download ProcDump from Windows Sysinternals - ProcDump v3.04 .

Move procdump.exe to the C:\windows\system32 folder.
Moving the file to this folder ensures that the file is available in the $PATH folder.

8) To create a dump file for the ImaSrv.exe process and to create a log file, complete the following procedure:

1) Make a note of the process ID from the Windows Task Manager, as shown in the following screen shot.User-added image
2) Run the following command to create a log file:
```
    procdump –ma 1340 c:\Dumps\
    C:\Program Files (x86)\Debugging Tools for Windows (x86)>umdh -p:1340 -f:c:\Logs\Log1.txt
```

Note: The procdump command requires the –ma parameter to record the complete memory space for the process.


Run the script from PowerShell or from the command prompt. Run the following command from PowerShell:
```
    PS> New-Capture.ps1 PID
```


The following is a sample output for the preceding command:

```dos
The process is running, current memory usage is 27.63671875 - preparing to dump (firefox)...
ProcDump v1.81 - Writes process dump files
Copyright (C) 2009-2010 Mark Russinovich
Sysinternals - www.sysinternals.com
Process:            firefox.exe (5880)
CPU threshold:      n/a
Commit threshold:   27 MB
Threshold seconds:  10
Number of dumps:    1
Hung window check:  Disabled
Exception monitor:  Disabled
Terminate monitor:  Disabled
Dump file:          c:\Dumps\dump1.dmp
Time        CPU  Duration
Process has hit memory usage spike threshold.
Writing dump file c:\Dumps\dump1_100810_134059.dmp...
Dump written.
ProcDump v1.81 - Writes process dump files
Copyright (C) 2009-2010 Mark Russinovich
Sysinternals - www.sysinternals.com
Process:            firefox.exe (5880)
CPU threshold:      n/a
Commit threshold:   47 MB
Threshold seconds:  10
Number of dumps:    1
Hung window check:  Disabled
Exception monitor:  Disabled
Terminate monitor:  Disabled
Dump file:          c:\Dumps\dump2.dmp
Time        CPU  Duration
Process has hit memory usage spike threshold.
Writing dump file c:\Dumps\dump2_100810_134129.dmp...
Dump written.
ProcDump v1.81 - Writes process dump files
Copyright (C) 2009-2010 Mark Russinovich
Sysinternals - www.sysinternals.com
Process:            firefox.exe (5880)
CPU threshold:      n/a
Commit threshold:   67 MB
Threshold seconds:  10
Number of dumps:    1
Hung window check:  Disabled
Exception monitor:  Disabled
Terminate monitor:  Disabled
Dump file:          c:\Dumps\dump3.dmp
Time        CPU  Duration
Process has hit memory usage spike threshold.
Writing dump file c:\Dumps\dump3_100810_134136.dmp...
Dump written.
ProcDump v1.81 - Writes process dump files
Copyright (C) 2009-2010 Mark Russinovich
Sysinternals - www.sysinternals.com
Process:            firefox.exe (5880)
CPU threshold:      n/a
Commit threshold:   87 MB
Threshold seconds:  10
Number of dumps:    1
Hung window check:  Disabled
Exception monitor:  Disabled
Terminate monitor:  Disabled
Dump file:          c:\Dumps\dump4.dmp
Time        CPU  Duration
Process has hit memory usage spike threshold.
Writing dump file c:\Dumps\dump4_100810_134204.dmp...
Dump written.
```