# Memory Control Script – Reclaiming Unused Memory

Here we will cover the use of the script to trim working sets of processes such that more memory becomes available.

The working set of a process is defined [here](https://msdn.microsoft.com/en-us/library/windows/desktop/cc441804(v=vs.85).aspx) which defines it as "the set of pages in the virtual address space of the process that are currently resident in physical memory".

Well, what it means is that processes can grab memory but not necessarily actually need to use it. This can be memory leaks or buffers and other pieces of memory that the developer(s) of an application have requested but, for whatever reasons, aren’t currently using.

That memory could be used by other processes, for other users on multi-session systems, but until the application returns it to the operating system, it can’t be-reused.

## Memory trimming 

Memory trimming is where the OS forces processes to empty their working sets. They don’t just discard this memory, since the processes may need it at a later juncture and it could already contain data, instead the OS writes it to the page file for them such that it can be retrieved at a later time if required. Windows will force memory trimming if available memory gets too low but at that point it may be too late and it is indiscriminate in how it trims.

The [following PowerShell script]() uses the API's [SetProcessWorkingSetSizeEx](https://msdn.microsoft.com/en-us/library/windows/desktop/ms686237(v=vs.85).aspx)


## Working Sets

The "working set" is short hand for "parts of memory that the current algorithm is using" and is determined by which parts of memory the CPU just happens to access. It is totally automatic to you. If you are processing an array and storing the results in a table, the array and the table are your working set.

The memory figures in question aren't actually a reliable indicator of how much memory a process is using.

A brief explanation of each of the memory relationships:

- Private Bytes are what the process is allocated, also with pagefile usage.
- Working Set is the non-paged Private Bytes plus memory-mapped files.
- Virtual Bytes are the Working Set plus paged Private Bytes and standby list.

the peak working set is the maximum amount of physical RAM that was assigned to the process in question.

It is not the maximum memory used at some time ("peak"), it's a coincidence that you have roughly the same number there. It is the presently used amount (used by "everyone", that is all programs and the OS).

The peak working set is a different thing. The working set is the amount of memory in a process (or, if you consider several processes, in all these processes) that is currently in physical memory. The peak working set is, consequently, the maximum value so far seen.
A process may allocate more memory than it actually ever commits ("uses"), and most processes will commit more memory than they have in their working set at one time. This is perfectly normal. Pages are moved in and out of working sets (and into the standby list) to assure that the computer, which has only a finite amount of memory, always has enough reserves to satisfy any memory needs.

## PowerShell Garbage Collection

Simply use…

```
    [System.GC]::GetTotalMemory(‘forcefullcollection’) | out-null
```
or

```
    [System.GC]::GetTotalMemory($true) | out-null
```

Where $true or ```‘forceFullCollection’``` is used to indicate that this method can wait for garbage collection to occur before returning.

Even though [System.GC]::Collect() does not force a garbage collection whilst executing in a pipeline or loop of an object, [System.GC]::GetTotalMemory() with either $true or ‘forceFullCollection’ does indeed successfully force a garbage collection. Wow! Go figure! My testing so far has found this to be a reliable method to use across different versions of PowerShell.