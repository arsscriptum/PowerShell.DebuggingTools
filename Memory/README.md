# Memory Control Script – Reclaiming Unused Memory

Here we will cover the use of the script to trim working sets of processes such that more memory becomes available.

The working set of a process is defined [here](https://msdn.microsoft.com/en-us/library/windows/desktop/cc441804(v=vs.85).aspx) which defines it as "the set of pages in the virtual address space of the process that are currently resident in physical memory".

Well, what it means is that processes can grab memory but not necessarily actually need to use it. This can be memory leaks or buffers and other pieces of memory that the developer(s) of an application have requested but, for whatever reasons, aren’t currently using.

That memory could be used by other processes, for other users on multi-session systems, but until the application returns it to the operating system, it can’t be-reused.

## Memory trimming 

Memory trimming is where the OS forces processes to empty their working sets. They don’t just discard this memory, since the processes may need it at a later juncture and it could already contain data, instead the OS writes it to the page file for them such that it can be retrieved at a later time if required. Windows will force memory trimming if available memory gets too low but at that point it may be too late and it is indiscriminate in how it trims.

The [following PowerShell script]() uses the API's [SetProcessWorkingSetSizeEx](https://msdn.microsoft.com/en-us/library/windows/desktop/ms686237(v=vs.85).aspx)