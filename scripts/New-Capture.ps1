

# 2008 R2 Memory Leak Capture Script
if($args.length -lt 1) {
    Write-Warning "You need to provide a PID as an argument";
    exit;
}
# Get the process ID that is passed as a parameter
$process_id = $args[0];
# Increment to use for memory dumps in MEGABYTES
$memory_increment = 20;
if(Get-Process -Id $process_id -ErrorAction SilentlyContinue) {
    # Set the Process Object
    $tmp = Get-Process -Id $process_id;
    $mem = $tmp.PrivateMemorySize / 1024 / 1024;
    $process_name = $tmp.ProcessName;
    Write-Host "The process is running, current memory usage is $mem - preparing to dump ($process_name)..."
    for($i = 1; $i -lt 5; $i++) {
        $dump = "dump" + $i;
        Write-Host $inc;
        procdump -m $mem -ma $process_id c:\Dumps\$dump;
        umdh.exe -p:$process_id -f:c:\Logs\Log$i.txt
        $mem += $memory_increment;
    }
}
else {
    Write-Warning "The process is not running.  You can view the running processes below.  Please try again";
    Get-Process;
    exit;
}