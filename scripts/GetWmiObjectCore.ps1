
$CoreWmi = "Get-CimInstance"
$DesktopWmi = "Get-WmiObject"

if($PSVersionTable.PSEdition -eq 'Core'){
    New-Alias -Name "$DesktopWmi" -Value "$CoreWmi" -Option AllScope -Scope Global -Force -ErrorAction Ignore
}
