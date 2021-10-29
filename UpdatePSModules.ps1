$modules = Get-installedmodule
$count = $modules.count
$percent = 0
$i = 0

foreach ($module in $modules)
{
    $percent = [math]::Round($I/$count*100)
    Write-Progress -Activity "Progress" -Status "$percent% Complete" -PercentComplete $percent
    write-host "Checking $($module.name)"
    $module | Update-Module -Force
    $I ++
}
