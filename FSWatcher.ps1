$searchPath = 'D:\Temp'
$fs_watcher = New-Object System.IO.FilesystemWatcher
$fs_watcher.Path = $searchPath
$fs_watcher.IncludeSubdirectories = $true
$fs_watcher.EnableRaisingEvents = $true
#
$changed = Register-ObjectEvent $fs_watcher "Changed" -Action {    write-host "$i : Changed: $($eventArgs.FullPath)" }
$created = Register-ObjectEvent $fs_watcher "Created" -Action {   write-host "$i : Created: $($eventArgs.FullPath)" }
$deleted = Register-ObjectEvent $fs_watcher "Deleted" -Action {   write-host "$i : Deleted: $($eventArgs.FullPath)"}
$renamed = Register-ObjectEvent $fs_watcher "Renamed" -Action { @{'Time'=Get-Date; 'Status'='Renamed'; 'Path'=$eventArgs.FullPath}}
