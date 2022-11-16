$SH_APP = New-Object -ComObject Shell.Application
$items = Get-ChildItem -Recurse | ? { $_.Name -match "(.mp3|.ogg|.wav)" } | select DirectoryName, Name

$length = [timespan]::new(0)
foreach ($item in $items) {
    $folder = $SH_APP.Namespace($item.DirectoryName)
    $details = $folder.ParseName($item.Name)
    $time = $folder.GetDetailsOf($details, 27) # Propaty what is able to get the Voice-Data's Time
    $sec = [timespan]::Parse($time)
    $length += $sec
}

Write-Host "--------------"
Write-Host $length.TotalMinutes
Write-Host "--------------"
