DO
 {
    $crashes = wevtutil.exe qe .\event_logger_query.dat /sq:true
    if ($crashes) {
        1 | Out-File C:\Users\Brendan\tools\scripts\overclocker\crashed.dat -NoNewLine
    } else {
        0 | Out-File C:\Users\Brendan\tools\scripts\overclocker\crashed.dat -NoNewLine
        Start-Sleep -s 10
    }
} While (!$crashes)
