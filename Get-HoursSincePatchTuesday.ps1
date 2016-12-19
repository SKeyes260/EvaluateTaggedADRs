Function Get-PatchTuesdayForDate {
    [CmdletBinding()] 
    PARAM( [Parameter(Position=1)] $VarDateTime )

    $FindNthDay=2
    $WeekDay='Tuesday'
    [datetime]$Today = $VarDateTime
    $todayM=$Today.Month.ToString()
    $todayY=$Today.Year.ToString()
    [datetime]$TestDate=$TodayM+'/01/'+$TodayY
    $Counter = 0
    while ($Counter -lt $FindNthDay ) {
        If ( $TestDate.DayofWeek -eq $WeekDay ) { $Counter += 1 }
        If ( $Counter -ne $FindNthDay ) { $TestDate=$TestDate.AddDays(1) }
    }
    Return $TestDate.Date
}


##########################
# MAIN
##########################

$Today = get-Date

$PT = Get-PatchTuesdayForDate -VarDateTime $Today

$HoursDiff = $Today - $PT

$HoursDiff

#$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")