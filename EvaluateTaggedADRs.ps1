# Evaluate Tagged ADRs for execution
# Any ADR with *SCHED-<DateCode>+[<HoursOffset>]* in its comment will be evaluated for execution
# The *SCHED-<DateCode>+[<HoursOffset>]*  tags are intended for monthly ADR's that require a complex schedule based off of the Patch Tuesday for a month

param([string]$SiteServer, [string]$SiteCode , [string]$InstanceName ) 


# This function will evaluate the current date against the tag text execute the ADR if appropriate
Function Evaluate-SCHEDTag {
    [CmdletBinding()] 
    PARAM(
        [Parameter(Position=1)] $TAG,
        [Parameter(Position=2)] [datetime]$EvaluationDate
    )

    $objTAG = New-Object –TypeName PSObject
    $objTAG | Add-Member –MemberType NoteProperty –Name TAG –Value $TAG
    $objTAG | Add-Member –MemberType NoteProperty –Name Recurrance –Value $TAG.SubString(0,1)
    $objTAG | Add-Member –MemberType NoteProperty –Name Weekday –Value  $TAG.Substring(2,2)
    $objTAG | Add-Member –MemberType NoteProperty –Name Interval –Value  $TAG.Substring(4,1)
    $objTAG | Add-Member –MemberType NoteProperty –Name Offset –Value $TAG.SubString(6,($Tag.Length - 7))
    $objTAG | Add-Member –MemberType NoteProperty –Name OffsetHours –Value $TAG.SubString(6,($Tag.Length - 9))
    $objTAG | Add-Member –MemberType NoteProperty –Name OffsetMinutes –Value $TAG.SubString(($Tag.length - 2),2)

    Log-Append -strLogFileName  $LogFileName  -strLogText  ("Extracted TAG Properties :"+$objTAG.Recurrance+" "+$objTAG.Weekday+" "+$objTAG.Interval+" "+$objTAG.OffsetHours+" "+$objTAG.OffsetMinutes)
    #If (!$ADR.LastRunTime) { $LastRunTime = [dateTime]("00/00/0000") } ELSE { $LastRunTime = [datetime](Convert-FromWMIDAte $ADR.LastRunTime) }
    $LastRunTime = $ADR.LastRunTime
    If ($LastRunTime) { $LastRunTime = Convert-FromWMIDate $LastRunTime }

    If ( $objTAG.Recurrance -eq "M" ) {  # MONTHLY
        $ExecutionDate = Get-MonthlyExecutionDate -objTAG $objTAG -EvalDate $EvaluationDate
        Log-Append -strLogFileName  $LogFileName  -strLogText  ("The Monthly Execution Date for this ADR is calculated to be : "+$ExecutionDate)  
        If (!$ADR.LastRunTime) {
            Log-Append -strLogFileName  $LogFileName  -strLogText  ("The ADR has never been executed before.") 
            If ($ExecutionDate -lt $EvaluationDate ) {
                Log-Append -strLogFileName  $LogFileName  -strLogText  ("It is not yet time to execute this ADR. Nothing to do")   
                $Result = Execute-ADR -SiteServer $SiteServer -SiteCode $SiteCode -ADRName $ADR.Name
            }
            ELSE { Log-Append -strLogFileName  $LogFileName  -strLogText  ("The ADR needs to be executed,  Executing the ADR named "+$ADR.Name+" on "+$SiteServer+":"+$SiteCode )  }
        }
        ELSE {
            Log-Append -strLogFileName  $LogFileName  -strLogText  ("The ADR named "+$ADR.Name+" was last executed  "+$LastRunTime) 
            If ( $LastRunTime -ge $ExecutionDate )  { Log-Append -strLogFileName  $LogFileName  -strLogText  ("The ADR has already executed for this month, nothing to do." ) }
            ELSE {
            If ($ExecutionDate -lt $EvaluationDate ) {
                Log-Append -strLogFileName  $LogFileName  -strLogText  ("The ADR needs to be executed,  Executing the ADR named "+$ADR.Name+" on "+$SiteServer+":"+$SiteCode )   
                $Result = Execute-ADR -SiteServer $SiteServer -SiteCode $SiteCode -ADRName $ADR.Name
            }
            ELSE { Log-Append -strLogFileName  $LogFileName  -strLogText  ("It is not yet time to execute this ADR. Nothing to do") }
            }
        }
    }

    Return $objTAG
}


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


# This function will return the SCHED tag from a string
Function Get-SCHEDTagFromString {
    [CmdletBinding()] 
    PARAM( [Parameter(Position=1)] $TagText )

    $SearchStart = "*SCHED-"
    $SearchEnd = "*"
    $TagStartIndex = $TagText.IndexOf($SearchStart,0)+ $SearchStart.length
    $TagEndIndex = $TagText.IndexOf($SearchEnd,$TagStartIndex) 
    Return $TagText.Substring($TagStartIndex,$TagEndIndex-$tagStartIndex )
}




# This function will return the execution date for a given TAG and eval date
Function Get-MonthlyExecutionDate {
    [CmdletBinding()] 
    PARAM(
        [Parameter(Position=1)] $objTAG,
        [Parameter(Position=2)] $EvalDate
 )

    SWITCH  ($objTAG.Weekday) {
         "Su"  { $Weekday = "Sunday"    }
         "Mo"  { $Weekday = "Monday"    }
         "Tu"  { $Weekday = "Tuesday"   }
         "We"  { $Weekday = "Wednesday" }
         "Th"  { $Weekday = "Thursday"  }
         "Fr"  { $Weekday = "Friday"    }
         "Sa"  { $Weekday = "Saturday"  }
    }
    
    $todayM=$EvalDate.Month.ToString()
    $todayY=$EvalDate.Year.ToString()
    [datetime]$TestDate=($TodayM+'/01/'+$TodayY+" 00:00")
    $Counter = 0
    while ($Counter -lt $objTAG.Interval ) {
        If ( $TestDate.DayofWeek -eq $WeekDay ) { $Counter += 1 }
        If ( $Counter -ne $objTAG.Interval ) { $TestDate=$TestDate.AddDays(1) }
    }
    [datetime]$TestDate = $TestDate.AddHours($objTAG.OffsetHours)
    [datetime]$TestDate = $TestDate.AddMinutes($objTAG.OffsetMinutes)
    Return $TestDate
}




Function Convert-FromWMIDate {
    [CmdletBinding()] 
    PARAM( [Parameter(Position=1)] $CIMDateTime )

    $sWBEM = New-Object -ComObject wbemscripting.swbemdatetime 
    $sWBEM.Value = $CIMDateTime
    Return $swbem.GetVarDate() 
}


Function Convert-ToWMIDate {
    [CmdletBinding()] 
    PARAM( [Parameter(Position=1)] $VarDateTime )

    $wmidate = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime($VarDateTime)
   Return $WMIDate
}


Function Execute-ADR {
   [CmdletBinding()] 
    PARAM(
        [Parameter(Position=1)] $SiteCode,
        [Parameter(Position=2)] $SiteServer,
        [Parameter(Position=3)] $ADRName
    )
    
    $ADRs=@()
    $ADRs = Get-WMIObject -ComputerName $SiteServer -Namespace "Root\SMS\Site_$SiteCode" -Class "sms_AutoDeployment" -Filter "Name = '$ADRName'"
    foreach ($Rule in $ADRs) { $Result = $Rule.EvaluateAutoDeployment() }
    Return $Result 
}



Function Log-Append () {
[CmdletBinding()]   
PARAM 
(   [Parameter(Position=1)] $strLogFileName,
    [Parameter(Position=2)] $strLogText )
    
    $strLogText = ($(get-date).tostring()+" ; "+$strLogText.ToString()) 
    Out-File -InputObject $strLogText -FilePath $strLogFileName -Append -NoClobber
}


###################################
#              MAIN               #
###################################


$TempFolder = "C:\TEMP\EvaluateTaggedADRs\"
$TodaysDate = ([datetime]::NOW)
 #$TodaysDate = ([datetime]"2016/07/13 00:00")    # for testing
$TodaysWMIDate = Convert-ToWMIDate ($TodaysDate)
$TodaysWMIDate = ($TodaysWMIDate.SubString(0,21)+"+***")
$CurrentMonth = $TodaysWMIDate.SubString(4,2)
$CurrentYear = $TodaysWMIDate.SubString(0,4)
$LogFileName = ($TempFolder+$InstanceName+"-"+$TodaysDate.Year+$TodaysDate.Month.ToString().PadLeft(2,"0")+$TodaysDate.Day.ToString().PadLeft(2,"0")+".log")
$PatchTuesday = Get-PatchTuesdayForDate($TodaysDate)
WRITE-Host $TodatsDate  $SiteServer $SiteCode  $InstanceName


If (!(Test-Path $TempFolder))  { $Result = New-Item $TempFolder -type directory }

# ______________________________________________________________________
# Milestone : Start  
Log-Append -strLogFileName  $LogFileName  -strLogText ("Starting the script to execute any ADR tagged with the *SCHED-<Code>+HoursOffset* tag...")
Log-Append -strLogFileName  $LogFileName  -strLogText ("    Todays Date : $TodaysDate"  )
Log-Append -strLogFileName  $LogFileName  -strLogText ("    Patch Tuesday : $PatchTuesday")
Log-Append -strLogFileName  $LogFileName  -strLogText ("    Site Server : $SiteServer"  )
Log-Append -strLogFileName  $LogFileName  -strLogText ("    Site Code : $SiteCode"    )


$ADRs = @() 
$ADRs = Get-WMIObject -ComputerName $SiteServer -Namespace ("root\sms\Site_"+$SiteCode) -Query "SELECT * FROM SMS_AutoDeployment Where Description LIKE '%*SCHED-%*%'"
Log-Append -strLogFileName  $LogFileName  -strLogText ("Identified "+@($ADRs).count+" Automatic Deployment Rules with the SCHED tag")

$ADRCount = 0
ForEach ( $ADR in $ADRs ) { 
    $ADRCount+=1 
    Log-Append -strLogFileName  $LogFileName  -strLogText ("#"+$ADRCount+"    ADRName("+$ADR["Name"]+")  LastRunTime("+$ADR.LastRunTime+")  Description("+$ADR["Description"]+")") 
    $TAG = Get-SCHEDTagFromString $ADR["Description"]
    $SCHED = Evaluate-SCHEDTag $TAG -EvaluationDate $Todaysdate

}
Log-Append -strLogFileName  $LogFileName  -strLogText ("Script Finished") 
Log-Append -strLogFileName  $LogFileName  -strLogText ("") 



