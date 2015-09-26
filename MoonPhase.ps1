#region Links
<#
http://en.wikipedia.org/wiki/Saros_(astronomy)
http://en.wikipedia.org/wiki/Year#Astronomical_years
http://en.wikipedia.org/wiki/Month

http://archive.org/stream/treatiseonspheri00balluoft/treatiseonspheri00balluoft_djvu.txt
https://archive.org/details/treatiseonspheri00balluoft

http://aa.usno.navy.mil/data/docs/JulianDate.php
http://quasar.as.utexas.edu/BillInfo/JulianDateCalc.html

#>
#endregion



#region Example Calls
# Starting point is Local date/time (along with time zone and DST offset)...
$Time = LocalDateTimeToJDN 2014 10 11 12 51 00 -DST 1 -ZoneCorrection 0

$JDN = (LocalDateTimeToJDN 2014 10 12 19 45 00 -DST 1 -ZoneCorrection 0).JDN

$SMA = SunMeanAnomaly $JDN
$MMA = MoonMeanAnomaly $JDN
$MLong = MoonLong $JDN
$MLat = MoonLat $JDN  # For (2014/10/11 12:51:00 -DST 1) returns -ve? (-3.2995540...)
$SL = SunLong $JDN
$NL = NutatLong $JDN

$JDN = (LocalDateTimeToJDN -SystemDate).JDN
$Moon = Moon $JDN

Moon (LocalDateTimeToJDN -SystemDate).JDN

#endregion


$(For ($Month=10; $Month -le 11; $Month++) {

    For ($Day = 1; $Day -le 31; $Day++) {

        $Moon = Moon (LocalDateTimeToJDN 2014 $Month $Day 21 30 00 -DST 1).JDN

        $PhaseLine = '#' * [Math]::Round($Moon.Phase * 20 ,0)


        $Appearance = Switch ($Moon.Phase) {
            {$_ -lt 0.01} {'New Moon'}
            {$_ -gt 0.4 -and $_ -lt 0.6} {'Quadrature'}
            {$_ -gt 0.98} {'Full Moon'}
            Default {''}
        }

        $Moon | 
        Add-Member -MemberType NoteProperty -Name Date -Value ("$Day $(If ($Month-eq 10){'October'}else{'November'})") -PassThru |
        Add-Member -MemberType NoteProperty -Name Appearance -Value $Appearance -PassThru |
        Add-Member -MemberType NoteProperty -Name PhaseLine -Value $PhaseLine -PassThru
    }
}) | ft -AutoSize



Function LocalDateTimeToJDN {
# Local Civil Time to JDN and Greenwich Date
# Returns [PsCustomObject] containing JDN and other formats

    [CmdletBinding(DefaultParameterSetName = 'SpecificDate')]
    [OutputType([PsCustomObject])]
    Param (
        [Parameter(ParameterSetName = 'SpecificDate', Mandatory = $True, Position = 0)]
        [Double]$LCYear,    # Local Civil Year
        [Parameter(ParameterSetName = 'SpecificDate', Mandatory = $True, Position = 1)]
        [Double]$LCMonth,   # Local Civil Month
        [Parameter(ParameterSetName = 'SpecificDate', Mandatory = $True, Position = 2)]
        [Double]$LCDay,     # Local Civil Day

        [Parameter(ParameterSetName = 'SpecificDate', Position = 3)]
        [Double]$LCHour = 0,    # Local Civil Hour
        [Parameter(ParameterSetName = 'SpecificDate', Position = 4)]
        [Double]$LCMinute = 0,  # Local Civil Minute
        [Parameter(ParameterSetName = 'SpecificDate', Position = 5)]
        [Double]$LCSecond = 0,  # Local Civil Second
         
        [Parameter(ParameterSetName = 'SpecificDate', Position = 6)]
        [Double]$ZoneCorrection = 0,   # Offset for local timezone, east or west of Greenwich
        [Parameter(ParameterSetName = 'SpecificDate', Position = 7)]
        [Double]$DST = 0,       # Offset for Daylight Saving Time

        [Parameter(ParameterSetName = 'SystemDate', Position = 0)]
        [Switch]$SystemDate
    ) 

    If ($SystemDate) {
        # Get the current date and time values from the system clock
        $Date = (Get-Date).ToUniversalTime()   # Get time, convert to UT

        # Use Get-Date again to split out the individual date/time components
        # (could use the -f operator with a format string, but this is easier...)
        $UtYear, $UtMonth, $UtDay, $UtHour, $UtMinute, $UtSecond = (get-date $Date -Format 'yyyy M d H m s').Split()
        $GmtHour = ([Math]::Abs($UtSecond)/60 +[Math]::Abs($UtMinute))/60 +[Math]::Abs($UtHour)


        #### Fix this later....!
        $LCYear = $UtYear
        $LCMonth = $UtMonth
        $LCDay = $UtDay

        $DST = If ((Get-Date).IsDaylightSavingTime()) {1} else {0}
        $ZoneCorrection = 0

    }
    else {
        # Use specific date/time values passed as params

        # Convert Hours-Minutes-Seconds to Decimal Hours
        $DecimalHours = ([Math]::Abs($LCSecond)/60 +[Math]::Abs($LCMinute))/60 +[Math]::Abs($LCHour)
    
        If (($LCHour -lt 0) -Or ($LCMinute -lt 0) -Or ($LCSecond -lt 0)) {
            $DecimalHours = -$DecimalHours
        }
        $GmtHour = $DecimalHours - $DST - $ZoneCorrection   # Convert to GMT by correcting for TimeZone and Daylight Time
    }


    ##### for dates around midnight 31 Dec, the UT date may be different from the local date...?

    # Now convert the date to Julian Day Number (JDN)
    # Move January and February dates into the previous year to enable to "average days per month" calculation to work


    If ($LCMonth -lt 3) {
        $Y = $LCYear - 1
        $M = $LCMonth + 12
    }
    else {
        $Y = $LCYear
        $M = $LCMonth
    }


    # Calculate around 1582.  Leap year contributions are only required after this

    $Century = Fix ($Y / 100)

    # The following is a correction for dates after 15 October 1582 to fix up the following:
    # 1. 10 days were skipped to bring the calendar back in line with solar time (these days had accumulated because the earlier leap year rules were not sufficiently accurate)
    # 2. After this date, centuries are not leap years unless they are divisible by 400

    $LeapDays = 2 - $Century + (Fix ($Century / 4))   # Assuming date is after 15 October 1582 (for now...)
    
    If ($LCYear -le 1582) {
        If (($LCYear -eq 1582) -And ($LCMonth -gt 10)) {
            # nop
        }
        Else {
            If (($LCYear -eq 1582) -And ($LCMonth -eq 10) -And ($LCDay -ge 15)) {
                # Nop
            }
            Else {
                $LeapDays = 0
            }
        }
    }
          
    If ($Y -lt 0) {
        $YearDays = Fix ((365.25 * $Y) - 0.75)
    }
    Else {
        $YearDays = Fix (365.25 * $Y)
    }

    # Cumulative days in months of the year (average month is ~30.6 days)
    # (This is only valid if January and February dates are moved into the previous year...)

    $MonthDays = Fix (30.6001 * ($M + 1))
    
    # Add all the terms.  1720994.5 is the number of days between midday on 1st January 4713 BCE and Midnight on 1st January 0001

    $JDN =  $YearDays + $LeapDays + $MonthDays + $LCDay + 1720994.5 + $GmtHour/24 


    # (Julian Day Number to Calendar Date)

    $I = Fix ($JDN + 0.5)
    $F = $JDN + 0.5 - $I
    $A = Fix (($I - 1867216.25) / 36524.25)
    
    If ($I -gt 2299160) {
        $B = $I + 1 + $A - (Fix ($A / 4))
    } 
    Else {
        $B = $I
    }
    
    #### Next steps need to be repeated for Local and GMT dates (e.g. to cater for dates/times around midnight Dec 31)

    $C = $B + 1524
    $D = Fix (($C - 122.1) / 365.25)
    $E = Fix (365.25 * $D)
    $G = Fix (($C - $E) / 30.6001)
    
    #### ZoneCorrection and DST variables are not set when using -SystemDate
    $JDCDecimalDay = $C - $E + $F - (Fix (30.6001 * $G)) + $ZoneCorrection/24 + $DST/24
    $JDCDay = Fix $JDCDecimalDay

    $JDCDecimalHour = (Frac $JDCDecimalDay) * 24
    $JDCHour = Fix $JDCDecimalHour
    $JDCMinute = Fix ((Frac $JDCDecimalHour) * 60)
    $JDCSecond = [Int](($JDCDecimalHour * 3600) % 60)
    

    $GmtDecimalDay = $C - $E + $F - (Fix (30.6001 * $G))
    $GmtDay = Fix $GmtDecimalDay

    $GmtDecimalHour = (Frac $GmtDecimalDay) * 24
    $GmtHour = Fix $GmtDecimalHour
    $GmtMinute = Fix ((Frac $GmtDecimalHour) * 60)
    $GmtSecond = [Int](($GmtDecimalHour * 3600) % 60)

    #### Next steps need to be repeated for Local and GMT dates (e.g. to cater for dates/times around midnight Dec 31)

    If ($G -lt 13.5) {
        $JDCMonth = $G - 1
    }
    Else {
        $JDCMonth = $G - 13
    }
    
    If ($JDCMonth -gt 2.5) {
        $JDCYear = $D - 4716
    } 
    Else {
        $JDCYear = $D - 4715
    }

    # T is the time since 1st Jan 1900 (epoch) expressed in units of Centuries 
    # Function CalcT {

    $DJ = $JDN - 2415020    # Days since midday 0m 0s 1900/1/1 
    $T = $DJ / 36525   # Convert to centuries 
    $T2 = $T * $T

    [PsCustomObject]@{
        JulianDayNumber = $JDN
        JDN             = $JDN
        LocalYear       = $JDCYear
        LocalMonth      = $JDCMonth
        LocalDay        = $JDCDay
        LocalHour       = $JDCHour
        LocalMinute     = $JDCMinute
        LocalSecond     = $JDCSecond
        GreenwichYear   = 0
        GreenwichMonth  = 0
        GreenwichDay    = $GmtDay
        GreenwichHour   = $GmtHour
        GreenwichMinute = $GmtMinute
        GreenwichSecond = $GmtSecond
        DecimalDay      = $GmtDecimalDay
        DecimalHour     = $GmtDecimalHour
        T               = $T                 # Don't include T/T2 - (In each function: $T=$JDN/36525; $T2=$T*$T (JDN already includes hour))
        T2              = $T2
    }
}


Function Moon {
    [OutputType([Double])]
    Param ($JDN)

    $SunLong  = SunLong $JDN
    $SunMeanAnomaly = SunMeanAnomaly $JDN

    $MoonLong = MoonLong $JDN
    $MoonLat  = MoonLat $JDN
    $MoonMeanAnomaly = MoonMeanAnomaly $JDN

    $CD = [Math]::Cos((Radians ($MoonLong - $SunLong))) * [Math]::Cos((Radians ($MoonLat)))
    $D = [Math]::Acos($CD)
    $SD = [Math]::Sin($D)
        
    $I = 0.1468 * $SD * (1 - 0.0549 * [Math]::Sin($MoonMeanAnomaly))
    $I = $I / (1 - 0.0167 * [Math]::Sin($SunMeanAnomaly))
    $I = [Math]::Pi - $D - (Radians $I)   # Age of the moon (0-Pi Radians)
    $AgeDays = $I * 29.53058868 / [Math]::Pi
    $K = (1 + [Math]::Cos($I)) / 2    # Convert age into phase (0-1)

        
    [PsCustomObject]@{
        PSTypeName = 'MoonInfo'
        AgeDays    = $AgeDays
        AgeDegrees = Degrees (2 * $I)
        Phase      = $K
    }
}





Function MoonPos {
    [OutputType([Double])]
    Param ($JDN)

    $Q = $JDN - 2415020    # Days since midday 0m 0s 1900/1/1 
    $T = $Q / 36525   # Convert to centuries 
    $T2 = $T * $T
    
    # http://en.wikipedia.org/wiki/Month
    $M1 = 27.32158213    # Tropical Month, moon returning to same poistion wrt vernal equinox
    $M2 = 365.2596407    # Anomalistic Year - one revolution between successive apsides (perihelion/aphelion)
    $M3 = 27.55455094    # Anomalistic Month - time between successive perigees
    $M4 = 29.53058868    # Synodic Period, time between successive new moons
    $M5 = 27.21222039    # Draconic Month, time for moon to return to the same node (crossing the plane of the earth's orbit)
    $M6 = 6798.363307    # period of the sidereal revolution of the moon's node (~223 revolutions/18 years)
        
    $M1 = Circle ($Q / $M1)    # Relative position in orbit (tropical month) based on known position at epoch
    $M2 = Circle ($Q / $M2)
    $M3 = Circle ($Q / $M3)
    $M4 = Circle ($Q / $M4)
    $M5 = Circle ($Q / $M5)
    $M6 = Circle ($Q / $M6)

    $ML = 270.434164 + $M1 - (0.001133 - 0.0000019 * $T) * $T2
    $MS = 358.475833 + $M2 - (0.00015 + 0.0000033 * $T) * $T2
    $MD = 296.104608 + $M3 + (0.009192 + 0.0000144 * $T) * $T2
    $ME1 = 350.737486 + $M4 - (0.001436 - 0.0000019 * $T) * $T2
    $MF = 11.250889 + $M5 - (0.003211 + 0.0000003 * $T) * $T2
    $NA = 259.183275 - $M6 + (0.002078 + 0.0000022 * $T) * $T2

    $A = Radians (51.2 + 20.2 * $T)
    $S1 = [Math]::Sin($A)
    $S2 = [Math]::Sin((Radians $NA))
    $B = 346.56 + (132.87 - 0.0091731 * $T) * $T
    $S3 = 0.003964 * [Math]::Sin((Radians $B))
    $C = Radians ($NA + 275.05 - 2.3 * $T)
    $S4 = [Math]::Sin($C)

    $ML = $ML + 0.000233 * $S1 + $S3 + 0.001964 * $S2
    $MS = $MS - 0.001778 * $S1
    $MD = $MD + 0.000817 * $S1 + $S3 + 0.002541 * $S2

    $MD = $MD + 0.000817 * $S1 + $S3 + 0.002541 * $S2     # $MD = Moon Mean Anomaly


    # Calculate Horizontal Parallax

    $MF = $MF + $S3 - 0.024691 * $S2 - 0.004328 * $S4
    $ME1 = $ME1 + 0.002011 * $S1 + $S3 + 0.001964 * $S2
    $E = 1 - (0.002495 + 0.00000752 * $T) * $T
    $E2 = $E * $E
        
            
    $ML  = Radians $ML    # Longitude
    $MS  = Radians $MS   
    $NA  = Radians $NA   
    $ME1 = Radians $ME1
    $MF  = Radians $MF
    $MD  = Radians $MD
   

    $PM = 0.950724 + 0.051818 * [Math]::Cos($MD) + 0.009531 * [Math]::Cos(2 * $ME1 - $MD)
    $PM = $PM + 0.007843 * [Math]::Cos(2 * $ME1) + 0.002824 * [Math]::Cos(2 * $MD)
    $PM = $PM + 0.000857 * [Math]::Cos(2 * $ME1 + $MD) + $E * 0.000533 * [Math]::Cos(2 * $ME1 - $MS)
    $PM = $PM + $E * 0.000401 * [Math]::Cos(2 * $ME1 - $MD - $MS)
    $PM = $PM + $E * 0.00032 * [Math]::Cos($MD - $MS) - 0.000271 * [Math]::Cos($ME1)
    $PM = $PM - $E * 0.000264 * [Math]::Cos($MS + $MD) - 0.000198 * [Math]::Cos(2 * $MF - $MD)
    $PM = $PM + 0.000173 * [Math]::Cos(3 * $MD) + 0.000167 * [Math]::Cos(4 * $ME1 - $MD)
    $PM = $PM - $E * 0.000111 * [Math]::Cos($MS) + 0.000103 * [Math]::Cos(4 * $ME1 - 2 * $MD)
    $PM = $PM - 0.000084 * [Math]::Cos(2 * $MD - 2 * $ME1) - $E * 0.000083 * [Math]::Cos(2 * $ME1 + $MS)
    $PM = $PM + 0.000079 * [Math]::Cos(2 * $ME1 + 2 * $MD) + 0.000072 * [Math]::Cos(4 * $ME1)
    $PM = $PM + $E * 0.000064 * [Math]::Cos(2 * $ME1 - $MS + $MD) - $E * 0.000063 * [Math]::Cos(2 * $ME1 + $MS - $MD)
    $PM = $PM + $E * 0.000041 * [Math]::Cos($MS + $ME1) + $E * 0.000035 * [Math]::Cos(2 * $MD - $MS)
    $PM = $PM - 0.000033 * [Math]::Cos(3 * $MD - 2 * $ME1) - 0.00003 * [Math]::Cos($MD + $ME1)
    $PM = $PM - 0.000029 * [Math]::Cos(2 * ($MF - $ME1)) - $E * 0.000029 * [Math]::Cos(2 * $MD + $MS)
    $PM = $PM + $E2 * 0.000026 * [Math]::Cos(2 * ($ME1 - $MS)) - 0.000023 * [Math]::Cos(2 * ($MF - $ME1) + $MD)
    $PM = $PM + $E * 0.000019 * [Math]::Cos(4 * $ME1 - $MS - $MD)   # Horizontal Parallax
        

    # Calculate Moon Latitude

    $G = 5.128189 * [Math]::Sin($MF) + 0.280606 * [Math]::Sin($MD + $MF)
    $G = $G + 0.277693 * [Math]::Sin($MD - $MF) + 0.173238 * [Math]::Sin(2 * $ME1 - $MF)
    $G = $G + 0.055413 * [Math]::Sin(2 * $ME1 + $MF - $MD) + 0.046272 * [Math]::Sin(2 * $ME1 - $MF - $MD)
    $G = $G + 0.032573 * [Math]::Sin(2 * $ME1 + $MF) + 0.017198 * [Math]::Sin(2 * $MD + $MF)
    $G = $G + 0.009267 * [Math]::Sin(2 * $ME1 + $MD - $MF) + 0.008823 * [Math]::Sin(2 * $MD - $MF)
    $G = $G + $E * 0.008247 * [Math]::Sin(2 * $ME1 - $MS - $MF) + 0.004323 * [Math]::Sin(2 * ($ME1 - $MD) - $MF)
    $G = $G + 0.0042 * [Math]::Sin(2 * $ME1 + $MF + $MD) + $E * 0.003372 * [Math]::Sin($MF - $MS - 2 * $ME1)
    $G = $G + $E * 0.002472 * [Math]::Sin(2 * $ME1 + $MF - $MS - $MD)
    $G = $G + $E * 0.002222 * [Math]::Sin(2 * $ME1 + $MF - $MS)
    $G = $G + $E * 0.002072 * [Math]::Sin(2 * $ME1 - $MF - $MS - $MD)
    $G = $G + $E * 0.001877 * [Math]::Sin($MF - $MS + $MD) + 0.001828 * [Math]::Sin(4 * $ME1 - $MF - $MD)
    $G = $G - $E * 0.001803 * [Math]::Sin($MF + $MS) - 0.00175 * [Math]::Sin(3 * $MF)
    $G = $G + $E * 0.00157 * [Math]::Sin($MD - $MS - $MF) - 0.001487 * [Math]::Sin($MF + $ME1)
    $G = $G - $E * 0.001481 * [Math]::Sin($MF + $MS + $MD) + $E * 0.001417 * [Math]::Sin($MF - $MS - $MD)
    $G = $G + $E * 0.00135 * [Math]::Sin($MF - $MS) + 0.00133 * [Math]::Sin($MF - $ME1)
    $G = $G + 0.001106 * [Math]::Sin($MF + 3 * $MD) + 0.00102 * [Math]::Sin(4 * $ME1 - $MF)
    $G = $G + 0.000833 * [Math]::Sin($MF + 4 * $ME1 - $MD) + 0.000781 * [Math]::Sin($MD - 3 * $MF)
    $G = $G + 0.00067 * [Math]::Sin($MF + 4 * $ME1 - 2 * $MD) + 0.000606 * [Math]::Sin(2 * $ME1 - 3 * $MF)
    $G = $G + 0.000597 * [Math]::Sin(2 * ($ME1 + $MD) - $MF)
    $G = $G + $E * 0.000492 * [Math]::Sin(2 * $ME1 + $MD - $MS - $MF) + 0.00045 * [Math]::Sin(2 * ($MD - $ME1) - $MF)
    $G = $G + 0.000439 * [Math]::Sin(3 * $MD - $MF) + 0.000423 * [Math]::Sin($MF + 2 * ($ME1 + $MD))
    $G = $G + 0.000422 * [Math]::Sin(2 * $ME1 - $MF - 3 * $MD) - $E * 0.000367 * [Math]::Sin($MS + $MF + 2 * $ME1 - $MD)
    $G = $G - $E * 0.000353 * [Math]::Sin($MS + $MF + 2 * $ME1) + 0.000331 * [Math]::Sin($MF + 4 * $ME1)
    $G = $G + $E * 0.000317 * [Math]::Sin(2 * $ME1 + $MF - $MS + $MD)
    $G = $G + $E2 * 0.000306 * [Math]::Sin(2 * ($ME1 - $MS) - $MF) - 0.000283 * [Math]::Sin($MD + 3 * $MF)
    
    $W1 = 0.0004664 * [Math]::Cos($NA)
    $W2 = 0.0000754 * [Math]::Cos($C)
    
    $BM = Radians($G) * (1 - $W1 - $W2)   # Latitude


    # Calculate Moon Longitude

    $L = 6.28875 * [Math]::Sin($MD) + 1.274018 * [Math]::Sin(2 * $ME1 - $MD)

    # Add further Corrections
    $L = $L + 0.658309 * [Math]::Sin(2 * $ME1) + 0.213616 * [Math]::Sin(2 * $MD)
    $L = $L - $E * 0.185596 * [Math]::Sin($MS) - 0.114336 * [Math]::Sin(2 * $MF)
    $L = $L + 0.058793 * [Math]::Sin(2 * ($ME1 - $MD))
    $L = $L + 0.057212 * $E * [Math]::Sin(2 * $ME1 - $MS - $MD) + 0.05332 * [Math]::Sin(2 * $ME1 + $MD)
    $L = $L + 0.045874 * $E * [Math]::Sin(2 * $ME1 - $MS) + 0.041024 * $E * [Math]::Sin($MD - $MS)
    $L = $L - 0.034718 * [Math]::Sin($ME1) - $E * 0.030465 * [Math]::Sin($MS + $MD)
    $L = $L + 0.015326 * [Math]::Sin(2 * ($ME1 - $MF)) - 0.012528 * [Math]::Sin(2 * $MF + $MD)
    $L = $L - 0.01098 * [Math]::Sin(2 * $MF - $MD) + 0.010674 * [Math]::Sin(4 * $ME1 - $MD)
    $L = $L + 0.010034 * [Math]::Sin(3 * $MD) + 0.008548 * [Math]::Sin(4 * $ME1 - 2 * $MD)
    $L = $L - $E * 0.00791 * [Math]::Sin($MS - $MD + 2 * $ME1) - $E * 0.006783 * [Math]::Sin(2 * $ME1 + $MS)
    $L = $L + 0.005162 * [Math]::Sin($MD - $ME1) + $E * 0.005 * [Math]::Sin($MS + $ME1)
    $L = $L + 0.003862 * [Math]::Sin(4 * $ME1) + $E * 0.004049 * [Math]::Sin($MD - $MS + 2 * $ME1)
    $L = $L + 0.003996 * [Math]::Sin(2 * ($MD + $ME1)) + 0.003665 * [Math]::Sin(2 * $ME1 - 3 * $MD)
    $L = $L + $E * 0.002695 * [Math]::Sin(2 * $MD - $MS) + 0.002602 * [Math]::Sin($MD - 2 * ($MF + $ME1))
    $L = $L + $E * 0.002396 * [Math]::Sin(2 * ($ME1 - $MD) - $MS) - 0.002349 * [Math]::Sin($MD + $ME1)
    $L = $L + $E2 * 0.002249 * [Math]::Sin(2 * ($ME1 - $MS)) - $E * 0.002125 * [Math]::Sin(2 * $MD + $MS)
    $L = $L - $E2 * 0.002079 * [Math]::Sin(2 * $MS) + $E2 * 0.002059 * [Math]::Sin(2 * ($ME1 - $MS) - $MD)
    $L = $L - 0.001773 * [Math]::Sin($MD + 2 * ($ME1 - $MF)) - 0.001595 * [Math]::Sin(2 * ($MF + $ME1))
    $L = $L + $E * 0.00122 * [Math]::Sin(4 * $ME1 - $MS - $MD) - 0.00111 * [Math]::Sin(2 * ($MD + $MF))
    $L = $L + 0.000892 * [Math]::Sin($MD - 3 * $ME1) - $E * 0.000811 * [Math]::Sin($MS + $MD + 2 * $ME1)
    $L = $L + $E * 0.000761 * [Math]::Sin(4 * $ME1 - $MS - 2 * $MD)
    $L = $L + $E2 * 0.000704 * [Math]::Sin($MD - 2 * ($MS + $ME1))
    $L = $L + $E * 0.000693 * [Math]::Sin($MS - 2 * ($MD - $ME1))
    $L = $L + $E * 0.000598 * [Math]::Sin(2 * ($ME1 - $MF) - $MS)
    $L = $L + 0.00055 * [Math]::Sin($MD + 4 * $ME1) + 0.000538 * [Math]::Sin(4 * $MD)
    $L = $L + $E * 0.000521 * [Math]::Sin(4 * $ME1 - $MS) + 0.000486 * [Math]::Sin(2 * $MD - $ME1)
    $L = $L + $E2 * 0.000717 * [Math]::Sin($MD - 2 * $MS)
        
    $Long = (NormaliseDegrees ($ML + $L))   # Longitude



    [PsCustomObject] @{
        PsTypeName         = 'MoonPos'
        MeanAnomaly        = (NormaliseRadians (Radians $MD))
        HorizontalParallax = $PM
        Latitude           = (NormaliseDegrees (Degrees $BM))    # Degrees
        Longitude          = $Long                               # Degrees
    }
}




Function MoonLat {
    [OutputType([Double])]
    Param ($JDN)

    $Q = $JDN - 2415020    # Days since midday 0m 0s 1900/1/1 
    $T = $Q / 36525   # Convert to centuries 
    $T2 = $T * $T

    
    # http://en.wikipedia.org/wiki/Month
    $M1 = 27.32158213    # Tropical Month, moon returning to same poistion wrt vernal equinox
    $M2 = 365.2596407    # Anomalistic Year - one revolution between successive apsides (perihelion/aphelion)
    $M3 = 27.55455094    # Anomalistic Month - time between successive perigees
    $M4 = 29.53058868    # Synodic Period, time between successive full moons
    $M5 = 27.21222039    # Draconic (Nodal) Month, time for moon to return to the same node (crossing the plane of the earth's orbit)
    $M6 = 6798.363307    # period of the sidereal revolution of the moon's nodes (~223 revolutions/18.61 years)
        
    $M1 = Circle ($Q / $M1)    # Relative position in orbit (tropical month) based on known position at epoch
    $M2 = Circle ($Q / $M2)
    $M3 = Circle ($Q / $M3)
    $M4 = Circle ($Q / $M4)
    $M5 = Circle ($Q / $M5)
    $M6 = Circle ($Q / $M6)

    $ML = 270.434164 + $M1 - (0.001133 - 0.0000019 * $T) * $T2  # Sun's Ecliptic Longitude?
    $MS = 358.475833 + $M2 - (0.00015 + 0.0000033 * $T) * $T2   # Sun Mean Anomaly?
    $MD = 296.104608 + $M3 + (0.009192 + 0.0000144 * $T) * $T2  
    $ME1 = 350.737486 + $M4 - (0.001436 - 0.0000019 * $T) * $T2
    $MF = 11.250889 + $M5 - (0.003211 + 0.0000003 * $T) * $T2
    $NA = 259.183275 - $M6 + (0.002078 + 0.0000022 * $T) * $T2

    $A = Radians (51.2 + 20.2 * $T)
    $S1 = [Math]::Sin($A)
    $S2 = [Math]::Sin((Radians $NA))
    $B = 346.56 + (132.87 - 0.0091731 * $T) * $T
    $S3 = 0.003964 * [Math]::Sin((Radians $B))
    $C = Radians ($NA + 275.05 - 2.3 * $T)
    $S4 = [Math]::Sin($C)
    $ML = $ML + 0.000233 * $S1 + $S3 + 0.001964 * $S2
    $MS = $MS - 0.001778 * $S1
    $MD = $MD + 0.000817 * $S1 + $S3 + 0.002541 * $S2     # $MD = Moon Mean Anomaly

    $MF = $MF + $S3 - 0.024691 * $S2 - 0.004328 * $S4
    $ME1 = $ME1 + 0.002011 * $S1 + $S3 + 0.001964 * $S2
    $E = 1 - (0.002495 + 0.00000752 * $T) * $T
    $E2 = $E * $E
        
            
    # $ML  = Radians $ML    # Not used
    $MS  = Radians $MS   
    $NA  = Radians $NA
    $ME1 = Radians $ME1
    $MF  = Radians $MF
    $MD  = Radians $MD


    $G = 5.128189 * [Math]::Sin($MF) + 0.280606 * [Math]::Sin($MD + $MF)
    $G = $G + 0.277693 * [Math]::Sin($MD - $MF) + 0.173238 * [Math]::Sin(2 * $ME1 - $MF)
    $G = $G + 0.055413 * [Math]::Sin(2 * $ME1 + $MF - $MD) + 0.046272 * [Math]::Sin(2 * $ME1 - $MF - $MD)
    $G = $G + 0.032573 * [Math]::Sin(2 * $ME1 + $MF) + 0.017198 * [Math]::Sin(2 * $MD + $MF)
    $G = $G + 0.009267 * [Math]::Sin(2 * $ME1 + $MD - $MF) + 0.008823 * [Math]::Sin(2 * $MD - $MF)
    $G = $G + $E * 0.008247 * [Math]::Sin(2 * $ME1 - $MS - $MF) + 0.004323 * [Math]::Sin(2 * ($ME1 - $MD) - $MF)
    $G = $G + 0.0042 * [Math]::Sin(2 * $ME1 + $MF + $MD) + $E * 0.003372 * [Math]::Sin($MF - $MS - 2 * $ME1)
    $G = $G + $E * 0.002472 * [Math]::Sin(2 * $ME1 + $MF - $MS - $MD)
    $G = $G + $E * 0.002222 * [Math]::Sin(2 * $ME1 + $MF - $MS)
    $G = $G + $E * 0.002072 * [Math]::Sin(2 * $ME1 - $MF - $MS - $MD)
    $G = $G + $E * 0.001877 * [Math]::Sin($MF - $MS + $MD) + 0.001828 * [Math]::Sin(4 * $ME1 - $MF - $MD)
    $G = $G - $E * 0.001803 * [Math]::Sin($MF + $MS) - 0.00175 * [Math]::Sin(3 * $MF)
    $G = $G + $E * 0.00157 * [Math]::Sin($MD - $MS - $MF) - 0.001487 * [Math]::Sin($MF + $ME1)
    $G = $G - $E * 0.001481 * [Math]::Sin($MF + $MS + $MD) + $E * 0.001417 * [Math]::Sin($MF - $MS - $MD)
    $G = $G + $E * 0.00135 * [Math]::Sin($MF - $MS) + 0.00133 * [Math]::Sin($MF - $ME1)
    $G = $G + 0.001106 * [Math]::Sin($MF + 3 * $MD) + 0.00102 * [Math]::Sin(4 * $ME1 - $MF)
    $G = $G + 0.000833 * [Math]::Sin($MF + 4 * $ME1 - $MD) + 0.000781 * [Math]::Sin($MD - 3 * $MF)
    $G = $G + 0.00067 * [Math]::Sin($MF + 4 * $ME1 - 2 * $MD) + 0.000606 * [Math]::Sin(2 * $ME1 - 3 * $MF)
    $G = $G + 0.000597 * [Math]::Sin(2 * ($ME1 + $MD) - $MF)
    $G = $G + $E * 0.000492 * [Math]::Sin(2 * $ME1 + $MD - $MS - $MF) + 0.00045 * [Math]::Sin(2 * ($MD - $ME1) - $MF)
    $G = $G + 0.000439 * [Math]::Sin(3 * $MD - $MF) + 0.000423 * [Math]::Sin($MF + 2 * ($ME1 + $MD))
    $G = $G + 0.000422 * [Math]::Sin(2 * $ME1 - $MF - 3 * $MD) - $E * 0.000367 * [Math]::Sin($MS + $MF + 2 * $ME1 - $MD)
    $G = $G - $E * 0.000353 * [Math]::Sin($MS + $MF + 2 * $ME1) + 0.000331 * [Math]::Sin($MF + 4 * $ME1)
    $G = $G + $E * 0.000317 * [Math]::Sin(2 * $ME1 + $MF - $MS + $MD)
    $G = $G + $E2 * 0.000306 * [Math]::Sin(2 * ($ME1 - $MS) - $MF) - 0.000283 * [Math]::Sin($MD + 3 * $MF)
    
    $W1 = 0.0004664 * [Math]::Cos($NA)
    $W2 = 0.0000754 * [Math]::Cos($C)
    
    $BM = Radians($G) * (1 - $W1 - $W2)

    Return (NormaliseDegrees (Degrees $BM))
}


Function MoonLong {
    [OutputType([Double])]
    Param ($JDN)

    $Q = $JDN - 2415020    # Days since midday 0m 0s 1900/1/1 
    $T = $Q / 36525   # Convert to centuries 
    $T2 = $T * $T

    # http://en.wikipedia.org/wiki/Month
    $M1 = 27.32158213    # Tropical Month, moon returning to same poistion wrt vernal equinox
    $M2 = 365.2596407    # Anomalistic Year - one revolution between successive apsides (perihelion/aphelion)
    $M3 = 27.55455094    # Anomalistic Month - time between successive perigees
    $M4 = 29.53058868    # Synodic Period, time between successive full moons
    $M5 = 27.21222039    # Draconic Month, time for moon to return to the same node (crossing the plane of the earth's orbit)
    $M6 = 6798.363307    # period of the sidereal revolution of the moon's node (~223 revolutions/18 years)
        
    $M1 = Circle ($Q / $M1)    # Relative position in orbit (tropical month) based on known position at epoch
    $M2 = Circle ($Q / $M2)
    $M3 = Circle ($Q / $M3)
    $M4 = Circle ($Q / $M4)
    $M5 = Circle ($Q / $M5)
    $M6 = Circle ($Q / $M6)

    $ML = 270.434164 + $M1 - (0.001133 - 0.0000019 * $T) * $T2
    $MS = 358.475833 + $M2 - (0.00015 + 0.0000033 * $T) * $T2
    $MD = 296.104608 + $M3 + (0.009192 + 0.0000144 * $T) * $T2
    $ME1 = 350.737486 + $M4 - (0.001436 - 0.0000019 * $T) * $T2
    $MF = 11.250889 + $M5 - (0.003211 + 0.0000003 * $T) * $T2
    $NA = 259.183275 - $M6 + (0.002078 + 0.0000022 * $T) * $T2

    $A = Radians (51.2 + 20.2 * $T)
    $S1 = [Math]::Sin($A)
    $S2 = [Math]::Sin((Radians $NA))
    $B = 346.56 + (132.87 - 0.0091731 * $T) * $T
    $S3 = 0.003964 * [Math]::Sin((Radians $B))
    $C = Radians ($NA + 275.05 - 2.3 * $T)
    $S4 = [Math]::Sin($C)
    $ML = $ML + 0.000233 * $S1 + $S3 + 0.001964 * $S2
    $MS = $MS - 0.001778 * $S1
    $MD = $MD + 0.000817 * $S1 + $S3 + 0.002541 * $S2
    $MF = $MF + $S3 - 0.024691 * $S2 - 0.004328 * $S4
    $ME1 = $ME1 + 0.002011 * $S1 + $S3 + 0.001964 * $S2
    $E = 1 - (0.002495 + 0.00000752 * $T) * $T
    $E2 = $E * $E
        
    # $ML  = Radians $ML    # Leave this value in degrees
    $MS  = Radians $MS
    $NA  = Radians $NA
    $ME1 = Radians $ME1
    $MF  = Radians $MF
    $MD  = Radians $MD

    $L = 6.28875 * [Math]::Sin($MD) + 1.274018 * [Math]::Sin(2 * $ME1 - $MD)

    # Add in further Corrections
    $L = $L + 0.658309 * [Math]::Sin(2 * $ME1) + 0.213616 * [Math]::Sin(2 * $MD)
    $L = $L - $E * 0.185596 * [Math]::Sin($MS) - 0.114336 * [Math]::Sin(2 * $MF)
    $L = $L + 0.058793 * [Math]::Sin(2 * ($ME1 - $MD))
    $L = $L + 0.057212 * $E * [Math]::Sin(2 * $ME1 - $MS - $MD) + 0.05332 * [Math]::Sin(2 * $ME1 + $MD)
    $L = $L + 0.045874 * $E * [Math]::Sin(2 * $ME1 - $MS) + 0.041024 * $E * [Math]::Sin($MD - $MS)
    $L = $L - 0.034718 * [Math]::Sin($ME1) - $E * 0.030465 * [Math]::Sin($MS + $MD)
    $L = $L + 0.015326 * [Math]::Sin(2 * ($ME1 - $MF)) - 0.012528 * [Math]::Sin(2 * $MF + $MD)
    $L = $L - 0.01098 * [Math]::Sin(2 * $MF - $MD) + 0.010674 * [Math]::Sin(4 * $ME1 - $MD)
    $L = $L + 0.010034 * [Math]::Sin(3 * $MD) + 0.008548 * [Math]::Sin(4 * $ME1 - 2 * $MD)
    $L = $L - $E * 0.00791 * [Math]::Sin($MS - $MD + 2 * $ME1) - $E * 0.006783 * [Math]::Sin(2 * $ME1 + $MS)
    $L = $L + 0.005162 * [Math]::Sin($MD - $ME1) + $E * 0.005 * [Math]::Sin($MS + $ME1)
    $L = $L + 0.003862 * [Math]::Sin(4 * $ME1) + $E * 0.004049 * [Math]::Sin($MD - $MS + 2 * $ME1)
    $L = $L + 0.003996 * [Math]::Sin(2 * ($MD + $ME1)) + 0.003665 * [Math]::Sin(2 * $ME1 - 3 * $MD)
    $L = $L + $E * 0.002695 * [Math]::Sin(2 * $MD - $MS) + 0.002602 * [Math]::Sin($MD - 2 * ($MF + $ME1))
    $L = $L + $E * 0.002396 * [Math]::Sin(2 * ($ME1 - $MD) - $MS) - 0.002349 * [Math]::Sin($MD + $ME1)
    $L = $L + $E2 * 0.002249 * [Math]::Sin(2 * ($ME1 - $MS)) - $E * 0.002125 * [Math]::Sin(2 * $MD + $MS)
    $L = $L - $E2 * 0.002079 * [Math]::Sin(2 * $MS) + $E2 * 0.002059 * [Math]::Sin(2 * ($ME1 - $MS) - $MD)
    $L = $L - 0.001773 * [Math]::Sin($MD + 2 * ($ME1 - $MF)) - 0.001595 * [Math]::Sin(2 * ($MF + $ME1))
    $L = $L + $E * 0.00122 * [Math]::Sin(4 * $ME1 - $MS - $MD) - 0.00111 * [Math]::Sin(2 * ($MD + $MF))
    $L = $L + 0.000892 * [Math]::Sin($MD - 3 * $ME1) - $E * 0.000811 * [Math]::Sin($MS + $MD + 2 * $ME1)
    $L = $L + $E * 0.000761 * [Math]::Sin(4 * $ME1 - $MS - 2 * $MD)
    $L = $L + $E2 * 0.000704 * [Math]::Sin($MD - 2 * ($MS + $ME1))
    $L = $L + $E * 0.000693 * [Math]::Sin($MS - 2 * ($MD - $ME1))
    $L = $L + $E * 0.000598 * [Math]::Sin(2 * ($ME1 - $MF) - $MS)
    $L = $L + 0.00055 * [Math]::Sin($MD + 4 * $ME1) + 0.000538 * [Math]::Sin(4 * $MD)
    $L = $L + $E * 0.000521 * [Math]::Sin(4 * $ME1 - $MS) + 0.000486 * [Math]::Sin(2 * $MD - $ME1)
    $L = $L + $E2 * 0.000717 * [Math]::Sin($MD - 2 * $MS)
        
    # Return (Degrees (NormaliseRadians ($ML + (Radians $L))))

    Return (NormaliseDegrees ($ML + $L))

}


Function MoonMeanAnomaly {
    [OutputType([Double])]
    Param ($JDN)

    $Q = $JDN - 2415020    # Days since midday 0m 0s 1900/1/1 
    $T = $Q / 36525   # Convert to centuries 
    $T2 = $T * $T
    
    # http://en.wikipedia.org/wiki/Month
    $M1 = 27.32158213    # Tropical Month, moon returning to same poistion wrt vernal equinox
    $M2 = 365.2596407    # Anomalistic Year - one revolution between successive apsides (perihelion/aphelion)
    $M3 = 27.55455094    # Anomalistic Month - time between successive perigees
    $M4 = 29.53058868    # Synodic Period, time between successive new moons
    $M5 = 27.21222039    # Draconic Month, time for moon to return to the same node (crossing the plane of the earth's orbit)
    $M6 = 6798.363307    # period of the sidereal revolution of the moon's node (~223 revolutions/18 years)
        
    $M1 = Circle ($Q / $M1)    # Relative position in orbit (tropical month) based on known position at epoch
    $M2 = Circle ($Q / $M2)
    $M3 = Circle ($Q / $M3)
    $M4 = Circle ($Q / $M4)
    $M5 = Circle ($Q / $M5)
    $M6 = Circle ($Q / $M6)

    $ML = 270.434164 + $M1 - (0.001133 - 0.0000019 * $T) * $T2
    $MS = 358.475833 + $M2 - (0.00015 + 0.0000033 * $T) * $T2
    $MD = 296.104608 + $M3 + (0.009192 + 0.0000144 * $T) * $T2
    $ME1 = 350.737486 + $M4 - (0.001436 - 0.0000019 * $T) * $T2
    $MF = 11.250889 + $M5 - (0.003211 + 0.0000003 * $T) * $T2
    $NA = 259.183275 - $M6 + (0.002078 + 0.0000022 * $T) * $T2

    $A = Radians (51.2 + 20.2 * $T)
    $S1 = [Math]::Sin($A)
    $S2 = [Math]::Sin((Radians $NA))
    $B = 346.56 + (132.87 - 0.0091731 * $T) * $T
    $S3 = 0.003964 * [Math]::Sin((Radians $B))
    $C = Radians ($NA + 275.05 - 2.3 * $T)
    $S4 = [Math]::Sin($C)

    $ML = $ML + 0.000233 * $S1 + $S3 + 0.001964 * $S2
    $MS = $MS - 0.001778 * $S1
    $MD = $MD + 0.000817 * $S1 + $S3 + 0.002541 * $S2

    Return (NormaliseRadians (Radians $MD))
}


Function MoonHP {
# Moon Horizontal Parallax
    [OutputType([Double])]
    Param ($JDN)

    $Q = $JDN - 2415020    # Days since midday 0m 0s 1900/1/1 
    $T = $Q / 36525   # Convert to centuries 
    $T2 = $T * $T

    
    # http://en.wikipedia.org/wiki/Month
    $M1 = 27.32158213    # Tropical Month, moon returning to same poistion wrt vernal equinox
    $M2 = 365.2596407    # Anomalistic Year - one revolution between successive apsides (perihelion/aphelion)
    $M3 = 27.55455094    # Anomalistic Month - time between successive perigees
    $M4 = 29.53058868    # Synodic Period, time between successive new moons
    $M5 = 27.21222039    # Draconic Month, time for moon to return to the same node (crossing the plane of the earth's orbit)
    $M6 = 6798.363307    # period of the sidereal revolution of the moon's node (~223 revolutions/18 years)
        
    $M1 = Circle ($Q / $M1)    # Relative position in orbit (tropical month) based on known position at epoch
    $M2 = Circle ($Q / $M2)
    $M3 = Circle ($Q / $M3)
    $M4 = Circle ($Q / $M4)
    $M5 = Circle ($Q / $M5)
    $M6 = Circle ($Q / $M6)

    $ML = 270.434164 + $M1 - (0.001133 - 0.0000019 * $T) * $T2
    $MS = 358.475833 + $M2 - (0.00015 + 0.0000033 * $T) * $T2
    $MD = 296.104608 + $M3 + (0.009192 + 0.0000144 * $T) * $T2
    $ME1 = 350.737486 + $M4 - (0.001436 - 0.0000019 * $T) * $T2
    $MF = 11.250889 + $M5 - (0.003211 + 0.0000003 * $T) * $T2
    $NA = 259.183275 - $M6 + (0.002078 + 0.0000022 * $T) * $T2

    $A = Radians (51.2 + 20.2 * $T)
    $S1 = [Math]::Sin($A)
    $S2 = [Math]::Sin((Radians $NA))
    $B = 346.56 + (132.87 - 0.0091731 * $T) * $T
    $S3 = 0.003964 * [Math]::Sin((Radians $B))
    $C = Radians ($NA + 275.05 - 2.3 * $T)
    $S4 = [Math]::Sin($C)
    $ML = $ML + 0.000233 * $S1 + $S3 + 0.001964 * $S2
    $MS = $MS - 0.001778 * $S1
    $MD = $MD + 0.000817 * $S1 + $S3 + 0.002541 * $S2     # $MD = Moon Mean Anomaly

    $MF = $MF + $S3 - 0.024691 * $S2 - 0.004328 * $S4
    $ME1 = $ME1 + 0.002011 * $S1 + $S3 + 0.001964 * $S2
    $E = 1 - (0.002495 + 0.00000752 * $T) * $T
    $E2 = $E * $E
        
            
    # $ML  = Radians $ML    # Not used
    $MS  = Radians $MS   
    # $NA  = Radians $NA    # Not used
    $ME1 = Radians $ME1
    $MF  = Radians $MF
    $MD  = Radians $MD
   

    $PM = 0.950724 + 0.051818 * [Math]::Cos($MD) + 0.009531 * [Math]::Cos(2 * $ME1 - $MD)
    $PM = $PM + 0.007843 * [Math]::Cos(2 * $ME1) + 0.002824 * [Math]::Cos(2 * $MD)
    $PM = $PM + 0.000857 * [Math]::Cos(2 * $ME1 + $MD) + $E * 0.000533 * [Math]::Cos(2 * $ME1 - $MS)
    $PM = $PM + $E * 0.000401 * [Math]::Cos(2 * $ME1 - $MD - $MS)
    $PM = $PM + $E * 0.00032 * [Math]::Cos($MD - $MS) - 0.000271 * [Math]::Cos($ME1)
    $PM = $PM - $E * 0.000264 * [Math]::Cos($MS + $MD) - 0.000198 * [Math]::Cos(2 * $MF - $MD)
    $PM = $PM + 0.000173 * [Math]::Cos(3 * $MD) + 0.000167 * [Math]::Cos(4 * $ME1 - $MD)
    $PM = $PM - $E * 0.000111 * [Math]::Cos($MS) + 0.000103 * [Math]::Cos(4 * $ME1 - 2 * $MD)
    $PM = $PM - 0.000084 * [Math]::Cos(2 * $MD - 2 * $ME1) - $E * 0.000083 * [Math]::Cos(2 * $ME1 + $MS)
    $PM = $PM + 0.000079 * [Math]::Cos(2 * $ME1 + 2 * $MD) + 0.000072 * [Math]::Cos(4 * $ME1)
    $PM = $PM + $E * 0.000064 * [Math]::Cos(2 * $ME1 - $MS + $MD) - $E * 0.000063 * [Math]::Cos(2 * $ME1 + $MS - $MD)
    $PM = $PM + $E * 0.000041 * [Math]::Cos($MS + $ME1) + $E * 0.000035 * [Math]::Cos(2 * $MD - $MS)
    $PM = $PM - 0.000033 * [Math]::Cos(3 * $MD - 2 * $ME1) - 0.00003 * [Math]::Cos($MD + $ME1)
    $PM = $PM - 0.000029 * [Math]::Cos(2 * ($MF - $ME1)) - $E * 0.000029 * [Math]::Cos(2 * $MD + $MS)
    $PM = $PM + $E2 * 0.000026 * [Math]::Cos(2 * ($ME1 - $MS)) - 0.000023 * [Math]::Cos(2 * ($MF - $ME1) + $MD)
    $PM = $PM + $E * 0.000019 * [Math]::Cos(4 * $ME1 - $MS - $MD)
        
    Return $PM
}


Function ECRA {
# Ecliptic Coordinates to Right Ascention
# (ELD As Double, ELM As Double, ELS As Double,   # DMS Longitude
#   BD As Double,  BM As Double, BS As Double,    # DMS Latitude
#   GD As Double,  GM As Double, GY As Double)    # ddMMYY (to calculate Obliquity of the Ecliptic)
Param ($ELD, $ELM, $ELS,
       $BD, $BM, $BS,
       $GD, $GM, $GY
      )
       
    $A = Radians (DMSDD $ELD $ELM $ELS)
    $B = Radians (DMSDD $BD $BM $BS)
    $C = Radians (Obliq $GD $GM $GY)
    $D = [Math]::Sin($A) * [Math]::Cos($C) - [Math]::Tan($B) * [Math]::Sin($C)
    $E = [Math]::Cos($A)
    $F = Degrees ([Math]::Atan2($E, $D))
    Return (Circle $F)
}


Function Obliq {
# Obliquity of the ecliptic
# (GD As Double, GM As Double, GY As Double) As Double
Param (
       $GD, $GM, $GY
      )

    $A = CDJD $GD $GM $GY
    $B = $A - 2415020
    $C = ($B / 36525) - 1

    $D = $C * (46.815 + $C * (0.0006 - ($C * 0.00181)))
    $E = $D / 3600
    Return (23.43929167 - $E + (NutatObl $GD $GM $GY))
}


Function NutatObl {
# Nutation of the Obliquity of the Ecliptic
#(GD As Double, GM As Double, GY As Double) As Double
Param (
       $GD, $GM, $GY
      )

    DJ = (CDJD GD GM GY) - 2415020
    T = DJ / 36525
    T2 = T * T

    $L1 = 279.6967 + 0.000303 * $T2 + (Circle (100.0021358 * $T))
    $L2 = 2 * (Radians $L1)
    
    $D1 = 270.4342 - 0.001133 * $T2 + (Circle (1336.855231 * $T))
    $D2 = 2 * (Radians $D1)

    $M1 = 358.4758 - 0.00015 * $T2 + (Circle (99.99736056 * $T))
    $M1 = (Radians $M1)

    $M2 = 296.1046 + 0.009192 * $T2 + (Circle (1325.552359 * $T))
    $M2 = (Radians $M2)

    $N1 = 259.1833 + 0.002078 * $T2 - (Circle (5.372616667 * $T))
    $N1 = (Radians $N1)
    $N2 = 2 * $N1

    $DDO = (9.21 + 0.00091 * $T) * [Math]::Cos($N1)
    $DDO = $DDO + (0.5522 - 0.00029 * $T) * [Math]::Cos($L2) - 0.0904 * [Math]::Cos($N2)
    $DDO = $DDO + 0.0884 * [Math]::Cos($D2) + 0.0216 * [Math]::Cos($L2 + $M1)
    $DDO = $DDO + 0.0183 * [Math]::Cos($D2 - $N2) + 0.0113 * [Math]::Cos($D2 + $M2)
    $DDO = $DDO - 0.0093 * [Math]::Cos($L2 - $M1) - 0.0066 * [Math]::Cos($L2 - $N2)

    Return ($DDO / 3600)
}



Function NutatLong {
    [OutputType([Double])]
    Param ($JDN)

    $DJ = $JDN - 2415020    # Days since midday 0m 0s 1900/1/1 
    $T = $DJ / 36525   # Convert to centuries 
    $T2 = $T * $T


    $A = 100.0021359 * $T
    $B = Circle $A
    $L1 = 279.6966778 + 0.0003025 * $T2 + $B    # Sun's mean longitude
    $L2 = 2 * (Radians $L1)

    $A = 1336.855231 * $T
    $B = Circle $A
        
    $D1 = 270.4342 - 0.001133 * $T2 + $B
    $D2 = 2 * (Radians $D1)
    $A = 99.99736056 * $T
    $B = Circle $A

    $M1 = 358.4758 - 0.00015 * $T2 + $B
    $M1 = Radians $M1
    $A = 1325.552359 * $T
    $B = Circle $A
    $M2 = 296.1046 + 0.009192 * $T2 + $B
    $M2 = Radians $M2
    $A = 5.372616667 * $T
    $B = Circle $A
    $N1 = 259.1833 + 0.002078 * $T2 - $B
    $N1 = Radians $N1
    $N2 = 2 * $N1

    $DP = (-17.2327 - 0.01737 * $T) * [Math]::Sin($N1)
    $DP = $DP + (-1.2729 - 0.00013 * $T) * [Math]::Sin($L2) + 0.2088 * [Math]::Sin($N2)
    $DP = $DP - 0.2037 * [Math]::Sin($D2) + (0.1261 - 0.00031 * $T) * [Math]::Sin($M1)
    $DP = $DP + 0.0675 * [Math]::Sin($M2) - (0.0497 - 0.00012 * $T) * [Math]::Sin($L2 + $M1)
    $DP = $DP - 0.0342 * [Math]::Sin($D2 - $N1) - 0.0261 * [Math]::Sin($D2 + $M2)
    $DP = $DP + 0.0214 * [Math]::Sin($L2 - $M1) - 0.0149 * [Math]::Sin($L2 - $D2 + $M2)
    $DP = $DP + 0.0124 * [Math]::Sin($L2 - $N1) + 0.0114 * [Math]::Sin($D2 - $M2)

    Return ($DP / 3600)
}


Function SunLong {
    [OutputType([Double])]
    Param ($JDN)

    $DJ = $JDN - 2415020    # Days since midday 0m 0s 1900/1/1 
    $T = $DJ / 36525   # Convert to centuries 
    $T2 = $T * $T

    $A = 100.0021359 * $T
    $B = Circle $A
    $MeanLong = 279.6966778 + 0.0003025 * $T2 + $B    # Sun's mean longitude
    

    # Calculate initial value for longitude (= anomaly) assuming a circular orbit (Mean anomaly)
    $A = 99.99736042 * $T
    $B = Circle $A
    $M1 = 358.47583 - (0.00015 + 0.0000033 * $T) * $T2 + $B    # $M1 = Mean anomaly

    $EC = 0.01675104 - 0.0000418 * $T - 0.000000126 * $T2      # Eccentricity of Sun-Earth orbit  (0.01675104 = value at epoch)

    # Calculate True anomaly from mean anomaly using Kepler's Equation
    $AM = Radians $M1
    $TrueAnomaly = Degrees (TrueAnomaly $AM $EC)

#    $AE = EccentricAnomaly $AM $EC   # Not used

    # Calculate additional corrections

    $A = 62.55209472 * $T
    $B = Circle $A
    $A1 = Radians (153.23 + $B)

    $A = 125.1041894 * $T
    $B = Circle $A
    $B1 = Radians (216.57 + $B)

    $A = 91.56766028 * $T
    $B = Circle $A
    $C1 = Radians (312.69 + $B)

    $A = 1236.853095 * $T
    $B = Circle $A
    $D1 = Radians (350.74 - 0.00144 * $T2 + $B)

    $E1 = Radians (231.19 + 20.2 * $T)

    # Add corrections

    $Corrections = 0.00134 * [Math]::Cos($A1) + 0.00154 * [Math]::Cos($B1) + 0.002 * [Math]::Cos($C1)
    $Corrections = $Corrections + 0.00179 * [Math]::Sin($D1) + 0.00178 * [Math]::Sin($E1)
    
    Return (NormaliseDegrees ($TrueAnomaly + $MeanLong - $M1 + $Corrections))


    # $H1, $D3 - values not used...
    
    $A = 183.1353208 * $T
    $B = Circle $A
    $H1 = Radians (353.4 + $B)

    $D3 = 0.00000543 * [Math]::Sin($A1) + 0.00001575 * [Math]::Sin($B1)
    $D3 = $D3 + 0.00001627 * [Math]::Sin($C1) + 0.00003076 * [Math]::Cos($D1)
    $D3 = $D3 + 0.00000927 * [Math]::Sin($H1)
}


Function TrueAnomaly {
# Calcucate the True Anomaly from the Mean Anomaly and the Eccentricity
# This iteratively solves Kepler's equation, M = E - e Sin E
# (where: M = Mean Anomaly; E = Eccentric Anomaly; e = Eccentricity)
    [OutputType([Double])]
    Param (
        [Double]$MeanAnomaly,    # Mean anomaly
        [Double]$Eccentricity     # Eccentricity
    )

    $MeanAnomaly = NormaliseRadians $MeanAnomaly
    $EccentricAnomaly = $MeanAnomaly    # Set the eccentric anomaly to a starting value

    # Iteratively solve Kepler's equation (E - e sin E = M) to find the eccentric anomaly, E

    While ($True) {
        $Difference = $EccentricAnomaly - ($Eccentricity * [Math]::Sin($EccentricAnomaly)) - $MeanAnomaly   # Calculate difference - close enough yet?
        
        If ([Math]::Abs($Difference) -lt 0.000001) {
            Break      # Finish when sufficiently accurate
        }
        
        # Calculate correction term for next iteration
        $Difference = $Difference / (1 - ($Eccentricity * [Math]::Cos($EccentricAnomaly)))    
        $EccentricAnomaly = $EccentricAnomaly - $Difference
    }

    # Calculate true anomaly from the eccentic anomaly and eccentricity
    $A = [Math]::Sqrt((1 + $Eccentricity) / (1 - $Eccentricity)) * [Math]::Tan($EccentricAnomaly / 2)
    Return (2 * [Math]::Atan($A))
}


Function EccentricAnomaly {
    [OutputType([Double])]
    Param ($AM, $EC)

    $TP = 2 * [Math]::Pi
    $M = $AM - $TP * (Fix ($AM / $TP))
    $AE = $M

    While ($True) {

        $D = $AE - ($EC * [Math]::Sin($AE)) - $M
        
        If ([Math]::Abs($D) -lt 0.000001) {
            Return $AE
        }
        
        $D = $D / (1 - ($EC * [Math]::Cos($AE)))
        $AE = $AE - $D
    }
}


Function SunMeanAnomaly {
    [OutputType([Double])]
    Param ($JDN)

    $DJ = $JDN - 2415020    # Days since midday 0m 0s 1900/1/1 
    $T = $DJ / 36525   # Convert to centuries 
    $T2 = $T * $T

    $A = 100.0021359 * $T 
    $B = Circle $A          # convert to circular offset
    $M1 = 358.47583 - (0.00015 + 0.0000033 * $T) * $T2 + $B
    $AM = NormaliseRadians (Radians $M1)

    Return $AM
}


#region Helper Functions

# Convert Degrees to Radians
Function Radians {
[OutputType([Double])]
Param ($Degrees)
    Return ($Degrees * [Math]::Pi / 180.0)
}



# Convert Radians to Degrees
Function Degrees {
[OutputType([Double])]
Param ($Radians)
    Return ($Radians * 180.0 / [Math]::Pi)
}



# Convert decimal value into circular revolutions and reurn the degree offset
# e.g. if $Factor = 3.75, rotate around a circle that many time and return the result of 270 degrees
Function Circle {
[OutputType([Double])]
Param ($Rotations)
    Return (360 * (Frac $Rotations))
}



# Normalise angle in radians to range 0 - 2 Pi
# Practical Astronomy = "Unwind"

Function NormaliseRadians {
[OutputType([Double])]
Param ([Double]$Radians)

    $Normalised = $Radians % (2 * [Math]::Pi)
    If ($Normalised -lt 0) {$Normalised += 2 * [Math]::Pi}

    Return $Radians % (2 * [Math]::Pi)   # Normalise to range 0 - 2 Pi
}


# Normalise angle in degress to range 0-360
Function NormaliseDegrees {
[OutputType([Double])]
Param ([Double]$Degrees)

    $Normalised = $Degrees % 360
    If ($Normalised -lt 0) {$Normalised += 360}

    Return $Normalised   # Normalise to range 0 - 360
}



#  Just wraps the [Math]::Truncate() function...

Function Fix {
[OutputType([Long])]
Param ($Decimal)
    Return [Math]::Truncate($Decimal)
}



Function Frac {
[OutputType([Double])]
Param ($Decimal)
    Return $Decimal - [Math]::Truncate($Decimal)
}


#endregion


# ---------------------------------------------------------------------------------------------------------------------------


#region Unused Functions

# ConvertTo-JulianDate
# Calculate the Julian date.  This is the number of days since midday on 4713BCE.

Function CDJD {
    [OutputType([Double])]
    Param (
        $Year, 
        $Month, 
        $Day,
        $Hour = 0
    )
    
    # Move January and February dates into the previous year to enable to "average days per month" calculation to work

    If ($Month -lt 3) {
        $Y = $Year - 1
        $M = $Month + 12
    }
    else {
        $Y = $Year
        $M = $Month
    }


    # Calculate around 1582.  Leap year contributions are only required after this

    $Century = Fix ($Y / 100)

    # The following is a correction for dates after 15 October 1582 to fix up the following:
    # 1. 10 days were skipped to bring the calendar back in line with solar time (these days had accumulated because the earlier leap year rules were not sufficiently accurate)
    # 2. After this date, centuries are not leap years unless they are divisible by 400

    $LeapDays = 2 - $Century + (Fix ($Century / 4))   # Assuming date is after 15 October 1582 (for now...)
    
    If ($Year -le 1582) {
        If (($Year -eq 1582) -And ($Month -gt 10)) {
            # nop
        }
        Else {
            If (($Year -eq 1582) -And ($Month -eq 10) -And ($Day -ge 15)) {
                # Nop
            }
            Else {
                $LeapDays = 0
            }
        }
    }
          
    If ($Y -lt 0) {
        $YearDays = Fix ((365.25 * $Y) - 0.75)
    }
    Else {
        $YearDays = Fix (365.25 * $Y)
    }

    # Cumulative days in months of the year (average month is ~30.6 days)
    # (This is only valid if January and February dates are moved into the previous year...)

    $MonthDays = Fix (30.6001 * ($M + 1))
    
    # Add all the terms and return.  1720994.5 is the number of days between midday on 1st January 4713 BCE and Midnight on 1st January 0001

    Return  $YearDays + $LeapDays + $MonthDays + $Day + 1720994.5 + $Hour/24 
}



# Hours-Minutes-Seconds to Decimal Hours
Function HMSDH { 
    [OutputType([Double])]
    Param ($H, $M, $S)

    $A = [Math]::Abs($S) / 60
    $B = ([Math]::Abs($M) + $A) / 60
    $C = [Math]::Abs($H) + $B
    
    If (($H -lt 0) -Or ($M -lt 0) -Or ($S -lt 0)) {
        Return (-$C)
    }
    Else {
        Return $C
    }    
}




# (Local Civil Time to Greenwich Date)
Function LctGDate {
    [OutputType([PsCustomObject])]
    Param (
        [Double]$LCH, # Local Civil Hour
        [Double]$LCM, # Local Civil Minute
        [Double]$LCS, # Local Civil Second
         
        [Double]$DS,   # Daylight Saving
        [Double]$ZC,   # Time Zone Correction
        
        [Double]$LD,   # Local Day
        [Double]$LM,   # Local Month
        [Double]$LY    # Local Year
    ) 
    
    $A = HMSDH $LCH $LCM $LCS    # Hours-Minutes-Seconds to Decimal Hours
    $B = $A - $DS - $ZC
    $C = $LD + ($B / 24)
    $D = CDJD $LY $LM $C           # $D is now the Julian Day Number at Grenwich 
    $E = JDCDate $D
    $E1 = Fix ($E.Day)

    [PsCustomObject]@{
        Year = $E.Year
        Month = $E.Month
        Day = $E1                       # This function returns an integer day number
        Hour = (24 * ($E.Day - $E1))     # Decimal hours in day (0-24))
    }
}




# (Julian Day Number to Calendar Date)
Function JDCDate {
    [OutputType([PsCustomObject])]
    Param (
        [Double]$JD  # Julian Day Number
    )

    $I = Fix ($JD + 0.5)
    $F = $JD + 0.5 - $I
    $A = Fix (($I - 1867216.25) / 36524.25)
    
    If ($I -gt 2299160) {
        $B = $I + 1 + $A - (Fix ($A / 4))
    } 
    Else {
        $B = $I
    }
    
    $C = $B + 1524
    $D = Fix (($C - 122.1) / 365.25)
    $E = Fix (365.25 * $D)
    $G = Fix (($C - $E) / 30.6001)
    
    $JDCDay = $C - $E + $F - (Fix (30.6001 * $G))
    
    If ($G -lt 13.5) {
        $JDCMonth = $G - 1
    }
    Else {
        $JDCMonth = $G - 13
    }
    
    If ($JDCMonth -gt 2.5) {
        $JDCYear = $D - 4716
    } 
    Else {
        $JDCYear = $D - 4715
    }

    [PsCustomObject]@{
        Year  = $JDCYear
        Month = $JDCMonth
        Day   = $JDCDay
    }
}



# ConvertTo-JulianDate
# Calculate the Julian date.  This is the number of days since midday on 4713BCE.

Function CDJD {
    [OutputType([Double])]
    Param (
        $Year, 
        $Month, 
        $Day,
        $Hour = 0
    )
    
    # Move January and February dates into the previous year to enable to "average days per month" calculation to work

    If ($Month -lt 3) {
        $Y = $Year - 1
        $M = $Month + 12
    }
    else {
        $Y = $Year
        $M = $Month
    }


    # Calculate around 1582.  Leap year contributions are only required after this

    $Century = Fix ($Y / 100)

    # The following is a correction for dates after 15 October 1582 to fix up the following:
    # 1. 10 days were skipped to bring the calendar back in line with solar time (these days had accumulated because the earlier leap year rules were not sufficiently accurate)
    # 2. After this date, centuries are not leap years unless they are divisible by 400

    $LeapDays = 2 - $Century + (Fix ($Century / 4))   # Assuming date is after 15 October 1582 (for now...)
    
    If ($Year -le 1582) {
        If (($Year -eq 1582) -And ($Month -gt 10)) {
            # nop
        }
        Else {
            If (($Year -eq 1582) -And ($Month -eq 10) -And ($Day -ge 15)) {
                # Nop
            }
            Else {
                $LeapDays = 0
            }
        }
    }
          
    If ($Y -lt 0) {
        $YearDays = Fix ((365.25 * $Y) - 0.75)
    }
    Else {
        $YearDays = Fix (365.25 * $Y)
    }

    # Cumulative days in months of the year (average month is ~30.6 days)
    # (This is only valid if January and February dates are moved into the previous year...)

    $MonthDays = Fix (30.6001 * ($M + 1))
    
    # Add all the terms and return.  1720994.5 is the number of days between midday on 1st January 4713 BCE and Midnight on 1st January 0001

    Return  $YearDays + $LeapDays + $MonthDays + $Day + 1720994.5 + $Hour/24 
}




# (Local Civil Time to Univeral Time)
Function LctUT {
    [OutputType([Double])]
    Param (
        [Double]$LCH, # Local Civil Hour
        [Double]$LCM, # Local Civil Minute
        [Double]$LCS, # Local Civil Second
         
        [Double]$DS,   # Daylight Saving
        [Double]$ZC,   # Time Zone Correction
        
        [Double]$LD,   # Local Day
        [Double]$LM,   # Local Month
        [Double]$LY    # Local Year
    ) 

    $A = HMSDH $LCH $LCM $LCS    # Hours-Minutes-Seconds to Decimal Hours
    $B = $A - $DS - $ZC
    $C = $LD + ($B / 24)
    $D = CDJD $C $LM $LY            # $D is now the Julian Day Number at Grenwich 
    $E = (JDCDate $D).Day               # Julian Day Number to Calendar Day
    $E1 = Fix $E
    Return (24 * ($E - $E1))     # (Return only time in hours (0-24))
}



# (Local Civil Time to Greenwich Year)   ---! Always returns an [INT]  (...returned from JCDYear)
Function LctGYear {
    [OutputType([Int])]
    Param (
        [Double]$LCH, # Local Civil Hour
        [Double]$LCM, # Local Civil Minute
        [Double]$LCS, # Local Civil Second
         
        [Double]$DS,   # Daylight Saving
        [Double]$ZC,   # Time Zone Correction
        
        [Double]$LD,   # Local Day
        [Double]$LM,   # Local Month
        [Double]$LY    # Local Year
    ) 

    $A = HMSDH $LCH $LCM $LCS    # Hours-Minutes-Seconds to Decimal Hours
    $B = $A - $DS - $ZC
    $C = $LD + ($B / 24)
    $D = CDJD $C $LM $LY            # $D is now the Julian Day Number at Grenwich 
    Return (JDCDate $D).Year          # Julian Day Number to Calendar Year
}


# (Local Civil Time to Greenwich Month)  ---! Always returns an [INT]  (...returned from JCDMonth)
Function LctGMonth {
    [OutputType([Int])]
    Param (
        [Double]$LCH, # Local Civil Hour
        [Double]$LCM, # Local Civil Minute
        [Double]$LCS, # Local Civil Second
         
        [Double]$DS,   # Daylight Saving
        [Double]$ZC,   # Time Zone Correction
        
        [Double]$LD,   # Local Day
        [Double]$LM,   # Local Month
        [Double]$LY    # Local Year
    ) 

    $A = HMSDH $LCH $LCM $LCS    # Hours-Minutes-Seconds to Decimal Hours
    $B = $A - $DS - $ZC
    $C = $LD + ($B / 24)
    $D = CDJD $C $LM $LY            # $D is now the Julian Day Number at Grenwich 
    Return (JDCDate $D).Month         # Julian Day Number to Calendar Month
}



# (Local Civil Time to Greenwich Day)    ---! Always returns an [INT]
Function LctGDay {
    [OutputType([Int])]
    Param (
        [Double]$LCH, # Local Civil Hour
        [Double]$LCM, # Local Civil Minute
        [Double]$LCS, # Local Civil Second
         
        [Double]$DS,   # Daylight Saving
        [Double]$ZC,   # Time Zone Correction
        
        [Double]$LD,   # Local Day
        [Double]$LM,   # Local Month
        [Double]$LY    # Local Year
    ) 
    
    $A = HMSDH $LCH $LCM $LCS    # Hours-Minutes-Seconds to Decimal Hours
    $B = $A - $DS - $ZC
    $C = $LD + ($B / 24)
    $D = CDJD $C $LM $LY            # $D is now the Julian Day Number at Grenwich 
    $E = (JDCDate $D).Day               # Julian Day Number to Calendar Day
    Return (Fix $E)              # (Return only integer part)  (?Why)
}







# Check that the given year/month/day correspond to a valid date in the Julian calendar.
# For Julian dates we can't use the .Net [DateTime]::TryParse() method (all .Net dates are
# considered to be in the Gregorian calendar) so we have to do it by hand...

Function ValidJulianDate {
[OutputType([Bool])]
Param ($Year, $Month, $Day)

    If ($Year  -IsNot [Int] -Or
        $Month -IsNot [Int] -Or
        $Day   -IsNot [Int]) {
        Throw 'Non-integer parameter passed to ValidJulianDate function'      # Coding error somewhere - bail out
    }

    If ($Year -lt -4713) {Return $False}
    If ($Year -gt 9999)  {Return $False}
    If ($Year -eq 0)     {Return $False}
    If ($Month -lt 1)    {Return $False}
    If ($Month -gt 12)   {Return $False}
    If ($Day -lt 1)      {Return $False}

    $DaysInMonth = @(0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)

    If ($Month -eq 2 -and $Year%4 -eq 0) {    # Leap years are every 4 years in the Julian Calendar
        $DaysInMonth[2] = 29
    }

    If ($Day -gt $DaysInMonth[$Month]) {Return $False}

    # All looks ok...
    Return $True
}



# Check that the given year/month/day correspond to a valid date in the Gregorian calendar.
# For Gregorian dates we can use the .Net [DateTime]::TryParse() method.
# (This function assumes the given date is actually in the Gregorian calendar - i.e. after 14th Oct 1582)

Function ValidGregorianDate {
[OutputType([Bool])]
Param ($Year, $Month, $Day)

    If ($Year  -IsNot [Int] -Or
        $Month -IsNot [Int] -Or
        $Day   -IsNot [Int]) {
        Throw 'Non-integer parameter passed to ValidGregorianDate function'      # Coding error somewhere - bail out
    }

    [DateTime] $NotUsed = 0
    Return [DateTime]::TryParse("$Year/$Month/$Day", [Ref]$NotUsed)
}



# Determines whether the given date is 4th October 1582 or earlier, i.e. is considered 
# a date on the Julian calendar

Function IsJulianDate {
[OutputType([Bool])]
Param ($Year, $Month, $Day)

    If ($Year  -IsNot [Int] -Or
        $Month -IsNot [Int] -Or
        $Day   -IsNot [Int]) {
        Throw 'Non-integer parameter passed to IsJulianDate function'      # Coding error somewhere - bail out
    }

    # All dates prior to 1582 are in the Julian calendar
    If ($Year -lt 1582) {Return $True}

    # All dates after 1582 are in the Gregorian calendar
    If ($Year -gt 1582) {Return $False}

    # For 1582, check before October (Julian) or after October (Gregorian)
    If ($Month -lt 10)  {Return $True}
    If ($Month -gt 10)  {Return $False}

    # For October 1582, check if before the 5th (Julian) of after the 14th (Gregorian)
    If ($Day -lt 5)     {Return $True}
    If ($Day -gt 14)    {Return $False}

    # If we get to here then we have a date in the range 5/10/1582 to 14/10/1582 which
    # is not a valid date in either the Julian or the Gregorian calendar 
    Throw 'This date is not valid as it does not exist in either the Julian or the Gregorian calendars.'
}




# This is valid for Julian dates only... (In the Julian calendar, leap years were every 4 years)

Function IsJulianLeap {
[OutputType([Bool])]
Param($Year)

    If ($Year -IsNot [Int]) {
        Throw 'Non-integer parameter passed to IsJulianLeap function'      # Coding error somewhere - bail out
    }

    If (($Year %4) -eq 0) {Return $True}     # If the year's divisible by 4, it's a leap year

    Return $False                            # ...otherwise, it's not a leap year
}


# This is valid for Gregorian dates only... (In the Julian calendar, leap years were every 4 years)

Function IsGregorianLeap {
[OutputType([Bool])]
Param($Year)

    If ($Year -IsNot [Int]) {
        Throw 'Non-integer parameter passed to IsGregorianLeap function'      # Coding error somewhere - bail out
    }

    If (($Year %400) -eq 0) { Return $True  }     # If it a century divisible by 400, it's a leap year

    If (($Year %100) -eq 0) { Return $False }     # ...otherwise, if it's a century not divisible by 400, it's not a leap year

    If (($Year %4) -eq 0)   { Return $True  }     # ...otherwise, if the year's divisible by 4, it's a leap year

    Return $False                             # ...otherwise, it's not a leap year
}



Function Test-CDJD {
Param(
    [DateTime] $StartDate,
    [Int] $Days
)

    For ($i=0; $i -lt $Days; $i++) {
        $TestDate = $StartDate.AddDays($i)
        $y = $TestDate.Year; $m = $TestDate.Month; $d = $TestDate.Day
        $j1 = Convert-DateToJulianDayNumber $y $m $d
        $j2 = CDJD $y $m $d
        If ($j1 -ne $j2) {
            Write-Output "Test Date $y/$m/$d"
            Write-Output "Convert-DateToJulianDayNumber: $j1"
            Write-Output "                         CDJN: $j2"
        }
    }
}



# Calculate the Julian Day Number (JDN) from the given date.
# The JDN is the number of days since midday on 1st January 4713 BCE
# (Where 1st January 4713 BCE = date in proleptic Gregorian calendar)

Function Convert-DateToJulianDayNumber {
[OutputType([Double])]
Param ($Year, $Month, $Day)

    If ($Year  -IsNot [Int] -Or
        $Month -IsNot [Int] -Or
        $Day   -IsNot [Int]) {
        Throw 'Non-integer parameter passed to Convert-DateToJulianDayNumber function'      # Coding error somewhere - bail out
    }

    # The following lookup table gives a value for the cummulative number of days in the year upto the 
    # end of the previous month.  For example, for February, $CummulativeDaysInPriorMonths[2] = 31

    #             Table Index:        0  1   2   3   4    5    6    7    8    9   10   11   12   13
    #             Prior Month:        -  - Jan Feb Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec
    #    Days in Prior Months:             +31 +28 +31  +30  +31  +30  +31  +31  +30  +31  +30  +31 
    $CummulativeDaysInPriorMonths = @(0, 0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365)

    If (IsJulianDate $Year $Month $Day) {

        # Julian date, before 5th October 1582, use standard calculation

        $JulianLeapDays = Fix (($Year -1) / 4)    #  Julian leap years were every four years; add a day for each leap year before this year
        $LeapDay = $(If ($Month -gt 2 -and (IsJulianLeap $Year)) {1} else {0})   # Leap day for this year?

        $JDN = $Year * 365 + $JulianLeapDays + $CummulativeDaysInPriorMonths[$Month] + $Day + $LeapDay

    } 
    Else {

        # After 14th October 1582, Gregorian Calendar, correct for missing 1582 days and calculate for fewer leap years

        $GregorianLeapDays = (Fix (($Year -1) / 4)) - (Fix (($Year -1582 -1) / 100)) + (Fix (($Year -1582 -1) / 400))
        $LeapDay = $(If ($Month -gt 2 -and (IsGregorianLeap $Year)) {1} else {0})   # Leap day for this year?
        # Correct for 'missing' 10 days 
        $JDN = $Year * 365 + $GregorianLeapDays + $CummulativeDaysInPriorMonths[$Month] + $Day + $LeapDay - 10

    }

    # Add all the terms and return.  1720994.5 is the number of days between midday on 1st January 4713 BCE and Midnight on 1st January 0001
    Return $JDN + 1721057.5
    
    
}


#endregion
