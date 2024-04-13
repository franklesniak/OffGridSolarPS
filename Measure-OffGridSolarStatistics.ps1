# Measure-OffGridSolarStatistics.ps1
# Version: 1.1.20240413.0

<#
.SYNOPSIS
    Analyzes solar and temperature data from https://nsrdb.nrel.gov/data-viewer to help
    design a resilient off-grid solar array with battery backup.

.DESCRIPTION
    This script reads in solar and temperature data from the National Solar Radiation
    Database (NSRDB) and calculates the following statistics:
    - Worst case solar power generation (rolling 24-hour period, in watts per square meter)
    - Worst case solar power generation (rolling three-day period, in watts per square meter)
    - Worst case solar power generation (rolling five-day period, in watts per square meter)
    - Worst case solar power generation (rolling seven-day period, in watts per square meter)
    - Worst case daily average "peak solar hours" (measured over a rolling 24-hour period, in hours)
    - Worst case daily average "peak solar hours" (measured over a rolling three-day period, in hours)
    - Worst case daily average "peak solar hours" (measured over a rolling five-day period, in hours)
    - Worst case daily average "peak solar hours" (measured over a rolling seven-day period, in hours)
    - Worst case daily average temperature (measured over a rolling 24-hour period, in degrees Celsius)
    - Worst case daily average temperature (measured over a rolling three-day period, in degrees Celsius)
    - Worst case daily average temperature (measured over a rolling five-day period, in degrees Celsius)
    - Worst case daily average temperature (measured over a rolling seven-day period, in degrees Celsius)

.PARAMETER PathToNSRDBDataFolder
    Path to the folder containing the extracted NSRDB data files.

.PARAMETER IgnoreYearSpecifiedInNSRDBData
    If specified, the script will ignore the year specified in the NSRDB data files.
    This is useful if you are using a "Typical Meteorological Year" dataset as that
    dataset combines multiple years of data into a single file.

.OUTPUTS
    A PSObject containing the calculated statistics.

.EXAMPLE
    $psobjectSolarStats = & .\Measure-OffGridSolarStatistics.ps1 -PathToNSRDBDataFolder "C:\Users\JDoe\NSRDBData"

.LINK
    https://github.com/franklesniak/OffGridSolarPS

#>

#region License ################################################################
# Copyright (c) 2024 Frank Lesniak
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of
# this software and associated documentation files (the "Software"), to deal in the
# Software without restriction, including without limitation the rights to use,
# copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
# Software, and to permit persons to whom the Software is furnished to do so,
# subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
# FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
# COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
# AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
# WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#endregion License ################################################################

[CmdletBinding()]
param (
    [parameter(Mandatory = $true)][ValidateScript({ Test-Path $_ -PathType Container })][string]$PathToNSRDBDataFolder,
    [switch]$IgnoreYearSpecifiedInNSRDBData
)

#region Functions ##################################################################
function Get-PSVersion {
    # Returns the version of PowerShell that is running, including on the original
    # release of Windows PowerShell (version 1.0)
    #
    # Example:
    # Get-PSVersion
    #
    # This example returns the version of PowerShell that is running. On versions
    # of PowerShell greater than or equal to version 2.0, this function returns the
    # equivalent of $PSVersionTable.PSVersion
    #
    # The function outputs a [version] object representing the version of
    # PowerShell that is running
    #
    # PowerShell 1.0 does not have a $PSVersionTable variable, so this function
    # returns [version]('1.0') on PowerShell 1.0

    #region License ############################################################
    # Copyright (c) 2024 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy
    # of this software and associated documentation files (the "Software"), to deal
    # in the Software without restriction, including without limitation the rights
    # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    # copies of the Software, and to permit persons to whom the Software is
    # furnished to do so, subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in
    # all copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    # SOFTWARE.
    #endregion License ############################################################

    #region DownloadLocationNotice #############################################
    # The most up-to-date version of this script can be found on the author's
    # GitHub repository at https://github.com/franklesniak/PowerShell_Resources
    #endregion DownloadLocationNotice #############################################

    $versionThisFunction = [version]('1.0.20240326.0')

    if (Test-Path variable:\PSVersionTable) {
        return ($PSVersionTable.PSVersion)
    } else {
        return ([version]('1.0'))
    }
}
#endregion Functions ##################################################################

$versionPS = Get-PSVersion
$boolIgnoreYearSpecifiedInNSRDBData = $false
if ($IgnoreYearSpecifiedInNSRDBData.IsPresent) {
    $boolIgnoreYearSpecifiedInNSRDBData = $true
}

$arrNSRDBDataFilePaths = @(Get-ChildItem -Path $PathToNSRDBDataFolder -Filter '*.csv' |
        Sort-Object -Property Name |
        ForEach-Object {
            if ($_.Name -notlike '*_fixed.csv') {
                $_.FullName
            }
        })

# Remove the leading two lines from each file
# Use [System.IO.File]::ReadAllLines() since it performs better than Get-Content
$arrRevisedFilePaths = @($arrNSRDBDataFilePaths | ForEach-Object {
        $strOriginalFilePath = $_

        # Strip off extension, add _fixed to the file name, and add extension back:
        $strRevisedFilePath = $strOriginalFilePath -replace '\.csv$', '_fixed.csv'

        $arrNSRDBData = [System.IO.File]::ReadAllLines($strOriginalFilePath)
        $arrNSRDBData = $arrNSRDBData[2..($arrNSRDBData.Length - 1)]
        [System.IO.File]::WriteAllLines($strRevisedFilePath, $arrNSRDBData)

        $strRevisedFilePath
    })

# Create the list to store received input data
if ($versionPS -ge ([version]'6.0')) {
    $listPSObjectNSRDBData = New-Object -TypeName 'System.Collections.Generic.List[PSObject]'
} else {
    # On Windows PowerShell (versions older than 6.x), we use an ArrayList instead
    # of a generic list
    # TODO: Fill in rationale for this
    $listPSObjectNSRDBData = New-Object -TypeName 'System.Collections.ArrayList'
}

$calendarGregorian = New-Object -TypeName System.Globalization.GregorianCalendar
$datetimekindUTC = [System.DateTimeKind]::Utc
$intCurrentYear = (Get-Date).Year

#region Collect Stats/Objects Needed for Writing Progress ##########################
$intProgressReportingFrequency = 1000
$intTotalItems = [int](($arrRevisedFilePaths.Count) * 365.2425 * 24) # May not be exactly correct, but we'll roll with it
$strProgressActivity = 'Importing data from NSRDB files into memory'
$strProgressStatus = 'Processing'
$strProgressCurrentOperationPrefix = 'Processing item'
$timedateStartOfLoop = Get-Date
# Create a queue for storing lagging timestamps for ETA calculation
$queueLaggingTimestamps = New-Object System.Collections.Queue
$queueLaggingTimestamps.Enqueue($timedateStartOfLoop)
#endregion Collect Stats/Objects Needed for Writing Progress ##########################

$intCounterLoop = 0
foreach ($strRevisedFilePath in $arrRevisedFilePaths) {
    $arrNSRDBData = Import-Csv -Path $strRevisedFilePath

    foreach ($objNSRDBData in $arrNSRDBData) {
        #region Report Progress ########################################################
        $intCurrentItemNumber = $intCounterLoop + 1 # Forward direction for loop
        if ((($intCurrentItemNumber -ge ($intProgressReportingFrequency * 3)) -and ($intCurrentItemNumber % $intProgressReportingFrequency -eq 0)) -or ($intCurrentItemNumber -eq $intTotalItems)) {
            # Create a progress bar after the first (3 x $intProgressReportingFrequency) items have been processed
            $timeDateLagging = $queueLaggingTimestamps.Dequeue()
            $datetimeNow = Get-Date
            $timespanTimeDelta = $datetimeNow - $timeDateLagging
            $intNumberOfItemsProcessedInTimespan = $intProgressReportingFrequency * ($queueLaggingTimestamps.Count + 1)
            $doublePercentageComplete = ($intCurrentItemNumber - 1) / $intTotalItems
            $intItemsRemaining = $intTotalItems - $intCurrentItemNumber + 1
            Write-Progress -Activity $strProgressActivity -Status $strProgressStatus -PercentComplete ($doublePercentageComplete * 100) -CurrentOperation ($strProgressCurrentOperationPrefix + ' ' + $intCurrentItemNumber + ' of ' + $intTotalItems + ' (' + [string]::Format('{0:0.00}', ($doublePercentageComplete * 100)) + '%)') -SecondsRemaining (($timespanTimeDelta.TotalSeconds / $intNumberOfItemsProcessedInTimespan) * $intItemsRemaining)
        }
        #endregion Report Progress ########################################################

        $psobjectComputedNSRDBData = New-Object -TypeName PSObject

        if ($boolIgnoreYearSpecifiedInNSRDBData) {
            $intYear = $intCurrentYear
        } else {
            $intYear = [int]($objNSRDBData.Year)
        }

        $arrDateTimeConstructorParams = @(
            $intYear,
            $objNSRDBData.Month,
            $objNSRDBData.Day,
            $objNSRDBData.Hour,
            $objNSRDBData.Minute,
            0,
            0,
            $calendarGregorian,
            $datetimekindUTC
        )

        $datetimeComputed = New-Object -TypeName 'System.DateTime' -ArgumentList $arrDateTimeConstructorParams

        $psobjectComputedNSRDBData | Add-Member -MemberType NoteProperty -Name 'DateTime' -Value $datetimeComputed
        $psobjectComputedNSRDBData | Add-Member -MemberType NoteProperty -Name 'GHI' -Value ([int]($objNSRDBData.GHI))
        $psobjectComputedNSRDBData | Add-Member -MemberType NoteProperty -Name 'Temperature' -Value ([double]($objNSRDBData.Temperature))

        # Add the updated object to the output list
        if ($versionPS -ge ([version]'6.0')) {
            $listPSObjectNSRDBData.Add($psobjectComputedNSRDBData)
        } else {
            [void]($listPSObjectNSRDBData.Add($psobjectComputedNSRDBData))
        }

        #region Post-Loop Progress Reporting ###########################################
        if ($intCurrentItemNumber -eq $intTotalItems) {
            Write-Progress -Activity $strProgressActivity -Status $strProgressStatus -Completed
        }
        if ($intCounterLoop % $intProgressReportingFrequency -eq 0) {
            # Add lagging timestamp to queue
            $queueLaggingTimestamps.Enqueue((Get-Date))
        }
        # Increment counter
        $intCounterLoop++
        #endregion Post-Loop Progress Reporting ###########################################
    }
}
Write-Progress -Activity $strProgressActivity -Status $strProgressStatus -Completed # Kill the progress bar just in case

$intMinGHI24HourPeriod = [int]::MaxValue
$datetimeMinGHI24HourPeriod = New-Object -TypeName 'System.DateTime'
$intMinGHI3DayPeriod = [int]::MaxValue
$datetimeMinGHI3DayPeriod = New-Object -TypeName 'System.DateTime'
$intMinGHI5DayPeriod = [int]::MaxValue
$datetimeMinGHI5DayPeriod = New-Object -TypeName 'System.DateTime'
$intMinGHI7DayPeriod = [int]::MaxValue
$datetimeMinGHI7DayPeriod = New-Object -TypeName 'System.DateTime'

$doubleMinAverageTemperature24HourPeriod = [double]::MaxValue
$datetimeMinAverageTemperature24HourPeriod = New-Object -TypeName 'System.DateTime'
$doubleMinAverageTemperature3DayPeriod = [double]::MaxValue
$datetimeMinAverageTemperature3DayPeriod = New-Object -TypeName 'System.DateTime'
$doubleMinAverageTemperature5DayPeriod = [double]::MaxValue
$datetimeMinAverageTemperature5DayPeriod = New-Object -TypeName 'System.DateTime'
$doubleMinAverageTemperature7DayPeriod = [double]::MaxValue
$datetimeMinAverageTemperature7DayPeriod = New-Object -TypeName 'System.DateTime'

# Create queues to store the last n-hours/days of data
$queueGHI24HourPeriod = New-Object -TypeName 'System.Collections.Generic.Queue[int]'
$queueGHI3DayPeriod = New-Object -TypeName 'System.Collections.Generic.Queue[int]'
$queueGHI5DayPeriod = New-Object -TypeName 'System.Collections.Generic.Queue[int]'
$queueGHI7DayPeriod = New-Object -TypeName 'System.Collections.Generic.Queue[int]'
$queueTemperature24HourPeriod = New-Object -TypeName 'System.Collections.Generic.Queue[double]'
$queueTemperature3DayPeriod = New-Object -TypeName 'System.Collections.Generic.Queue[double]'
$queueTemperature5DayPeriod = New-Object -TypeName 'System.Collections.Generic.Queue[double]'
$queueTemperature7DayPeriod = New-Object -TypeName 'System.Collections.Generic.Queue[double]'

#region Collect Stats/Objects Needed for Writing Progress ##########################
$intProgressReportingFrequency = 300
$intTotalItems = $listPSObjectNSRDBData.Count
$strProgressActivity = 'Computing stats for the range of data provided'
$strProgressStatus = 'Processing'
$strProgressCurrentOperationPrefix = 'Processing item'
$timedateStartOfLoop = Get-Date
# Create a queue for storing lagging timestamps for ETA calculation
$queueLaggingTimestamps = New-Object System.Collections.Queue
$queueLaggingTimestamps.Enqueue($timedateStartOfLoop)
#endregion Collect Stats/Objects Needed for Writing Progress ##########################

# Iterate through the list of PSObjects
$intCounterLoop = 0
$listPSObjectNSRDBData | Sort-Object -Property DateTime | ForEach-Object {
    $psobjectNSRDBData = $_

    #region Report Progress ########################################################
    $intCurrentItemNumber = $intCounterLoop + 1 # Forward direction for loop
    if ((($intCurrentItemNumber -ge ($intProgressReportingFrequency * 3)) -and ($intCurrentItemNumber % $intProgressReportingFrequency -eq 0)) -or ($intCurrentItemNumber -eq $intTotalItems)) {
        # Create a progress bar after the first (3 x $intProgressReportingFrequency) items have been processed
        $timeDateLagging = $queueLaggingTimestamps.Dequeue()
        $datetimeNow = Get-Date
        $timespanTimeDelta = $datetimeNow - $timeDateLagging
        $intNumberOfItemsProcessedInTimespan = $intProgressReportingFrequency * ($queueLaggingTimestamps.Count + 1)
        $doublePercentageComplete = ($intCurrentItemNumber - 1) / $intTotalItems
        $intItemsRemaining = $intTotalItems - $intCurrentItemNumber + 1
        Write-Progress -Activity $strProgressActivity -Status $strProgressStatus -PercentComplete ($doublePercentageComplete * 100) -CurrentOperation ($strProgressCurrentOperationPrefix + ' ' + $intCurrentItemNumber + ' of ' + $intTotalItems + ' (' + [string]::Format('{0:0.00}', ($doublePercentageComplete * 100)) + '%)') -SecondsRemaining (($timespanTimeDelta.TotalSeconds / $intNumberOfItemsProcessedInTimespan) * $intItemsRemaining)
    }
    #endregion Report Progress ########################################################

    $datetimeCurrent = $psobjectNSRDBData.DateTime
    $intGHI = $psobjectNSRDBData.GHI
    $doubleTemperature = $psobjectNSRDBData.Temperature

    # Update the stacks
    $queueGHI24HourPeriod.Enqueue($intGHI)
    $queueGHI3DayPeriod.Enqueue($intGHI)
    $queueGHI5DayPeriod.Enqueue($intGHI)
    $queueGHI7DayPeriod.Enqueue($intGHI)
    $queueTemperature24HourPeriod.Enqueue($doubleTemperature)
    $queueTemperature3DayPeriod.Enqueue($doubleTemperature)
    $queueTemperature5DayPeriod.Enqueue($doubleTemperature)
    $queueTemperature7DayPeriod.Enqueue($doubleTemperature)

    # Update the 24-hour period statistics
    if ($queueGHI24HourPeriod.Count -eq 24) {
        $intSumGHI24HourPeriod = $queueGHI24HourPeriod | Measure-Object -Sum | Select-Object -ExpandProperty Sum

        if ($intSumGHI24HourPeriod -lt $intMinGHI24HourPeriod) {
            $intMinGHI24HourPeriod = $intSumGHI24HourPeriod
            $datetimeMinGHI24HourPeriod = $datetimeCurrent
        }
    }
    if ($queueTemperature24HourPeriod.Count -eq 24) {
        $doubleAverageTemperature24HourPeriod = $queueTemperature24HourPeriod | Measure-Object -Average | Select-Object -ExpandProperty Average

        if ($doubleAverageTemperature24HourPeriod -lt $doubleMinAverageTemperature24HourPeriod) {
            $doubleMinAverageTemperature24HourPeriod = $doubleAverageTemperature24HourPeriod
            $datetimeMinAverageTemperature24HourPeriod = $datetimeCurrent
        }
    }

    # Update the 3-day period statistics
    if ($queueGHI3DayPeriod.Count -eq 72) {
        $intSumGHI3DayPeriod = $queueGHI3DayPeriod | Measure-Object -Sum | Select-Object -ExpandProperty Sum

        if ($intSumGHI3DayPeriod -lt $intMinGHI3DayPeriod) {
            $intMinGHI3DayPeriod = $intSumGHI3DayPeriod
            $datetimeMinGHI3DayPeriod = $datetimeCurrent
        }
    }
    if ($queueTemperature3DayPeriod.Count -eq 72) {
        $doubleAverageTemperature3DayPeriod = $queueTemperature3DayPeriod | Measure-Object -Average | Select-Object -ExpandProperty Average

        if ($doubleAverageTemperature3DayPeriod -lt $doubleMinAverageTemperature3DayPeriod) {
            $doubleMinAverageTemperature3DayPeriod = $doubleAverageTemperature3DayPeriod
            $datetimeMinAverageTemperature3DayPeriod = $datetimeCurrent
        }
    }

    # Update the 5-day period statistics
    if ($queueGHI5DayPeriod.Count -eq 120) {
        $intSumGHI5DayPeriod = $queueGHI5DayPeriod | Measure-Object -Sum | Select-Object -ExpandProperty Sum

        if ($intSumGHI5DayPeriod -lt $intMinGHI5DayPeriod) {
            $intMinGHI5DayPeriod = $intSumGHI5DayPeriod
            $datetimeMinGHI5DayPeriod = $datetimeCurrent
        }
    }
    if ($queueTemperature5DayPeriod.Count -eq 120) {
        $doubleAverageTemperature5DayPeriod = $queueTemperature5DayPeriod | Measure-Object -Average | Select-Object -ExpandProperty Average

        if ($doubleAverageTemperature5DayPeriod -lt $doubleMinAverageTemperature5DayPeriod) {
            $doubleMinAverageTemperature5DayPeriod = $doubleAverageTemperature5DayPeriod
            $datetimeMinAverageTemperature5DayPeriod = $datetimeCurrent
        }
    }

    # Update the 7-day period statistics
    if ($queueGHI7DayPeriod.Count -eq 168) {
        $intSumGHI7DayPeriod = $queueGHI7DayPeriod | Measure-Object -Sum | Select-Object -ExpandProperty Sum

        if ($intSumGHI7DayPeriod -lt $intMinGHI7DayPeriod) {
            $intMinGHI7DayPeriod = $intSumGHI7DayPeriod
            $datetimeMinGHI7DayPeriod = $datetimeCurrent
        }
    }
    if ($queueTemperature7DayPeriod.Count -eq 168) {
        $doubleAverageTemperature7DayPeriod = $queueTemperature7DayPeriod | Measure-Object -Average | Select-Object -ExpandProperty Average

        if ($doubleAverageTemperature7DayPeriod -lt $doubleMinAverageTemperature7DayPeriod) {
            $doubleMinAverageTemperature7DayPeriod = $doubleAverageTemperature7DayPeriod
            $datetimeMinAverageTemperature7DayPeriod = $datetimeCurrent
        }
    }

    # Pop the oldest data from the stacks
    if ($queueGHI24HourPeriod.Count -eq 24) {
        [void]($queueGHI24HourPeriod.Dequeue())
    }
    if ($queueGHI3DayPeriod.Count -eq 72) {
        [void]($queueGHI3DayPeriod.Dequeue())
    }
    if ($queueGHI5DayPeriod.Count -eq 120) {
        [void]($queueGHI5DayPeriod.Dequeue())
    }
    if ($queueGHI7DayPeriod.Count -eq 168) {
        [void]($queueGHI7DayPeriod.Dequeue())
    }
    if ($queueTemperature24HourPeriod.Count -eq 24) {
        [void]($queueTemperature24HourPeriod.Dequeue())
    }
    if ($queueTemperature3DayPeriod.Count -eq 72) {
        [void]($queueTemperature3DayPeriod.Dequeue())
    }
    if ($queueTemperature5DayPeriod.Count -eq 120) {
        [void]($queueTemperature5DayPeriod.Dequeue())
    }
    if ($queueTemperature7DayPeriod.Count -eq 168) {
        [void]($queueTemperature7DayPeriod.Dequeue())
    }

    #region Post-Loop Progress Reporting ###########################################
    if ($intCurrentItemNumber -eq $intTotalItems) {
        Write-Progress -Activity $strProgressActivity -Status $strProgressStatus -Completed
    }
    if ($intCounterLoop % $intProgressReportingFrequency -eq 0) {
        # Add lagging timestamp to queue
        $queueLaggingTimestamps.Enqueue((Get-Date))
    }
    # Increment counter
    $intCounterLoop++
    #endregion Post-Loop Progress Reporting ###########################################
}

$psobjectSolarStats = New-Object -TypeName PSObject
$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'WorstCaseSolarPowerGeneration24HourPeriod' -Value $intMinGHI24HourPeriod
$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'WorstCasePeakSolarHours24HourPeriod' -Value ([double]($intMinGHI24HourPeriod / 1 / 1000))
$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'DateTimeWorstCaseSolarPowerGeneration24HourPeriod' -Value $datetimeMinGHI24HourPeriod

$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'WorstCaseSolarPowerGeneration3DayPeriod' -Value $intMinGHI3DayPeriod
$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'WorstCasePeakSolarHours3DayPeriod' -Value ([double]($intMinGHI3DayPeriod / 3 / 1000))
$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'DateTimeWorstCaseSolarPowerGeneration3DayPeriod' -Value $datetimeMinGHI3DayPeriod

$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'WorstCaseSolarPowerGeneration5DayPeriod' -Value $intMinGHI5DayPeriod
$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'WorstCasePeakSolarHours5DayPeriod' -Value ([double]($intMinGHI5DayPeriod / 5 / 1000))
$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'DateTimeWorstCaseSolarPowerGeneration5DayPeriod' -Value $datetimeMinGHI5DayPeriod

$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'WorstCaseSolarPowerGeneration7DayPeriod' -Value $intMinGHI7DayPeriod
$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'WorstCasePeakSolarHours7DayPeriod' -Value ([double]($intMinGHI7DayPeriod / 7 / 1000))
$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'DateTimeWorstCaseSolarPowerGeneration7DayPeriod' -Value $datetimeMinGHI7DayPeriod

$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'WorstCaseAverageTemperature24HourPeriod' -Value $doubleMinAverageTemperature24HourPeriod
$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'DateTimeWorstCaseAverageTemperature24HourPeriod' -Value $datetimeMinAverageTemperature24HourPeriod

$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'WorstCaseAverageTemperature3DayPeriod' -Value $doubleMinAverageTemperature3DayPeriod
$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'DateTimeWorstCaseAverageTemperature3DayPeriod' -Value $datetimeMinAverageTemperature3DayPeriod

$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'WorstCaseAverageTemperature5DayPeriod' -Value $doubleMinAverageTemperature5DayPeriod
$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'DateTimeWorstCaseAverageTemperature5DayPeriod' -Value $datetimeMinAverageTemperature5DayPeriod

$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'WorstCaseAverageTemperature7DayPeriod' -Value $doubleMinAverageTemperature7DayPeriod
$psobjectSolarStats | Add-Member -MemberType NoteProperty -Name 'DateTimeWorstCaseAverageTemperature7DayPeriod' -Value $datetimeMinAverageTemperature7DayPeriod

return $psobjectSolarStats
