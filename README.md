# OffGridSolarPS

Analyzes solar and temperature data from <https://nsrdb.nrel.gov/data-viewer> to help design a resilient off-grid solar array with battery backup

## Instructions

Visit the [NSRDB database](https://nsrdb.nrel.gov/data-viewer) and search it for your location.

Then:

- For USA locations, select "USA & Americas (30, 60 min / 4km / 1998-20xx)".
- Check the checkbox for Temperature and GHI.
- Under "Select Year", click "Select All".
- For "Select Interval", leave "60 minutes" selected.
The script requires a 60-minute reporting interval.
- Check the checkbox for "Include Leap Day".
- Leave "Convert UTC to Local Time" unchecked.
- Enter your email address and click "Download".
An email will be sent with a link to a ZIP file.

Once you receive the email, download and extract the ZIP file to an empty folder. You will need this folder path to run the script.

## Usage

In PowerShell, run:

```powershell
$solarstats = & .\Measure-OffGridSolarStatistics.ps1 -PathToNSRDBDataFolder "C:\Users\JDoe\NSRDBData"

# The results are stored in $solarstats and can be viewed by typing: $solarstats
# Note that timestamps are reported in UTC
```

## Example

The author's location revealed the following statistics:

```powershell
PS C:\Users\flesniak> $results = & '.\Github\OffGridSolarPS\Measure-OffGridSolarStatistics.ps1' -PathToNSRDBDataFolder 'C:\Users\flesniak\Downloads\1deb49b30ae560262d64813009fa24a2'
PS C:\Users\flesniak> $results

WorstCaseSolarPowerGeneration24HourPeriod         : 114
WorstCasePeakSolarHours24HourPeriod                : 0.114
DateTimeWorstCaseSolarPowerGeneration24HourPeriod : 1/7/2019 8:30:00 PM
WorstCaseSolarPowerGeneration3DayPeriod           : 893
WorstCasePeakSolarHours3DayPeriod                  : 0.297666666666667
DateTimeWorstCaseSolarPowerGeneration3DayPeriod   : 12/25/2009 2:30:00 PM
WorstCaseSolarPowerGeneration5DayPeriod           : 1779
WorstCasePeakSolarHours5DayPeriod                  : 0.3558
DateTimeWorstCaseSolarPowerGeneration5DayPeriod   : 12/26/2009 7:30:00 PM
WorstCaseSolarPowerGeneration7DayPeriod           : 2763
WorstCasePeakSolarHours7DayPeriod                  : 0.394714285714286
DateTimeWorstCaseSolarPowerGeneration7DayPeriod   : 12/27/2009 5:30:00 PM
WorstCaseAverageTemperature24HourPeriod           : -27.2458333333333
DateTimeWorstCaseAverageTemperature24HourPeriod   : 1/31/2019 11:30:00 AM
WorstCaseAverageTemperature3DayPeriod             : -21.9930555555556
DateTimeWorstCaseAverageTemperature3DayPeriod     : 2/1/2019 5:30:00 AM
WorstCaseAverageTemperature5DayPeriod             : -17.3941666666667
DateTimeWorstCaseAverageTemperature5DayPeriod     : 2/1/2019 6:30:00 AM
WorstCaseAverageTemperature7DayPeriod             : -16.7839285714286
DateTimeWorstCaseAverageTemperature7DayPeriod     : 2/1/2019 5:30:00 AM
```

- The `WorstCaseSolarPowerGeneration` stats are reported in watts per meter squared (W/m^2)
- The `WorstCasePeakSolarHours` stats are reported in hours; these are normalized to represent the equivalent number of hours in the day to 1000 W/m^2 solar radiation.
This is a useful stat for determining how many watt-hours (Wh) a given solar panel will generate if you know its standardized wattage rating. For example, on a day with 4 peak solar hours, a fixed 100-watt panel would be expected to generate 400 watt-hours (not including any loss of power due to system inefficiencies).
- The `DateTime` properties are reported in UTC and are backward-looking.
For example, a `DateTimeWorstCaseSolarPowerGeneration24HourPeriod` of `1/7/2019 8:30:00 PM` means that the 24-hour timeframe from `1/6/2019 8:30:00 PM` UTC and `1/7/2019 8:30:00 PM` UTC was the worst 24 hours for solar power generation.
- The `WorstCaseAverageTemperature` stats are reported in degrees Celcius.

## Remarks

For less critical systems, You may find it beneficial to compare the analysis of the selected "USA & Americas (30, 60 min / 4km / 1998-20xx)" to the data from "USA & Americas - Typical Meteorological Year", as the latter may show a less-pessimistic.
