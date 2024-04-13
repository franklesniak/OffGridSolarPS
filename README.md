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

## Remarks

For less critical systems, You may find it beneficial to compare the analysis of the selected "USA & Americas (30, 60 min / 4km / 1998-20xx)" to the data from "USA & Americas - Typical Meteorological Year", as the latter may show a less-pessimistic.
