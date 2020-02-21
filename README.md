# Time-Series-Explorer
Tool to synchronize, interpolate gaps, analyze and plot any time-series data for multiple parameters, e.g. logged data from BEMS (Building Energy Management Systems) with sub-hourly logging. Includes options for correction for daylight saving time, user-definable date formats, basic statistics (e.g. weighted mean and median values), and carpet plot.

### Key features
- Can handle a virtually unlimited number of logged parameters, each of which have independent time stamps. You simply paste two columns of data for each parameter into sheet 'Input'. The first column is the date/time (Excel format), and the second column is the logged data value. For example:
Column A: Parameter 1 time-stamps
Column B: Parameter 1 values
Column C: Parameter 2 time-stamps
Column D: Parameter 2 values
Column E... and so on up to a maximum of over 8000 parameters.
- Can optionally handle cumulative parameters, e.g. logged energy use (e.g. kWh pulses, from smart electricity meters). For cumultive parameters you can choose to output rate of change, e.g. electrical power.
- Can handle any logger interval. The logger interval need not be constant, and can vary for the different parameters. Building Energy Managenens Systems (BEMS) systems typically have a logger interval of 5 or 10 minutes.
- The output in sheet 'Output' synchronizes the time-stamps of all the input parameters. The output time-step is user-definable. You can also choose the type of output values (e.g. first, min, mean, median, max, or last value within the time step) when the logger interval is shorter than the output time step. This feature enables you to calculate, hourly mean values, for example.
- You can choose whether periods with no data are filled by linear interpolation or left empty.
- Can correct for Daylight Saving Time (DST) if you wish. Output is always in standard time (user defined time-zone, e.g. UTC+1 in Europe), but with an additional column for DST hour of day, so that you can correctly assess time schedules.
- Can calculate weighted averages or sums of input parameters, e.g. area-weighted temperature in a building. Input the weights (e.g. zone area) for each parameter in row 1 above the parameter's date/time column.
- Carpet plot to visualize calculated hourly output.
- Includes rudimentary error-checking, such as wrong date-format, or dates/values pasted into the wrong column.

### History
- This tool is a further development of the original spreadsheet tool 'BEMS-hourly-values'. The name was changed because 'Time-Series Explorer" can output any time step, and has more general application, not only for BEMS.

### Installation and activation
- Simply download the spreadsheet file and open it in Microsoft Excel. No installation or registration is needed.
- However, this is a Visual Basic macro-enabled spreadhseet. You must activate macros for it to function: 
  - When you open the file for the first time in Excel, you will see a yellow bar at the top of the window, with the message *"PROTECTED VIEW Be careful... [Enable Editing]"*. Click on the 'Enable Editing' button. 
  - Next, depending on the security settings on your installation of Microsoft Excel, a red bar may appear at the top of the window, with the message *"BLOCKED CONTENT Macros in this document have been disabled..."*. This can be solved by moving the file to a directory on your PC that you designate for files that you trust, then open the file in Excel. To designate a 'trusted directory', click on **File > Options > Trust Center > Trust Center Settings > Trusted Locations > Add new location**, then browse to a directory, e.g. C:\TEMP\. Finally check that **Trust Center Settings > Trusted Documents > Disable Trusted Documents**  is not ticked.
  - When macros are properly activated, **BEMS-hourly-values** shows a small square splash-screen when you open the file. This splash screen shows the licence info, and advises you if an update is available for download from GitHub. Simply press 'Close' button to close the splash-screen. 
  - If still no splash-screen appears, then you might be able to fix it by menu option **File > Options > Trust Center > Trust Center Settings > Macro Settings > Disable all macros with notification**, which is a suitable level of security.
  
### License & warranty
- Distributed under the GLP v3 lisence (https://www.gnu.org/licenses/gpl-3.0.en.html). Please acknowledge/attribute use of this software in your report/publication with an appropriate author citation and URL to this site.
- Provided without warranty of any kind.

### Author & copyright owner
© Peter.Schild@OsloMet.no
