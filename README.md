# ZapMVImager

## Description
Tool to extract MV imager data from log files and show them 
on a chart or export them to different types of files.

## Installation
To get the EXE files for the app, you have to compile them. For 
this you need the .Net SDK 6 or 7 and the .Net Framework 4.7.2.
You get both in the internet. Search with Google for them. They 
are free. Download first the .Net 6/7 SDK and install it. Then 
download the .Net Framework 4.7.2. Install it too. Test it by 
opening a Powershell and run "dotnet build". You should see the
version of the installed .Net SDK.

On this webpage go in the right upper part. There is a green button 
"<> Code". Open it and select "Download ZIP". You will download 
the source code as a Zip file. Open it and copy the content at 
a place you remember.

Open a Powershell, go to the folder where you stored the content
of the Zip file and call the command "dotnet build". The compiler 
takes some time. When it is ready, you can go to the folder 
"bin\Debug\net472". There you will find the app "ZapMVImager.exe".
Start it.

## How to use

### Add logs
With the two buttons “Add file(s)” and “Add folder(s)” you select 
the log files. If you use “Add file(s)” you could import one or 
more log text files. If you change the filter for the extension 
(lower right side), it is also possible to add zip files. 
Normally you should find them on the Service drive of your TPS 
as backup files from the TDS (normally named as OP_DP1007_2023-04-14_143610.zip). 
They contain the log files from the TDS for the given day in a 
subfolder. All TreatmentView log files in a zip file are added. 
With the “Add folder” you could import one or more folders. All 
TreatmentView log files that are found recursively in this folders 
are used. If there are zip files inside of this folders, they are 
examined as above described.

### Extract data
If you have all the log files added, that should be used, you use 
the button “Extract”. The app starts to go through all this log 
files. The files are normally ordered after date and log file 
extension (if a log file is greater than 25 MB, it is renamed to 
log.1 and a new one is used). It checks all the lines and extracts 
the MV imager entries. You could see the progress in the status 
bar (number of different plans, number of entries and the file name, 
which is checked in this moment. Takes some time dependent by the 
number of logfiles you want to import.

### Examin chart
After this process finished, you could select the found plans in a 
combo box. If the same plan name is found in log files from different 
dates (fractions or makeup fractions), you could also select the date 
you want to examine. In the lower part of the screen you see the chart 
of the found entries. If you move the mouse around in the chart, the 
values for this entry are shown in the status bar at the lower part. 
There you find the date and time of this beam, isocenter and node with 
axial and oblique positions, planned, delivered and detected MUs, 
difference between delivered and detected MUs in percent and the cumulative 
difference over time. Zoom with the scroll wheel of the mouse, left mouse 
button for panning, click with middle mouse button to reset zoom and so on.

#### Chart manipulation
Following mouse operations are supported.  
**Left click drag**: Pan  
**Right click drag**: Zoom  
**Middle click drag**: Zoom region  
**Alt left click drag**: Zoom region  
**Scroll wheel**: Zoom to cursor  
**Right click**: Show context menu  
**Middle click**: Automatically adapt zoom to view whole data  
**Ctrl left click drag**: Pan horizontally  
**Shift left click drag**: Pan vertically  
**Ctrl right click drag**: Zoom horizontally  
**Shift right click drag**: Zoom vertically  
**Ctrl shift right click drag**: Zoom evenly  

#### Entry details
When you move with the mouse cusror inside of the chart, you get a 
cursor in the chart. With this, you could select, which entry do you 
want to examine. The values then are shown in the statusbar. From left to 
right you find:
1. Date and time when the measurement was done  
2. Isocenter and node with axial and oblique position of beam. Cave: The node  
must not be the same as the beam number. So it could be, that you start with 
the beam #1 at node #0. 
3. Planned MU as given by the plan
4. Delivered MU, which could slightly differ from the planned MU
5. Imager MU as it is detekted by the MV imager
6. Difference between delivered MUs and MV imager MUs in percent
7. Difference of sum of delivered MUs and MV imager MUs up to this beam in percent

#### Context menu
With the context menu of the chart, you could do the following things.  
**Copy**: Copy the current chart to the clipboard. Title is added.  
**Show cumulative for inside 10 %**: Add another plot, that shows the cumaltive 
dose difference in percent but only for the points, which difference is between 
-10.0 % and 10.0. The idea behind is, that when dose is more than 10 % off, 
there must something totally wrong. So we drop this measurements.  
**Open in new window**: Open a new window with the current chart as content. Title is added.  

### Export
At the end it is possible with the “Export” button to export the data in 
different formats. 

#### CSV - Comma separated values
Export the data belonging to the selected plan and date to a text file, 
each entry of chart in an extra line with a leading header. By default a 
semicolon is used as separator between the columns. Another character 
could be choosen in the config file. The entry is called "CSVSeparator".

#### XLSX - Excel OpenXML format
Export data belonging to the selected plan and date or for all dates into 
a XLSX file in OpenXML format, which then could be used with MS Excel or 
OpenOffice Calc. To distingush between this two you use the filter in the 
combo box below the text field, where you could input the filename, which 
you want to use for saving the data.

If you decide to get data for all dates, you get a XLSX file, which contain 
a worksheet for each date. Perhaps you have to change the name to for saving.

#### PNG or JPG - Graphics formats
Export the chart as you see it on the screen. Only difference is, that the 
chart gets a title containing plan name and date. There are two possible 
formats: PNG and JPG. App distinguish between them by the extension of the 
name.
