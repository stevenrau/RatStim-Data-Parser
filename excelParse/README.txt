///////////////////////////////////////////////////////////////////////////////
//                 ----------------------------------                        //
//                 RatStim Data Parser README / Manual                       //
//                 -----------------------------------                       //
//                                                                           //
// Written by: Steven Rau                                                    //
// For more help see: https://sourceforge.net/projects/excelparser/          //
// For source code, see: https://github.com/stevenrau/RatStim-Data-Parser    //
///////////////////////////////////////////////////////////////////////////////

--------------------------------------------------------------------------------

PLEASE NOTE:

This application is built for a very specific purpose with very little 
flexibility at this time. There is zero guarantee of program quality for any 
usage of this application outside of its intended purpose.

--------------------------------------------------------------------------------

Functionality:

The RatStim Data Parser takes in one or more CSV files (formatted in the
predetermined rat ID/stimulus layout) and sorts them according to rat ID and
stimulus value. The average values for each stimulus across all input files are
output to a 'Master' Excel file, and the individual sorted entries are kept
in an 'Intermediate' Excel file. This allows data to be examined at a higher
level in the 'Master' file, while still allowing closr analysis of the organized
data in the 'Intermediate' file.

--------------------------------------------------------------------------------

Requirements:

-To install and run successfully, the RatStim Data Parser application requires
.NET framework 3.5 or higher. All versions of Windows from XP onwards should
have this installed by deafult. If the installation fails, try to download and 
install .NET here: http://www.microsoft.com/net/downloads. If the newest .NET 
version is not compatible with your current Windows operating system, try 
chosing an older .NET version from the 'Earlier .NET Versions' tab.

--------------------------------------------------------------------------------
Instructions:

-To install the program, click on the setup.exe link in the download package
folder.
-If successful, the program should create a start menu shortcut 
under the RatStim folder for easy access.

- At the start screen, there are options to 'Browse' for input files, 'Clear 
input files', 'Save As...' and 'Sort'.

- The first thing you will want to do is click the 'Browse' button and then
select all of the input files you wish to sort. 
- You can click the 'Clear input files' button at any time if you wish to 
remove all the previously selected files and re-select others.
- Then, click the 'Save As...' button to selet the location where you want to 
save your output file. The file name you provide will be the location of the
'Master' output file, while the 'Intermediate' output file will be located in
a subdirectory named '<filename>_INTERMEDIATE_DATA'. If the output file you 
selected already exists, you will be prompted with a confirmation that you want 
to overwrite the existing 'Master' and 'Intermediate' files with that name.

- Once you have selected the input and output file locations, click the 'Sort'
button to organize the data and calculate the desired values.
- If successful, your output files will be saved to the location given and you
will be greeted with a success message window. From here you can click 'View 
file' to open the 'Master' Excel file.

-You can then repeat the process as many times as you would like.

