# Bunker Report App
The aim of this Desktop application is to:
1. Extract information from Bunker Reports sent from the Vessels to the Technical Department of the Company in a *Microsoft Speadsheet Form (.xlsx)* 
2. Create a database file in .xlsx format as well that contains either **Weekly** information or a general **Monthly/Yearly** overview. 

Bunker Report App is a Windows Forms Desktop Application developed Visual Studio IDE using C# as a programming language. The following open-source libraries were also used in this project:
  - [EPPlus][epplus]
  - [LinqToExcel][ltex]

## Installation Instructions

**Step 1)** Make sure you have the compressed folder containing the application named *Bunker_Project.zip* in your computer.  
**Step 2)** Unzip the folder to a destination in your computer
**Step 3)** Double click on the *setup.exe*
**Step 4)** This will automatically run the application for the first time. From now on you can double clic on the *BunkerReports.application* to use the App

#### Requirements 
> Windows Operating System with **.NET Framework 4.6.1** installed in the target machine


## Getting Started

This application helps you create a Weekly Bunker Report in just a few steps:

  - Gather all the bunker reports sent by the vessels in a Folder. This Folder may include files of other type as well, but make sure that all the *.xlsx files in the Folder are the Bunker Reports to be processed! The Bunker Reports need to be formatted strictly as the Sample.xlsx
  - Decide whether you want to create a New or Update an already existing Weekly Bunker Report.
  - For a New Report choose the Directory for the Output.
  - To Update an already existing Weekly Bunker Report specify the Directory and the filename of the file to be updated. 

> The Fleet.txt file contains the names of vessels in the company's bulk carrier fleet. Based on that information the application wil also create a *.txt file as an output stating which bunker reports are missing for this week's bunker reporting. 

## Updating Fleet.txt
The user can customize the Fleet.txt file based on the company's latest fleet information. 

Support
---

Dimitra E. Anevlavi: dimitranevlavi@gmail.com

License
---

[![License: CC BY 4.0](https://img.shields.io/badge/License-CC%20BY%204.0-lightgrey.svg)](https://creativecommons.org/licenses/by/4.0/)






[//]: # (These are reference links used in the body of this note and get stripped out when the markdown processor does its job. There is no need to format nicely because it shouldn't be seen. Thanks SO - http://stackoverflow.com/questions/4823468/store-comments-in-markdown-syntax)


   [epplus]: https://github.com/JanKallman/EPPlus
   [ltex]: https://github.com/paulyoder/LinqToExcel
   [cvhelp]:https://joshclose.github.io/CsvHelper/
   
