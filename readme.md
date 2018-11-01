# Bunker Report App
The aim of this Desktop application is to:
1. Extract information from Bunker Reports sent from the Vessels to the Technical Department of the Company in a *Microsoft Speadsheet Form (.xlsx)* 
2. Create a database file in .xlsx format as well that contains either **Weekly** information or a general **Monthly/Yearly** overview. 

Bunker Report App is a Windows Forms Desktop Application developed Visual Studio IDE using C# as a programming language. The following open-source libraries were also used in this project:
  - [EPPlus][epplus]
  - [LinqToExcel][ltex]

## Installation Instructions

**Step 1)** Make sure you have the compressed folder containing the application named *Release_Update.zip* in your computer. If you don't you can download it from [here][BetaRelease]

**Step 2)** Unzip the folder to a destination in your computer

**Step 3)** Double click on the *setup.exe*

**Step 4)** This will automatically run the application for the first time. From now on you can double click on the *BunkerReports.application* to use the App

#### Requirements 
> Windows Operating System with **.NET Framework 4.6.1** installed in the target machine


## Getting Started with Weekly Bunker Reports

This application helps you create a Weekly Bunker Report in just a few steps:

  - **Gather all the bunker reports sent by the vessels in a Folder**. This Folder may include files of other type as well, but make sure that all the *.xlsx files in the Folder are the Bunker Reports to be processed! The Bunker Reports need to be formatted strictly as the Sample.xlsx
  - **Determine the acceptable range for difference in Fuel Oil or Diesel Oil**. The default values in metric tonnes are: 
  - > -15 < FO < 15 
  - > -5 < DO < 5
  - Please check the box with the **Weekly** label.
  - Decide whether you want to create a New or Update an already existing Weekly Bunker Report.
    -  For a New Report choose the Directory for the Output. The final file will be created in a folder named *Database* in the user-specified Directory.   
    - To Update an already existing Weekly Bunker Report specify the Directory and the filename of the file to be updated. Please note that the file extention should NOT be included in the filename.
  - When the process is complete a database file in .xlsx format will be created, along with a .txt file with the same name. This .txt file contains information that helps the user to identify which reports where missing based on current Fleet Vessel Names. The exact location of the .txt file is included in the .txt so that the user can modify the Vessel Names if required! 

## Getting Started with Monthly/Yearly Bunker Reports

This application helps you create a Weekly Bunker Report in just a few steps:

  - **Gather all the bunker reports sent by the vessels in a Folder**. This Folder may include files of other type as well, but make sure that all the *.xlsx files in the Folder are the Bunker Reports to be processed! The Bunker Reports need to be formatted strictly as the Sample.xlsx
  - **Determine the acceptable range for difference in Fuel Oil or Diesel Oil**. The default values in metric tonnes are: 
  - > -15 < FO < 15 
  - > -5 < DO < 5
  - Please check the box with the **NonWeekly** label.
  - Decide whether you want to create a New or Update an already existing Weekly Bunker Report.
    -  For a New Report choose the Directory for the Output. The final file will be created in a folder named *Database* in the user-specified Directory.   
    - To Update an already existing Weekly Bunker Report specify the Directory and the filename of the file to be updated. Please note that the file extention should NOT be included in the filename.
  - When the process is complete a database file in .xlsx format will be created, along with a .txt file with the same name. This .txt file contains information that helps the user to identify which reports where missing based on current Fleet Vessel Names. The exact location of the .txt file is included in the .txt so that the user can modify the Vessel Names if required! 

## Comments

> If you wish to create New **Weekly** and **Monthly** Bunker reports is it possible, if you just check both boxes on the top right.
**BUT for Updating already existing files, you should open the app two(2) seperate times for each separate task!**


## How to update the Fleet.txt
> Note that: The Fleet.txt file contains the names of vessels in the company's bulk carrier fleet. Based on that information the application wil also create a .txt file as an output stating which bunker reports are missing for this week's bunker reporting. 
The user can customize the Fleet.txt file based on the company's latest fleet information. The filepath is written in any .txt output file after a successfull run of the application.

Support
---

Dimitra E. Anevlavi: dimitranevlavi@gmail.com

License
---

[![License: CC BY 4.0](https://img.shields.io/badge/License-CC%20BY%204.0-lightgrey.svg)](https://creativecommons.org/licenses/by/4.0/)






[//]: # (These are reference links used in the body of this note and get stripped out when the markdown processor does its job. There is no need to format nicely because it shouldn't be seen. Thanks SO - http://stackoverflow.com/questions/4823468/store-comments-in-markdown-syntax)


   [epplus]: https://github.com/JanKallman/EPPlus
   [ltex]: https://github.com/paulyoder/LinqToExcel
   [BetaRelease]: https://github.com/demieane/MarineBunkerReport/blob/master/Release_Update.zip
   
