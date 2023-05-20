# Portfolio-Profect-Excel-VBA

## Table of contents

* [Project Detail](#project-overview)
* [Prerequisite](#prerequisite)
* [Project Detail](#project-detail)


## Project Overview

This project is desing to showcase the code I create using [Excel](https://www.microsoft.com/en-us/microsoft-365/excel) fomular[VBA](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office) to automate task.
You can find the Excel file with the example data and VBA attched in this [folder](/Excel) the VBA code used in the file can be found [here](/Script).

## Prerequisite

To use this project you would need an [Excel](https://www.microsoft.com/en-us/microsoft-365/excel) programe with [Macro enable](https://www.excelcampus.com/vba/enable-developer-tab/).

Please clone this project with below code

```
git clone https://github.com/Chalermdej-l/Portfolio-Profect-Excel-VBA
```

This will download the [Excel file](/Excel) and [code](/Script).

## Project Detail

After clone the project you will find 2 excel file

1.Project_DashBoard

You can find the Dashboard create using excel function in `Portfolio 3` tab

![dashboard](/image/Excel_dashboard.png)


This dashboard show the monthly summary of the suppport call center stat like avg call per minute how many call take about what subject what agent receive the most call etc..

The data use in this dashboard can be found in `Portfolio 3_2` in this tab you can see the raw data and the calculation use for the dashboard.

2.Portfolio Project_VBA

Please note you may have to [enable trust](https://support.microsoft.com/en-us/topic/a-potentially-dangerous-macro-has-been-blocked-0952faa0-37e7-4316-b61d-5b5ed6024216) to this file as microsoft diable Macro by default.

2.1 First code `Portfolio1` use this [VBA code](/Script/Portfolio1_Loopandopenfile.bas) this code is desing to load the Excel and Text file data into this file

Please select `Import File` there will be a popup window to choose the file once choose the content will be insert into this sheet

![VBAtab1.png](/image/VBAtab1.png)

You can select `Clear Data` to remove all data in this page to load a new file

2.2 Second code `Portfolio 2` use this [VBA code](/Script/Portfolio2_LoopCellAndSavePDFdf.bas) this code is desing to export the data into WHT pdf format in  `Portfolio 2_2` tab

![wht](/image/Whtformat.png)

Please select `Save as pdf` button there will be a window to choose which Ref no. to start generated the data from.

![VBAtab2.png](/image/VBAtab2.png)

Once hit run the code will run depend on how many data there ar the code might take a minute or two to run 

After done there will be a folder create in the Desktop name `Print` folder this will store all the output pdf file generated

The code will also open this folder to check the output for any issue.

2.3 Thrid code `Portfolio 3` is desing to export the data into a text file for wht tax filing via the [RD Prep](https://efiling.rd.go.th/rd-cms/) program.

This sofeware can be use to file the Wht tax but you would need to prepare the data into the correct text format.

![VBAtab3.png](/image/VBAtab3.png)

This code will use the data in `Portfolio 3_2` and seperate them into different address type like province, City, Post code along with necessary data for the tax filing like tax rate tax amount.

You can chnage the number in `A2` cell to change between WHT3 and WHT53 form you can insert 3 or 53 in this cell and the fomula will change accordingly

Please select `Export03` button this will create a `WHTaxGenerate` folder in the Desktop with the text file generated.

The code will also open this folder to check the output file for any error.
