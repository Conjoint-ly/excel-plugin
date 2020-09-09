# Conjoint.ly Excel plugin
A free companion plugin for Excel that helps with charting [Conjoint.ly](https://conjointly.com/) outputs, including simulations charts from the Conjoint.ly [online simulator](https://conjointly.com/guides/conjoint-preference-share-simulator/) (scenario modelling and [price elasticity](https://conjointly.com/guides/understanding-price-elasticity-of-demand/) charts), colouring for [TURF analysis](https://conjointly.com/blog/turf-analysis/), and other useful utility functions.

#### Installation Guide (Windows 10)

To install the plugin first download the file by selecting ConjointlyExcelPlugin-v2.xlam, then clicking the button Download. 

![img](GuideImages/download.PNG)

Once the file is downloaded, move the file to an apropriate location in your file system. Rename the file to ConjointlyExcelPlugin-v2.install.

![img](GuideImages/Rename.PNG)

Right click on the file and select properties. At the bottom of the pop up box is the option to unblock the file. Select unblock and then apply.

![img](GuideImages/unblock.PNG)

Once this is completed, close all current instances of Microsoft Office. By double clicking on the file you will be prompted with the following message. 

![img](GuideImages/Prompt.PNG)

Select Yes to confirm the installation. To confirm that the program has installed successfully, check that the tab Conjoint.ly now appears at the top of your screen.

##### Optional Step - Previous Conjoint.ly Plugin Installed 
If you have installed the Conjoint.ly Plugin before September 2020, you will need to disable the exisiting plugin.

To do this, navigate to `Options` - `Excel Add-ins`. Select `Go` in the bottom left hand corner of the screen.

![img](GuideImages/Deselect.PNG)

You will now have `Conjointlyexcelplugin.Xlam` installed as well as `Conjointlyexcelplugin`. Deselect the non xlam variant

#### Functionality

![img](GuideImages/Tool bar.PNG)

| Button | Functionality |  
| ---------------------|--------------------------------------------------------------|
| *Go to Conjoint.ly** | Redirects the user to [Conjoint.ly](https://conjointly.com/) |
| **Make solid table** | Draw solid borders around all selected cells and will make the top row text bold|
| **Centre across cells** | Place the text from the first cell across all empty cells in that row without merging the cells|
| **Re-colour chart from cells** |Re-colour the selected chart based on cell background colours|
| **Trace for a range** | You know how you can trace     precedents and dependents for one cell in Excel? These buttons let you     check precedents and dependents for a range. These functions only find     precedents and dependents within the sheet. These functions do not show     arrows, but instead select the precedents and dependents.|
