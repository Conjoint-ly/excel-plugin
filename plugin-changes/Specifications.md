# Conjoint.ly Excel Specifications - July 2020

## 1.1 - Changes to existing functions

#### 1.1.1 	Open List of Experiments
The button needs to be renamed to “Go to Conjoint.ly”.
 When selected, the button needs to redirect the user to www.conjoint.ly 
![Alt text](https://raw.githubusercontent.com/Conjoint-ly/excel-plugin/master/plugin-changes/images/Open%20List.png)

#### 1.1.2 	Precedents/Dependents
The functionality of both buttons needs to be combined into one button.
This single button needs to be named to Precedents/Dependents 
When selected should open up a two line drop down box with the two seperate options
![Alt text](https://raw.githubusercontent.com/Conjoint-ly/excel-plugin/master/plugin-changes/images/Precedents.png)

#### 1.1.3	Re-colour chart from cells
The current functionality needs to be expanded upon
When selected, the button will change the colors of the selected graph to the background colors of the data cells that it is dependent on
![Alt text](https://raw.githubusercontent.com/Conjoint-ly/excel-plugin/master/plugin-changes/images/Recolor.png)
When selected, the button should open a modal dialog box that has options for where to apply the colors, as well as what to do with the labels of the chart.
![Alt text](https://raw.githubusercontent.com/Conjoint-ly/excel-plugin/master/plugin-changes/images/Modal%20box.PNG)

The default selected values for colors should be Fill, Lines, Marker and for label formatting should be Value

When dark colors are applied to the graph, the label color should be changed to white to make it easier to see.

#### 1.1.4	Kill custom styles, Wrap formula in IF statement, Find Red, Make all FALSE red in selection, Make all errors red on sheet
These buttons need to be entirely deleted
![Alt text](https://raw.githubusercontent.com/Conjoint-ly/excel-plugin/master/plugin-changes/images/Excel%20Bar.PNG)

#### 1.1.5	ELASTICITY() function
The functionality needs to be changed to show the order of arguments

## 1.2 New functionality

#### 1.2.1	New Button - Copy conditional formatting
When selected, this button needs to copy the conditional formatting from a section as static formatting. 

An example use case of this is so that we can easily combine different formatting. In the following image we have a section of data formatted through color scales and a section formatted through data bars.
![Alt text](https://raw.githubusercontent.com/Conjoint-ly/excel-plugin/master/plugin-changes/images/Format%20Step%201.PNG)

Through this new button we want to be able to convert conditional formatting to static formatting, so that we can easily combine the formatting, as shown below. 

![Alt text](https://raw.githubusercontent.com/Conjoint-ly/excel-plugin/master/plugin-changes/images/Format%20Step%202.PNG)

Note - In the above example the two formatting sections are the same data, so applying both rules at once to the same section could achieve the same formatting. This is not what we want, as typically the data will be different 

#### 1.2.2	New Button - Hide Zero
The goal of this button is to remove the value labels from our charts when the value is close to zero. 

An example use case of this is the following graph 

![Alt text](https://raw.githubusercontent.com/Conjoint-ly/excel-plugin/master/plugin-changes/images/Graph%20Before.PNG)

Where column 2 and column 6 have values very close to zero. Selecting this button will transform the graph to the below graph, where columns 2 and 6 have had their label automatically deleted.

![Alt text](https://raw.githubusercontent.com/Conjoint-ly/excel-plugin/master/plugin-changes/images/Graph%20After.PNG)

We believe the regular express to achieve this is `[>0.01]#%;"";"";""`but we need this to be checked 


