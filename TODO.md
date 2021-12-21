### TODO
1. Change cell color for jobName in the viewport 
   1. based on the customer the job belongs to
2. Redisign the layout of the viewport, easier to understand and has most relevant information
3. Is there an easy way to load info for ML jobs when doing an initial shop load?
   1. Having the information available at the beginning will mean we can set the **warning** status flag
4. Job ViewPort should be moved down as the inspector is selecting a cell they need to scroll down for
5. Can the charts be more animated by simply switching the ranges if they exist already??
   1. Run a test on this in another boook
6. If the user chooses not to refresh the shop load data, we will still need to clear the routine cells and arrows, since the arrow objects have no reference
7. When calculating the AQL for a job, need to make sure we are taking the maximum between the two of ProdQty and RunQty. 
   1. Also the Formula for the *Current AQL* should take into account if the Prod Qty is higher, then it doenst make sense to do a proportional value anymore.

### General
1. Need to add error handling to called functions and routines, if nothing else then to let us know where errors occurred, save some time on debugging.

### ShopStatus
1. Calls to activate the last active cell at the end of chart setting should probably be prefix with an
`on error resume next`
2. Routine and characteristic grouping, allow the user to select characteristics for a routine to view how they have been running
3. Doughnut chart should be formatted to have leading zeros for employee if they are less than 4 characters and dont have a "?" in them
4. Doughnut chart for setup%



### MeasurementInfo Sheet
#### Test
#### Insp Bar Chart
1. Add datalabels to the end of the bars that show the current amount of inspections done. Use SUM() of the range
   1. If and IP_{Somethign} has less than the current required AQL amount, we can turn either the data label red of the name of the routine red.
2. When collecting data, whenever we have one or more columns of inspections by employees, We should add a **TOTAL** column at the end of this
3. Then when creating the bar chart we should take this series and change its visibility to the *fill:None* and set data labels for this series and set this *position:base*. 
4. Finally iterate through each of the labels, grab the i*th* position of the Xvalues and depending on its routine type, we may want to change the formatting of either the routine or the data label.
5. We need a way to custom sort the order of the 

<br>


#### Older
-----------------------------------------

#### Possible Logic
2. The routines that need to be captured for inspection will be determined by the cell that the job is running on. If its a *SWISS* cell, then we should be ready to check for all FA and IP routines that do not have any 'MILL' in the name. Vice-versa for the *MILL* cells.
   1. The FA Type will also play a part in this for the FA_FIRST and FA_MINI and such....

#### TODO
1. Any functional use for tying callback methods to Charts?



