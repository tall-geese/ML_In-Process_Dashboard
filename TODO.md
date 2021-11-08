### TODO

-----------------------------------------
#### Next up is...
- we need to get the machine type back from the Query
- Determine if the machine type is checked off as ML Ready
- if it is ready, then we also will need the ML Frequency sheet to be be visible and the Customer's AQL,
    too truly be 'ready'

- Using the runQty or Qty Complete, whichever is greater of the two, determine the number of AQL freq Inspections
- The number greater of the two numbers / number of inspections due is the frequency, may need to math.floor() here
- With the above complete, we can say for sure that yes, its true that it is MeasurLink ready, lets have some conditional hilighting to differentiate the good Jobs

- If the setup is 100%, then we should detect the number of number of shifts worked for the possibility of 1X shift routines.




#### Possible Logic
1. As Jesse is clicking from job to job, the viewport so refresh the amount of inspections captured?
   1. Are we also going to be updating all of the Epicor information as well? Whenever a job is selected? This information will be updated less frequently for sure.
   2. Whenever a job cell is selected, we need to be applying a highlight to that entire row to show that This is the job currently selected.
2. The routines that need to be captured for inspection will be determined by the cell that the job is running on. If its a *SWISS* cell, then we should be ready to check for all FA and IP routines that do not have any 'MILL' in the name. Vice-versa for the *MILL* cells.
   1. The FA Type will also play a part in this for the FA_FIRST and FA_MINI and such....
3. As we are selecting the next row of machine/jobs if the current selection is only one cell or every cell is in the current row, then in addition to loading the viewport information we should contextually highlight that cell


#### Information
Cant capture everything on the Beginning sheet, otherwise we wont have any room for the viewports

1. Machine
1. JobNumer
2. Part number
3. Drawing Number
4. Revision
5. Part Description
6. Status?


For the setup jobs that are not 100%, 
 - perhaps we load up a different set of graphic(s)
 - Maybe all we can show here it he Setup% completed and hours in so far?
 - Is there a field where ew can ge the est remaining hour of just the setup?


#### TODO
1. Design the chart styles that we need and save their templates for future use. Hopefully we can access these the same through VBA
2. The templates seem to always point to a default directory in the C Drive. If we cant change this we may need to have a sub program for first time runners to pull in the chart templates from the J drive to their local directory
   1. Doughnut charts for RM Hrs, % of all inspections
      1. (Can these graphs have callback functions when we click on them?)
      2. Yes they can, and the clalbacks will even trigger when the graph is clicked on a protected worksheet, where the user will not be able to change any other information regarding the graph


### Changes
1. Est remaining hours now adds the estimated setup hours * the percentage of the setup complete
2. Replace Total Hours we an estimated production hours as it goes nicer with the pie chart comparison. 
   1. This is based off the qty Completed  * the production standard and adds the est setup hours *  the percentage of the setup completed

