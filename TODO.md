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


#### Information
Cant capture everything on the Beginning sheet, otherwise we wont have any room for the viewports

1. Machine
1. JobNumer
2. Part number
3. Drawing Number
4. Revision
5. Part Description
6. Status?





