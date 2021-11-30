### TODO
- Change cell color for jobName in the viewport
  - based on the customer the job belongs to

- Add another col to ShopLoad for completed# Good inspections
  - and #shifts worked, and how many inspections we should be expecting

### ShopStatus
1. Time series cleanup sub needs to be changed
   1. The bar chart and time series are going to be sharing the same space...

### MeasurementInfo Sheet
1. The main callable should return the range of routines
   1. We should be offsetting one by one and looking at the cell above the first one adn seeing that there is an emlpoyee # there
2. Create a Cleanup() subroutine and link it to the ShopLoad

#### Test
Can we actually pass a collection of range objects?
I kind of recall that this doesnt work....
`yes, this will work but just be sure to declare the variable explcitly as a range before iterating through the collection`

#### Insp Bar Chart
1. Before we go setting the vertical lines, first lets get a bar graph set conditionally in the viewport
2. Our lines dont need to reference a literal range, they can use a literal value (figure out how this is done in the code)
   1. So this information should be set in the ShopLoad tab, along with our AQL req, the current insps required (calculated) and # of shifts worked

#### older
-----------------------------------------
#### Next up is...
- The number greater of the two numbers / number of inspections due is the frequency, may need to math.floor() here
- If the setup is 100%, then we should detect the number of number of shifts worked for the possibility of 1X shift routines.


#### Possible Logic
2. The routines that need to be captured for inspection will be determined by the cell that the job is running on. If its a *SWISS* cell, then we should be ready to check for all FA and IP routines that do not have any 'MILL' in the name. Vice-versa for the *MILL* cells.
   1. The FA Type will also play a part in this for the FA_FIRST and FA_MINI and such....

#### TODO
1. Any functional use for tying callback methods to Charts?



