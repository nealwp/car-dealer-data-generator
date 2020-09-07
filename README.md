# car-dealer-data-generator
Randomly generates a dataset of car dealer data. Output as CSV.

Also created a VBA version for users who do not have Python.

To utilize VBA version:
 * Open MS Access or MS Excel
 * In Access go to Database Tools > Visual Basic. In Excel go to Developer > Visual Basic. 
      * If you do not have the Developer Tab in Excel, right-click the ribbon and select "Customize the ribbon" and ensure the "Developer" box is checked in the right-hand pane.
 * From the VBA window, select File > Import File and browse to the .bas and .cls files. You will have to import each file seperately. 
 * In the makeCSVdata module, update the strPath and carCount variables to the desired values
 * Click the green play button on the ribbon or press F5.
 * File will output to desired path
