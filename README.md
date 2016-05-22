# vba-column-width
Excel VBA macro to optimize column widths and formatting for printing.

## Use
* **Just run this macro:** Save the `vba-column-widths.xlam` file somewhere on your local machine. (Note: Do not use **Open > Save As** for this.) Then launch Excel and go to **File > Options > Add-ins > Go [Excel Add-ins] > Browse...** Navigate to the `vba-column-widths.xlam` file you just saved and click **Open**. "Vba-Column-Widths" should appear in the **Add-ins available** list, with the box next to it checked. Click **OK**. You should now have a button (green box) in the Quick Access Toolbar that says *Format To Print* when you hover over it--this launches the macro. You can also use the keyboard shortcut **Ctrl-Shift-F**.
* **Use as part of another macro:** Import the `ColumnWidths.bas` module into your project.
* **Make changes to primary macro:** The `vba-column-widths.xlsm` file is for development. When your changes are ready to go, save the file as `.xlam` to be used as an Add-in. It also doesn't hurt to export the `ColumnWidths` module for safekeeping.

## To do
* Test on Excel for Mac 2011
* Error handling:
  * General
  * Specified font doesn't exist
  * Requested paper size is not available
* Expand `GetPageWidth` to include all options
* See if it's possible to return only paper sizes that are currently available on the printer, not that the printer is capable of using.
* Test if numbers appear as *####* because they don't fit in the column width (specifically ISBN). This is somewhat dependent on View > Zoom level, so may not be possible.
* Is it possible to better handle columns where only a few values are long, throwing off the .Columns.AutoFit method? Without actually looping through every cell in the sheet. 
