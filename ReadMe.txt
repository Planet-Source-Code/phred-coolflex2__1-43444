Version 1.0

This ocx is an extension of the EasyFlex.ocx that was originally developed by Muhammad Jauhari Saealal and submitted to FreeVBCode.  The original ocx used a textbox to enable easy editing functionality of an empty FlexGrid control.  

I needed a grid control that had more functionality so I've added the ability to define columns as combo boxes instead of text boxes.  I've also added the ability to define columns as checkboxes.  I've also exposed many more properties and methods of the underlying FlexGrid then the EasyGrid exposed.

The supplied sample app shows some of the new functionality.
Let me know if you find any bugs or have ideas for an upgrade.  
Thanks,
Jconwell@costco.com



Version 2.0   Aug 16, 2001

Added 3 new Events
CellCheckBoxClick: fires when a check box is clicked.  returns the column that the cell is in and the new value of the check box.

CellComboBoxClick: fires when a combo box is clicked.  returns the column that the cell is in and the new value of the check box.

CellComboBoxChange: fires when a combo box is changed.  returns the column that the cell is in.

RowColChange: same as the msflexgrid event.  Occurs when the currently active cell changes to a different cell

Added 7 new Properties
Col: same as the msflexgrid Col property
Row: same as the msflexgrid Row property
FixedColls: same as the msflexgrid FixedColls property
FixedRows: same as the msflexgrid FixedRows property
SortonHeaderClick: turns on or off grid sorting when a fixed column is clicked

Added 2 new Methods
HideCol: Pass in a col number to hide or show that column
HideRow: Pass in a row number to hide or show that row

Version 3.0 Feb 22 2003
Added command button in grid. Use "CoolFlex1_CellButtonClick" sub to code events after click.

Version 4.0 Apr 22, 2003
If you set a checkbox cell to "C" or "U" on load, the grid will display the checkbox accordingly.


Have fun!
phred@qti.net