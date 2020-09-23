VERSION 5.00
Object = "*\A..\..\..\..\..\DOCUME~1\Family\Desktop\NEWCHE~1\VB_COO~4\CoolFlex.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   1815
   ClientTop       =   2415
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   11040
   Begin CoolFlexGrid.CoolFlex CoolFlex1 
      Height          =   3255
      Left            =   480
      TabIndex        =   13
      Top             =   720
      Width           =   10095
      _extentx        =   17806
      _extenty        =   5741
   End
   Begin VB.CommandButton cmdHideCol 
      Caption         =   "Hide Col 3"
      Height          =   330
      Left            =   9660
      TabIndex        =   11
      Top             =   5040
      Width           =   1275
   End
   Begin VB.CommandButton cmdHideRow 
      Caption         =   "Hide Row 3"
      Height          =   330
      Left            =   8295
      TabIndex        =   10
      Top             =   5040
      Width           =   1275
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort"
      Height          =   330
      Left            =   1470
      TabIndex        =   9
      Top             =   5040
      Width           =   1275
   End
   Begin VB.CommandButton cmdMouse 
      Caption         =   "Where is Mouse"
      Height          =   330
      Left            =   105
      TabIndex        =   8
      Top             =   5040
      Width           =   1275
   End
   Begin VB.CommandButton cmdText 
      Caption         =   "Get Cell Text"
      Height          =   330
      Left            =   2835
      TabIndex        =   7
      Top             =   5040
      Width           =   1275
   End
   Begin VB.CommandButton cmdClearCbo1 
      Caption         =   "Clear cbo 1"
      Height          =   330
      Left            =   1470
      TabIndex        =   6
      Top             =   4620
      Width           =   1275
   End
   Begin VB.CommandButton cmdCboCount 
      Caption         =   "cbo 5 Count"
      Height          =   330
      Left            =   4200
      TabIndex        =   5
      Top             =   4620
      Width           =   1275
   End
   Begin VB.CommandButton cmdLockGrid 
      Caption         =   "Lock Grid"
      Height          =   330
      Left            =   8295
      TabIndex        =   4
      Top             =   4620
      Width           =   1275
   End
   Begin VB.CommandButton cmdClearGrid 
      Caption         =   "Clear Grid"
      Height          =   330
      Left            =   105
      TabIndex        =   3
      Top             =   4620
      Width           =   1275
   End
   Begin VB.CommandButton cmdGridLines 
      Caption         =   "Grid Lines"
      Height          =   330
      Left            =   9660
      TabIndex        =   2
      Top             =   4620
      Width           =   1275
   End
   Begin VB.CommandButton cmdClearCbo5 
      Caption         =   "Clear cbo 5"
      Height          =   330
      Left            =   2835
      TabIndex        =   1
      Top             =   4620
      Width           =   1275
   End
   Begin VB.CommandButton cmdCboRemove 
      Caption         =   "Remove Item from cbo 5"
      Height          =   330
      Left            =   5565
      TabIndex        =   0
      Top             =   4620
      Width           =   2640
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Top             =   5040
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    With Me.CoolFlex1

        .Cols = 8
        .Rows = 4

        'Turn editing on
        .EditEnable = True
        .EditTextLenght = 50

        'use this method to turn sorting on or off when the user clicks on the header column
        .SortOnHeaderClick True, SortGenericAscending
        .TextMatrix(0, 0) = "ID"
        .ColWidth(0) = 500
        'set column 1 to type Checkbox
        .TextMatrix(0, 1) = "Do"
        .ColWidth(1) = 330
        '.ColType 1, eCheckbox
        .TextMatrix(1, 1) = "C"
        'Call CoolFlex1_CellCheckBoxClick(1, 1)
        .TextMatrix(0, 2) = "Item"
        .ColWidth(2) = 1500
        'set column 2 to type Combobox
        .ColType 2, eCombobox
        .ComboBoxAddItem 2, "Col1Item1"
        .ComboBoxItemData(2, 0) = 11111
        .ComboBoxAddItem 2, "Col1Item2"
        .ComboBoxItemData(2, 1) = 22222
        .ComboBoxAddItem 2, "Col1Item3"
        .ComboBoxItemData(2, 2) = 33333
        .ComboBoxAddItem 2, "Col1Item4"
        .ComboBoxItemData(2, 3) = 44444
        .TextMatrix(0, 3) = "Last Match Date"
        .ColWidth(3) = 2000
        .TextMatrix(0, 4) = "Start Time"
        .ColWidth(4) = 1500
        .TextMatrix(0, 5) = "Stop Time"
        .ColWidth(5) = 1500
        .TextMatrix(0, 6) = "Duration"
        .ColWidth(6) = 1500
        'set column 7 to type Button
        .ColType 7, eButton
        .ColWidth(7) = 700

        'set column 1 to type Combobox
        'add 4 items to the column 1 combobox and set each items itemdata

        'set column 3 to type Checkbox
        '.ColType 3, eCheckbox

        'set column 1 to type Combobox
        '.ColType 5, eCombobox
        'add 6 items to the column 5 combobox
        '.ComboBoxAddItem 5, "Col5Item1"
        '.ComboBoxAddItem 5, "Col5Item2"
        '.ComboBoxAddItem 5, "Col5Item3"
        '.ComboBoxAddItem 5, "Col5Item4"
        '.ComboBoxAddItem 5, "Col5Item5"
        ' .ComboBoxAddItem 5, "Col5Item6"

        .ColType 1, eCheckbox

    End With 'ME.COOLFLEX1

End Sub

Private Sub cmdCboCount_Click()

  'gets item count for column 5 combobox

    MsgBox CoolFlex1.ComboBoxListCount(5)

End Sub

Private Sub cmdCboRemove_Click()

  'removex an item from column 5 combobox

  Dim x As Integer

    x = InputBox("What item do you want to remove from combobox column 5? (remember its 0 indexed)")
    CoolFlex1.ComboBoxRemoveItem 5, x

End Sub

Private Sub cmdClearCbo1_Click()

  'clears all items from column 1 combobox

    CoolFlex1.ComboBoxClear 1

End Sub

Private Sub cmdClearCbo5_Click()

  'clears all items from column 5 combobox

    CoolFlex1.ComboBoxClear 5

End Sub

Private Sub cmdClearGrid_Click()

  'clears whole grid

    CoolFlex1.Clear

End Sub

Private Sub cmdGridLines_Click()

  'changes grid line property

  Static x As Boolean

    If x = False Then
        CoolFlex1.GridLines = GridNone
        x = True
      Else 'NOT X...
        CoolFlex1.GridLines = GridFlat
        x = False
    End If

End Sub

Private Sub cmdLockGrid_Click()

  'shows how to enable and disable grid

    If cmdLockGrid.Caption = "Lock Grid" Then
        CoolFlex1.GridEnabled = False
        cmdLockGrid.Caption = "UnlockGrid"
      Else 'NOT CMDLOCKGRID.CAPTION...
        CoolFlex1.GridEnabled = True
        cmdLockGrid.Caption = "LockGrid"
    End If

End Sub

'this method used used to hide and unhide a column
'One thing though, it does not work well with column that have checkboxes in them.
'the checkbox picture does not go away
Private Sub cmdHideCol_Click()

    If cmdHideCol.Caption = "Hide Col 3" Then
        CoolFlex1.HideCol(3) = True
        cmdHideCol.Caption = "Unhide Col 3"
      Else 'NOT CMDHIDECOL.CAPTION...
        CoolFlex1.HideCol(3) = False
        cmdHideCol.Caption = "Hide Col 3"
    End If

End Sub

'this method used used to hide and unhide a row.
'One thing though, it does not work well with rows that have checkboxes in them.
'the checkbox picture does not go away
Private Sub cmdHideRow_Click()

    If cmdHideRow.Caption = "Hide Row 3" Then
        CoolFlex1.HideRow(3) = True
        cmdHideRow.Caption = "Unhide Row 3"
      Else 'NOT CMDHIDEROW.CAPTION...
        CoolFlex1.HideRow(3) = False
        cmdHideRow.Caption = "Hide Row 3"
    End If

End Sub

'Event that fires when a cells checkbox is clicked
Private Sub CoolFlex1_CellCheckBoxClick(ColIndex As Long, Value As Integer)

  'Call re_init_grid
  'MsgBox "checkbox was clicked"

End Sub

'Event that fires when a cells combobox is clicked
Private Sub CoolFlex1_CellComboBoxClick(ColIndex As Long, Value As String)

  'MsgBox "combobox was clicked"

End Sub

Private Sub cmdMouse_Click()

  'return row,col cord of mouse

    MsgBox CoolFlex1.MouseRow & ", " & CoolFlex1.MouseCol

End Sub

Private Sub cmdSort_Click()

  'sorts current column

    CoolFlex1.Sort = SortGenericAscending

End Sub

Private Sub cmdText_Click()

  'two ways to get cell text

    MsgBox CoolFlex1.Text
    MsgBox CoolFlex1.TextMatrix(2, 2)

End Sub

Private Sub CoolFlex1_CellButtonClick(RowIndex As Long, Value As Integer)

    Form1.Label1.Caption = "Save was clicked"

    MsgBox "The save button was clicked for row " & RowIndex, vbOKOnly, "Information on Click"

End Sub


