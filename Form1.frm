VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Editable Data Grid"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List3 
      DragIcon        =   "Form1.frx":0000
      Height          =   840
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   1935
   End
   Begin VB.ListBox List2 
      DragIcon        =   "Form1.frx":0442
      Height          =   1230
      Left            =   5400
      Sorted          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   " Select a state "
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List1 
      DragIcon        =   "Form1.frx":0884
      Height          =   2010
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   " Select a city "
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\Program Files\Microsoft Visual Studio\VB98\Biblio.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\Program Files\Microsoft Visual Studio\VB98\Biblio.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Publishers"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0CC6
      DragIcon        =   "Form1.frx":0CDB
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Zip Codes (Drag a zip code from the list box below and drop on into a zip code cell in the grid.)"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   7335
   End
   Begin VB.Label Label1 
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   2760
      Width           =   6015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intHeight As Integer
Dim intCol As Integer
Dim intRow As Integer

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

'Important Note:
'The ADO Data Control's (adodc1) ConnectionString property is
'set to the following:
'
'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Program Files\Microsoft Visual Studio\VB98\Biblio.mdb;Persist Security Info=False
'
'If the BIBLIO.MDB database resides in a different directory on your
'machine you will need to change the Data Source location. On most
'machines this file will exist on the C drive.
'
'Another thing:
'If anyone knows how to drag & drop an item onto a Datagrid
'based on the X and Y coordinates of the mouse cursor please
'let me know!!! Email: kusluski@mail.ic.net

Private Sub DataGrid1_ButtonClick(ByVal ColIndex As Integer)
Dim strItem As String
On Error Resume Next
With DataGrid1
strItem = .Text
'Set height, move, select item, make visible, and
'give focus to list box
Select Case ColIndex
  Case 4
    List1.Height = (.Height / .RowHeight - (intRow - 1)) * .RowHeight
    List1.Move .Left + .Columns(4).Left, _
     .Top + .RowTop(.Row) + .RowHeight, _
     .Columns(4).Width
    If Len(strItem) Then
       List1 = strItem
    Else
       List1.ListIndex = 0
    End If
    List1.Visible = True
    List1.SetFocus
  Case 5
    If intRow > 4 Then 'place above cell
       List2.Height = (intRow + 1) * .RowHeight
       List2.Move .Left + .Columns(5).Left, _
        .Top + .RowHeight + (intRow * 1.4), _
        .Columns(5).Width
    Else 'place below cell
       List2.Height = (.Height / .RowHeight - (intRow + 1)) * .RowHeight
       List2.Move .Left + .Columns(5).Left, _
        .Top + .RowTop(.Row) + .RowHeight, _
        .Columns(5).Width
    End If
    If Len(strItem) Then
       'Find match in Listbox
       Dim n As Integer
       For n = 0 To List2.ListCount - 1
           If strItem = Left(List2.List(n), 2) Then
              'List2.Selected(n) = True
              List2.ListIndex = n
              Exit For
           End If
       Next
    Else
       List2.ListIndex = 0
    End If
    List2.Visible = True
    List2.SetFocus
End Select
End With
End Sub

Private Sub DataGrid1_DragDrop(Source As Control, X As Single, Y As Single)
If TypeOf Source Is ListBox Then
   'force a left mouse click
   mouse_event &H2, 0, 0, 0, 0
   mouse_event &H4, 0, 0, 0, 0
   DoEvents  'very important!
   If DataGrid1.Col = 6 Then
      DataGrid1.Text = Source.Text
      Source.Drag vbEndDrag
   Else 'incorrect field - cancel drag operation
      Source.Drag vbCancel
   End If
End If
End Sub

Private Sub DataGrid1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'If DataGrid1.Col = 6 Then
   If State = vbEnter Or State = vbOver Then
      Source.DragIcon = List2.DragIcon     'drop icon
   ElseIf State = vbLeave Then
      Source.DragIcon = DataGrid1.DragIcon 'no-drop icon
   End If
'Else
'   Source.DragIcon = DataGrid1.DragIcon    'no-drop icon
'End If
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
Dim strColName As String
Dim strSort As String
Static bSortAsc As Boolean
Static strPrevCol As String
On Error GoTo SortErr
strColName = DataGrid1.Columns(ColIndex).DataField
'If user clicked on same column check
'the previous state, in order to toggle
'between sorting ascending or descending.
'If this is the first time the user clicks on a column
'then sort ascending.
If strColName = strPrevCol Then
   If bSortAsc Then
      strSort = strColName & " DESC"
      bSortAsc = False
   Else
      strSort = strColName & " ASC"
      bSortAsc = True
   End If
Else
   strSort = strColName & " ASC"
   bSortAsc = True
End If
strPrevCol = strColName

Adodc1.Recordset.Sort = strSort
Label1.Caption = "Sorted by " & strSort

Exit Sub

SortErr:

Label1.Caption = "Can't sort!"

End Sub

Private Sub DataGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Label1.Caption = "X=" & X & ",Y=" & Y
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
intCol = DataGrid1.Col
intRow = DataGrid1.Row
'Label1.Caption = DataGrid1.Row
If List1.Visible Then
   List1.Visible = False
ElseIf List2.Visible Then
   List2.Visible = False
End If
End Sub

Private Sub DataGrid1_Scroll(Cancel As Integer)
If List1.Visible Then
   List1.Visible = False
ElseIf List2.Visible Then
   List2.Visible = False
End If
End Sub

Private Sub Form_Click()
If List1.Visible Then
   List1.Visible = False
ElseIf List2.Visible Then
   List2.Visible = False
End If
End Sub

Private Sub Form_Load()
Dim objConn As ADODB.Connection
Dim objRec As ADODB.Recordset
Dim objColumns As Columns
Dim strData As String
Dim varBM As Variant
On Error Resume Next
Me.Caption = "Dropdown Listboxes for City & State, Drag & Drop Zip Code, Column Sort"
Label1.Caption = Adodc1.Recordset.RecordCount & " records."
varBM = DataGrid1.Bookmark
Call AutosizeGridCols(DataGrid1, 5000, 5000)
DataGrid1.Bookmark = varBM
'add split to grid
DataGrid1.Splits.Add 1
DataGrid1.Splits.Item(0).ScrollGroup = 1 'setting to # greater than 1 will create unsynchronized split
DataGrid1.Splits.Item(0).SizeMode = dbgExact 'dbgScalable (default)
DataGrid1.Splits.Item(0).Size = 3000 'sets width of split in pixels
DataGrid1.Splits.Item(0).AllowFocus = False 'determines whether user can select or not
DataGrid1.Splits.Item(0).AllowRowSizing = False 'determines whether rows can be re-sized
DataGrid1.Splits.Item(0).AllowSizing = True 'determines whether split can be re-sized
DataGrid1.Splits.Item(0).Locked = True 'don't allow editing in split
intHeight = List1.Height
Set objConn = New ADODB.Connection
objConn.ConnectionString = Adodc1.ConnectionString
objConn.Open
Set objRec = New ADODB.Recordset
'Note: SQL below sorts the data. If this is used set the Sorted property
'of the ListBox to False
'objRec.Open "SELECT DISTINCT City FROM Publishers WHERE City IS NOT NULL ORDER BY City", objConn, adOpenForwardOnly
objRec.Open "SELECT DISTINCT City FROM Publishers WHERE City IS NOT NULL", objConn, adOpenForwardOnly
List1.AddItem "" 'add a blank item
Do While Not objRec.EOF  'populate the list box
   List1.AddItem objRec!City
   objRec.MoveNext
Loop
objRec.Close
objRec.Open "SELECT DISTINCT State FROM Publishers WHERE State IS NOT NULL", objConn, adOpenForwardOnly
List2.AddItem "" 'add a blank item
Do While Not objRec.EOF  'populate the list box
   Select Case Trim(objRec!State)
     Case "AK"
       strData = "AK - Arkansas"
     Case "CA"
       strData = "CA - California"
     Case "GA"
       strData = "GA - Georgia"
     Case "IL"
       strData = "IL - Illinois"
     Case "IN"
       strData = "IN - Indiana"
     Case "MA"
       strData = "MA - Massachussets"
     Case "MD"
       strData = "MD - Montana"
     Case "MI"
       strData = "MI - Michigan"
     Case "MN"
       strData = "MN - Minnesota"
     Case "NC"
       strData = "NC - North Carolina"
     Case "NJ"
       strData = "NJ - New Jersey"
     Case "NY"
       strData = "NY - New York"
     Case "OR"
       strData = "OR - Oregon"
     Case "PA"
       strData = "PA - Pennsylvania"
     Case "TX"
       strData = "TX - Texas"
     Case "WA"
       strData = "WA - Washington"
     Case Else
       strData = objRec!State
   End Select
   List2.AddItem strData
   objRec.MoveNext
Loop
objRec.Close
objRec.Open "SELECT DISTINCT Zip FROM Publishers WHERE Zip IS NOT NULL", objConn, adOpenForwardOnly
Do While Not objRec.EOF
   List3.AddItem objRec!zip
   objRec.MoveNext
Loop
Set objColumns = DataGrid1.Columns 'create Columns Object
objColumns.Item(0).Alignment = dbgRight 'right align column data
objColumns.Item(0).AllowSizing = False  'disable changing of column width
objColumns.Item(0).Caption = "Pub ID #" 'change column heading
objColumns.Item(0).Locked = True 'don't allow it to be selected
objColumns.Item(4).Button = True 'display button when selected
objColumns.Item(5).Button = True
objColumns.Item(5).Width = 2000  'set column width to 2000 pixels
'objColumns.Item(7).NumberFormat = "(###)###-####"  'format a number
objRec.Close
Set objRec = Nothing
objConn.Close
Set objConn = Nothing
End Sub

Private Sub Label1_Click()
If List1.Visible Then
   List1.Visible = False
ElseIf List2.Visible Then
   List2.Visible = False
End If
End Sub

Private Sub Label2_Click()
If List1.Visible Then
   List1.Visible = False
ElseIf List2.Visible Then
   List2.Visible = False
End If
End Sub

Private Sub List1_Click()
On Error Resume Next
DataGrid1.Text = List1
List1.Visible = False
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
   List1.Visible = False
ElseIf KeyCode = vbKeyReturn Then
   DataGrid1.Text = List1.Text
   List1.Visible = False
Else
   'This code keeps the list box displayed after a key is pressed
   SendKeys "{ENTER}"
   MsgBox ""
End If
End Sub

Private Sub List1_LostFocus()
List1.Visible = False
End Sub

Private Sub List2_Click()
On Error Resume Next
DataGrid1.Text = Left(List2.Text, 2)
List2.Visible = False
End Sub

Private Sub List2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
   List2.Visible = False
ElseIf KeyCode = vbKeyReturn Then
   DataGrid1.Text = Left(List2.Text, 2)
   List2.Visible = False
Else
   'This code keeps the list box displayed after a key is pressed
   SendKeys "{ENTER}"
   MsgBox ""
End If
End Sub

Private Sub List2_LostFocus()
List2.Visible = False
End Sub

Private Sub List3_Click()
If List1.Visible Then
   List1.Visible = False
ElseIf List2.Visible Then
   List2.Visible = False
End If
End Sub

Private Sub List3_DblClick()
If DataGrid1.Col = 6 Then
   DataGrid1.Text = List3.Text
End If
End Sub

Private Sub List3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
   List3.DragIcon = List1.DragIcon
   List3.Drag vbBeginDrag 'initiate the drag
End If
End Sub

Private Sub AutosizeGridCols(ByRef msFG As DataGrid, ByVal MaxRowsToParse As Integer, ByVal MaxColWidth As Integer)
Dim I As Integer
Dim J As Integer
Dim txtString As String
Dim intTempWidth As Integer
Dim intBiggestWidth As Integer
Dim intRows As Integer
Const intPadding = 150
With msFG
   For I = 0 To .Columns.Count - 1
       'Loops through every column
       .Col = I
       'Set the active colunm
       intRows = .ApproxCount
       'Set the number of rows
       If intRows > MaxRowsToParse Then intRows = MaxRowsToParse
          'If there are more rows of data, reset
          'intRows to the MaxRowsToParse constant
          intBiggestWidth = 0
          'Reset some values to 0
          For J = 0 To intRows - 1
              'check up to MaxRowsToParse # of rows and obtain
              'the greatest width of the cell contents
              .Row = J
              txtString = .Text
              intTempWidth = TextWidth(txtString) + intPadding
              'The intPadding constant compensates for text insets
              'You can adjust this value above as desired.
              If intTempWidth > intBiggestWidth Then intBiggestWidth = intTempWidth
              'Reset intBiggestWidth to the intMaxCol
              'Width value if necessary
          Next
          .Columns.Item(I).Width = intBiggestWidth
   Next
   'Now check to see if the columns aren't
   'as wide as the grid itself.
   'If not, determine the difference and expand each column proportionately
   'to fill the grid
   intTempWidth = 0
   For I = 0 To .Columns.Count - 1
       intTempWidth = intTempWidth + .Columns.Item(I).Width
       'Add up the width of all the columns
   Next
   If intTempWidth < msFG.Width Then
      ' Compute the width of the columns to the width of the grid control
      ' and if necessary expand the columns.
      intTempWidth = Fix((msFG.Width - intTempWidth) / .Columns.Count)
      ' Determine the amount od width expansion needed by each column
      For I = 0 To .Columns.Count - 1
          .Columns.Item(I).Width = .Columns.Item(I).Width + intTempWidth
          ' add the necessary width to each column
      Next
   End If
End With
End Sub
