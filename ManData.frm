VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ManData 
   Caption         =   "Database Processing"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11880
   LinkTopic       =   "ManData"
   MDIChild        =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "Update"
      Height          =   252
      Left            =   2160
      TabIndex        =   16
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Refresh"
      Height          =   252
      Left            =   1920
      TabIndex        =   15
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete Query"
      Height          =   252
      Left            =   6960
      TabIndex        =   14
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Query"
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox textSQL 
      Height          =   1080
      Left            =   8400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   1680
      Width           =   3330
   End
   Begin VB.ListBox QryList 
      Height          =   840
      Left            =   4920
      TabIndex        =   8
      Top             =   1680
      Width           =   3375
   End
   Begin VB.ListBox FldList 
      Height          =   840
      Left            =   8400
      TabIndex        =   6
      Top             =   240
      Width           =   3375
   End
   Begin VB.ListBox TblList 
      Height          =   840
      Left            =   4920
      TabIndex        =   5
      Top             =   240
      Width           =   3375
   End
   Begin VB.TextBox txtSQL 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1215
      Width           =   4620
   End
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7560
      Width           =   11655
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "ManData.frx":0000
      Height          =   4455
      Left            =   120
      TabIndex        =   22
      Top             =   3000
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7858
      _Version        =   393216
   End
   Begin VB.Label Queryname 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label percent 
      Height          =   255
      Left            =   7560
      TabIndex        =   20
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Percent%:"
      Height          =   255
      Left            =   6720
      TabIndex        =   19
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "records:"
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label reccount 
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Query Definition"
      Height          =   255
      Left            =   8400
      TabIndex        =   11
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "Fields(Variables)"
      Height          =   255
      Left            =   8520
      TabIndex        =   10
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Queries (Double click to load query)"
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label Label5 
      Caption         =   "Table(Click for details)"
      Height          =   495
      Left            =   4800
      TabIndex        =   7
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   135
      TabIndex        =   4
      Top             =   75
      Width           =   1725
   End
   Begin VB.Label Label3 
      Caption         =   "SQL Query"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   3
      Top             =   720
      Width           =   1680
   End
   Begin VB.Label Label2 
      Caption         =   "Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   2
      Top             =   2760
      Width           =   1590
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   135
      TabIndex        =   0
      Top             =   360
      Width           =   4605
   End
   Begin VB.Menu DatabaseManagement 
      Caption         =   "Database"
      Begin VB.Menu OpenDB 
         Caption         =   "Open Database"
      End
      Begin VB.Menu CloseDatabase 
         Caption         =   "Close Database"
      End
   End
   Begin VB.Menu ExecSQL 
      Caption         =   "Execute SQL Query"
   End
   Begin VB.Menu Statmenu 
      Caption         =   "Statistical Processing"
   End
   Begin VB.Menu GraphChart 
      Caption         =   "Graphs"
      Begin VB.Menu LineGraphMenu 
         Caption         =   "Line Graph"
      End
      Begin VB.Menu PieGraphMenu 
         Caption         =   "Pie Graph"
      End
   End
   Begin VB.Menu LinkToImageProcForm 
      Caption         =   "GoToImageProcessingForm"
   End
   Begin VB.Menu End 
      Caption         =   "End"
   End
End
Attribute VB_Name = "ManData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DB As Database
Dim tbl As TableDef
Dim TName As String
Dim idx As Index
Dim qry As QueryDef
Dim initialreccount As Integer
Public statistical_processing As Integer
Public linegraphprocessing As Integer
Public piegraphprocessing As Integer

Private Sub CloseDatabase_Click()
Unload Me
Load ManData
ManData.Show
End Sub

Private Sub Command1_Click() ' SAVE QUERY
Dim qdfNew As QueryDef
Dim qdfname As String

On Error GoTo SQLError

qdfname = InputBox("Give a name for the query:", , qdfname)
Set qdfNew = DB.CreateQueryDef(qdfname, _
            txtSQL.Text)
QryList.AddItem qdfNew.Name
QryList.refresh

SQLError:
    MsgBox Err.Description
End Sub

Private Sub Command2_Click() 'REMOVE QUERY FROM DATABASE
Dim qdfrem As QueryDef
Dim qdfname As String
Dim msg, style, title, response
If QryList.ListIndex = -1 Then
   MsgBox ("Please select a query from the list below")
Else
  qdfname = DB.QueryDefs(QryList.ListIndex).Name
  msg = "Query to be deleted: "
  style = vbOKCancel
  title = "QUERY DELETE"
  response = MsgBox(msg & qdfname, style, title)
  If response = vbOK Then   ' User chose OK
     QryList.RemoveItem (QryList.ListIndex)
     DB.QueryDefs.Delete (qdfname)
     QryList.refresh
  Else
  End If
End If
End Sub

Private Sub Command3_Click() ' Clear txtSQL in box
txtSQL.Text = ""
End Sub

Private Sub Command5_Click() 'RESET
MSFlexGrid1.Clear
End Sub

Private Sub data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'This will display the current record position for dynasets and snapshots
  Data1.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1)
End Sub

Private Sub End_Click()
Unload Me
End Sub

Private Sub ExecSQL_Click()
On Error GoTo SQLError
       
    
  'Only for Counting initial amount of records
    Data1.RecordSource = TblList.List(0)
    Data1.refresh
    Data1.Recordset.MoveLast
    Data1.Recordset.MoveFirst
    initialreccount = Data1.Recordset.RecordCount
    'Counting records & percentage of records
    Data1.RecordSource = txtSQL
    Data1.refresh
    Data1.Recordset.MoveLast
    Data1.Recordset.MoveFirst
    reccount.Caption = Data1.Recordset.RecordCount
    percent.Caption = Int((Data1.Recordset.RecordCount * 100) / initialreccount)
    Exit Sub
    
SQLError:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
statistical_processing = 0
linegraphprocessing = 0
piegraphprocessing = 0
End Sub

Private Sub LineGraphMenu_Click()
linegraphprocessing = 1
Load LineGraph
LineGraph.Show
End Sub

Private Sub LinkToImageProcForm_Click()
Load ImageProc
ImageProc.Show
ImageProc.SetFocus
End Sub

Private Sub PieGraphMenu_Click()
piegraphprocessing = 1
Load PieGraph
PieGraph.Show
End Sub

Private Sub MSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If statistical_processing And Not linegraphprocessing And Not piegraphprocessing Then
  StatForm.Show
  StatForm.SetFocus
  Else
  If linegraphprocessing And Not statistical_processing And Not piegraphprocessing Then
  LineGraph.Show
  LineGraph.SetFocus
  Else
  If piegraphprocessing And Not statistical_processing And Not linegraphprocessing Then
  PieGraph.Show
  PieGraph.SetFocus
  End If
 End If
End If
End Sub

Private Sub QryList_Click()
Dim qry As QueryDef
    textSQL.Text = DB.QueryDefs(QryList.ListIndex).SQL
End Sub

Private Sub QryList_DblClick()
txtSQL.Text = textSQL.Text
Queryname.Caption = DB.QueryDefs(QryList.ListIndex).Name
End Sub

Private Sub Statmenu_Click()
Load StatForm
statistical_processing = 1
StatForm.Show
End Sub

Private Sub TblList_Click()
Dim fld As Field
Dim idx As Index
    If Left(TblList.Text, 2) = "  " Then Exit Sub
    FldList.Clear
    For Each fld In DB.TableDefs(TblList.Text).Fields
        FldList.AddItem fld.Name
    Next
End Sub
Private Sub OpenDB_Click()
On Error GoTo NoDatabase
CommonDialog1.CancelError = True
CommonDialog1.Filter = "Databases|*.MDB"
CommonDialog1.ShowOpen

'Initializations
'Data1.DatabaseName = Null
'Data1.RecordSource = Null
'TblList.Clear
'FldList.Clear
'QryList.Clear

Data1.DatabaseName = CommonDialog1.FileName
Data1.refresh
' Open the database
    If CommonDialog1.FileName <> "" Then
        Set DB = OpenDatabase(CommonDialog1.FileName)
        Label1.Caption = CommonDialog1.FileName
    End If
' Clear the ListBox controls
TblList.Clear
FldList.Clear
QryList.Clear
txtSQL.Text = ""
textSQL.Text = ""

For Each tbl In DB.TableDefs
        ' EXCLUDE SYSTEM TABLES
 If Left(tbl.Name, 4) <> "MSys" And Left(tbl.Name, 4) <> "USys" Then
        TblList.AddItem tbl.Name
 End If
Next
    
Debug.Print "There are " & DB.QueryDefs.Count & " queries in the database"
' Process each stored query
    For Each qry In DB.QueryDefs
        QryList.AddItem qry.Name
    Next
    
If Err = 0 Then
        Label1.Caption = CommonDialog1.FileName
    Else
        MsgBox Err.Description
End If

NoDatabase:
    On Error GoTo 0
End Sub
