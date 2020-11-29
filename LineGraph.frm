VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Begin VB.Form LineGraph 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LineGraph"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   277
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton draw 
      Caption         =   "Draw line"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Return 
      Caption         =   "Return"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton refresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3720
      Width           =   975
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   0
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   720
      ScaleHeight     =   42.287
      ScaleMode       =   0  'User
      ScaleWidth      =   72.219
      TabIndex        =   0
      Top             =   600
      Width           =   4485
   End
   Begin VB.Line Line2 
      X1              =   48
      X2              =   48
      Y1              =   200
      Y2              =   216
   End
   Begin VB.Label Label5 
      Caption         =   "y"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   3000
      Width           =   255
   End
   Begin VB.Line Line1 
      X1              =   368
      X2              =   344
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Label Label4 
      Caption         =   "x"
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "X="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1275
   End
   Begin VB.Label Label8 
      Caption         =   "Y="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4320
      TabIndex        =   3
      Top             =   3120
      Width           =   1515
   End
   Begin VB.Label Label3 
      Caption         =   "Y"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "X"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "LineGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim sortedmatrix() As Double
Dim unsortedmatrix() As Double
Dim k As Integer
Dim currX, currY As Single
Dim XMin As Double, XMax As Double, YMin As Double, YMax As Double
Function FunctionEval1(ByVal X As Long) As Double
    ScriptControl1.ExecuteStatement "X=" & X
    FunctionEval1 = X
End Function

Private Sub draw_Click()
Dim t  As Double
Dim XPixels As Integer
'GraphForm.ScaleMode = 6
'GraphForm.Scale (0, 0)-(80, 24)
   griddata
   YMin = sortedmatrix(1) '0
    YMax = sortedmatrix(k) '24
    XMin = sortedmatrix(1) '0
    XMax = sortedmatrix(k) '80
    Picture1.Cls
    'makeframe
    Picture1.ScaleMode = 6
    Picture1.Scale (XMin, YMin)-(XMax, YMax)
    show_labels
    Picture1.ForeColor = RGB(0, 0, 255)

'Picture1.PSet (unsortedmatrix(1), FunctionEval1(unsortedmatrix(1)))
For i = 2 To k
    t = sortedmatrix(1) + (sortedmatrix(k) - sortedmatrix(1)) * i / k 'unsortedmatrix(k) 't = XMin + (XMax - XMin) * i / XPixels
    Picture1.Line -(t, FunctionEval1(t))
    
Next
End Sub

Private Sub Form_Load()
'GraphForm.ScaleMode = 6
'Picture1.ScaleMode = 6
'GraphForm.Scale (0, 0)-(80, 24)
'Picture1.Scale (0, 0)-(80, 24)
End Sub

Private Sub show_labels()
Label2.Caption = "X " & Format$(Picture1.ScaleX(currX, Picture1.ScaleMode, 1), "#.000")
Label3.Caption = "Y " & Format$(Picture1.ScaleY(currY, Picture1.ScaleMode, 1), "#.000")
Label7.Caption = "(" & Picture1.ScaleLeft & ", " & Picture1.ScaleTop & ")"
Label8.Caption = "(" & Picture1.ScaleLeft + Picture1.ScaleWidth & ", " & Picture1.ScaleTop + Picture1.ScaleHeight & ")"
End Sub

Private Sub make_frame_Click()
makeframe
End Sub

Private Sub makeframe()
For i = 2 To XMax Step 2 'Making Upper & Lower scales
'If i >= YMax / 2 Then 'Or i = 40 Or i = 60 Then
'Picture1.Line (i, 0)-(i, 2)  'Pic1 Upper scale main points 20,40,60
'Picture1.Line (i, 24)-(i, 22) 'Pic1 Lower scale main points 20,40,60
'Else
Picture1.Line (i, 0)-(i, 1)   'Pic1 Upper frame scale
Picture1.Line (i, YMax)-(i, YMax - 1) 'Pic1 Lower frame scale
'End If
Next
For i = 2 To YMax Step 2
'If i = 12 Then
'Picture1.Line (0, i)-(2, i) 'Pic1 Left scale main point 12
'Picture1.Line (80, i)-(78, i) 'Pic1 Right scale main point 12
'Else
Picture1.Line (0, i)-(1, i) 'Pic1 Left frame scale
Picture1.Line (XMax, i)-(XMax - 1, i) 'Pic1 Right frame scale
'End If
Next
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    'Picture1.refresh
    'Picture1.Line (X, Picture1.ScaleTop)-(X, Picture1.ScaleTop + Picture1.ScaleHeight)
    'Picture1.Line (Picture1.ScaleLeft, Y)-(Picture1.ScaleLeft + Picture1.ScaleWidth, Y)
    Label2.Caption = "X " & Format$(X, "#.000")
    Label3.Caption = "Y " & Format$(Y, "#.000")
    currX = X
    currY = Y
End Sub

Private Sub refresh_Click()
Picture1.refresh
End Sub

Private Sub return_Click()
ManData.linegraphprocessing = 0
Unload Me
End Sub

Private Sub griddata()
totalsum = 0
    For j = ManData.MSFlexGrid1.Col To ManData.MSFlexGrid1.ColSel
     For i = ManData.MSFlexGrid1.Row To ManData.MSFlexGrid1.RowSel
      If IsNumeric(ManData.MSFlexGrid1.TextMatrix(i, j)) Then
        totalsum = totalsum + ManData.MSFlexGrid1.TextMatrix(i, j)
      Else
       MsgBox ("Κάνε την επιλογή των δεδομένων")
       Exit Sub
      End If
     Next
    Next
   ' Total number of rows, total number of columns, total number of cells
     totalrows = (ManData.MSFlexGrid1.RowSel - ManData.MSFlexGrid1.Row) + 1
     totalcols = (ManData.MSFlexGrid1.ColSel - ManData.MSFlexGrid1.Col) + 1
     totalcells = totalrows * totalcols
     
     ReDim sortedmatrix(totalcells)
     ReDim unsortedmatrix(totalcells)
     'assigning the values of TextMatrix() array to new arrays
     k = 1
     For j = ManData.MSFlexGrid1.Col To ManData.MSFlexGrid1.ColSel
     For i = ManData.MSFlexGrid1.Row To ManData.MSFlexGrid1.RowSel
        'MsgBox ManData.MSFlexGrid1.TextMatrix(i, j)
        sortedmatrix(k) = ManData.MSFlexGrid1.TextMatrix(i, j)
        unsortedmatrix(k) = ManData.MSFlexGrid1.TextMatrix(i, j)
        'MsgBox "sortedmatrix[" & k & "]=" & sortedmatrix(k)
        k = k + 1
     Next
     Next
     k = k - 1
     'MsgBox "k=" & k
     'Sorting (increment) the array
     For i = 1 To k - 1
     For j = i + 1 To k
      If sortedmatrix(i) > sortedmatrix(j) Then
        temp = sortedmatrix(i)
        sortedmatrix(i) = sortedmatrix(j)
        sortedmatrix(j) = temp
      End If
      Next
      Next
End Sub

Private Sub ScriptControl1_Error()
    Debug.Print ScriptControl1.Error.Number
    Debug.Print ScriptControl1.Error.Text
End Sub
