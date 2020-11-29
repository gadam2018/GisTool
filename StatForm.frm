VERSION 5.00
Begin VB.Form StatForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Statistical Processing"
   ClientHeight    =   3210
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5505
   Begin VB.CommandButton Command2 
      Caption         =   "Calculate"
      Height          =   372
      Left            =   2520
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   2880
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Return"
      Height          =   372
      Left            =   4440
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2292
   End
   Begin VB.Label reslabel 
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label resultlabel 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   5
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "StatForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sortedmatrix() As Double
Dim k As Integer

Private Sub Command1_Click() ' Επιστροφή
ManData.statistical_processing = 0
   Unload Me
End Sub

Private Sub Form_Load() 'Load my Stat Functions ()
  List1.AddItem "ArithmeticMean()"
  List1.AddItem "AverageDeviation()"
  List1.AddItem "Median()"
  'List1.AddItem "WeightedArithmeticMean()"
  'List1.AddItem "GeometricMean()"
  List1.AddItem "HarmonicMean()"
  'List1.AddItem "Mode()"
  'List1.AddItem "Range()"
  'List1.AddItem "MeanDeviation()"
  List1.AddItem "StandardDeviation()"
  'List1.AddItem "SquaresMean()"
  'List1.AddItem "Sampling()"
  'List1.AddItem "Ratio()"
  'List1.AddItem "Interval()"
  'List1.AddItem "NormalDistribution()"
  List1.refresh
  Label1.Visible = False
  'Text1.Visible = False
  'Text1.Text = ""
  Label1.Caption = ""
  resultlabel.Caption = ""
End Sub

Private Sub List1_Click()
'Text1.Text = ""
'Text1.Visible = False
Label1.Caption = ""
resultlabel.Caption = ""
Label1.Visible = False
reslabel.Visible = False
If List1.ListIndex = -1 Then
   MsgBox ("Select the function")
Else
   '***** IF Αριθμητικός Μέσος Function *****
   If List1.Text = "ArithmeticMean()" Then
   Me.Caption = "Calculating ArithmeticMean"
      Label1.Visible = True
      reslabel.Visible = True
      reslabel.Caption = "Result of Arithmentic Mean:"
      'Text1.Visible = True
      Label1.Caption = "Select the data"
   End If
   '***** IF Μέση Τιμή Function *****
   If List1.Text = "AverageDeviation()" Then
   Me.Caption = "Calculating AverageDeviation"
      Label1.Visible = True
      reslabel.Visible = True
      reslabel.Caption = "Result of Average Deviation:"
      'Text1.Visible = True
      Label1.Caption = "Select the data"
   End If
   '***** IF Διάμεση Τιμή Function *****
   If List1.Text = "Median()" Then
   Me.Caption = "Calculating Median"
      Label1.Visible = True
      reslabel.Visible = True
      reslabel.Caption = "Result of Median:"
      'Text1.Visible = True
      Label1.Caption = "Select the data"
   End If
   '***** IF Αρμονικός Μέσος Function *****
   If List1.Text = "HarmonicMean()" Then
   Me.Caption = "Calculating HarmonicMean"
      Label1.Visible = True
      reslabel.Visible = True
      reslabel.Caption = "Result of Harmonic Mean:"
      'Text1.Visible = True
      Label1.Caption = "Select the data"
   End If
    '***** IF Τυπική Απόκλιση Function *****
   If List1.Text = "StandardDeviation()" Then
   Me.Caption = "Calculating StandardDeviation"
      Label1.Visible = True
      reslabel.Visible = True
      reslabel.Caption = "Result of Standard Deviation:"
      'Text1.Visible = True
      Label1.Caption = "Select the data"
   End If
   '***** IF GeometricMean Function *****
   If List1.Text = "GeometricMean()" Then
      Label1.Visible = True
      reslabel.Visible = True
      reslabel.Caption = "Result of Geometric mean:"
      'Text1.Visible = True
      Label1.Caption = "Please select the numeric data cells."
   End If
End If
End Sub

Private Sub Command2_Click() 'Calculate Function
'********** IF Αριθμητικός Μέσος Function *********
If List1.Text = "ArithmeticMean()" Then
    'arg1 = Text1.Text
    totalsum = 0
    For j = ManData.MSFlexGrid1.Col To ManData.MSFlexGrid1.ColSel
     For i = ManData.MSFlexGrid1.Row To ManData.MSFlexGrid1.RowSel
      If IsNumeric(ManData.MSFlexGrid1.TextMatrix(i, j)) Then
        totalsum = totalsum + ManData.MSFlexGrid1.TextMatrix(i, j)
      Else
       MsgBox ("Select the data")
       Exit Sub
      End If
     Next
    Next
   ' Total number of rows, total number of columns, total number of cells
     totalrows = (ManData.MSFlexGrid1.RowSel - ManData.MSFlexGrid1.Row) + 1
     totalcols = (ManData.MSFlexGrid1.ColSel - ManData.MSFlexGrid1.Col) + 1
     totalcells = totalrows * totalcols
     arithmean = Round(totalsum / totalcells, 4)
     'MsgBox (totalsum)
     resultlabel.Caption = arithmean
    'arg2 = ManData.MSFlexGrid1.Clip
    'MsgBox (arg2)
     'arg2 = ManData.MSFlexGrid1.Text
    'If arg1 <> "" And IsNumeric(arg1) Then  ' Show the result
        'MsgBox (ArithmeticMean(arg1))
    'Else
        'If arg2 <> "" And IsNumeric(arg2) Then
           ' MsgBox (ArithmeticMean(arg2))
       ' End If
    'End If
   ' Text1.Text = ""
    'arg1 = Null
    'arg2 = Null
    On Error GoTo CalcError
    Exit Sub
 Me.Caption = "Statistical Processing"
 End If
 '********** IF Μέση Τιμή Function *********
If List1.Text = "AverageDeviation()" Then
    'arg1 = Text1.Text
    totalsum = 0
    For j = ManData.MSFlexGrid1.Col To ManData.MSFlexGrid1.ColSel
     For i = ManData.MSFlexGrid1.Row To ManData.MSFlexGrid1.RowSel
      If IsNumeric(ManData.MSFlexGrid1.TextMatrix(i, j)) Then
        totalsum = totalsum + ManData.MSFlexGrid1.TextMatrix(i, j)
      Else
       MsgBox ("Select the data")
       Exit Sub
      End If
     Next
    Next
   ' Total number of rows, total number of columns, total number of cells
     totalrows = (ManData.MSFlexGrid1.RowSel - ManData.MSFlexGrid1.Row) + 1
     totalcols = (ManData.MSFlexGrid1.ColSel - ManData.MSFlexGrid1.Col) + 1
     totalcells = totalrows * totalcols
     arithmean = Round(totalsum / totalcells, 4)
     
    totalabssum = 0
    For j = ManData.MSFlexGrid1.Col To ManData.MSFlexGrid1.ColSel
     For i = ManData.MSFlexGrid1.Row To ManData.MSFlexGrid1.RowSel
      If IsNumeric(ManData.MSFlexGrid1.TextMatrix(i, j)) Then
        totalabssum = totalabssum + Abs(ManData.MSFlexGrid1.TextMatrix(i, j) - arithmean)
      Else
       MsgBox ("Select the data")
       Exit Sub
      End If
     Next
    Next
     avedev = Round(totalabssum / totalcells, 4)
     'MsgBox (totalsum)
     resultlabel.Caption = avedev
    On Error GoTo CalcError
    Exit Sub
 End If
 '********** IF Διάμεση Τιμή Function *********
If List1.Text = "Median()" Then
    'arg1 = Text1.Text
    totalsum = 0
    For j = ManData.MSFlexGrid1.Col To ManData.MSFlexGrid1.ColSel
     For i = ManData.MSFlexGrid1.Row To ManData.MSFlexGrid1.RowSel
      If IsNumeric(ManData.MSFlexGrid1.TextMatrix(i, j)) Then
        totalsum = totalsum + ManData.MSFlexGrid1.TextMatrix(i, j)
      Else
       MsgBox ("Select the data")
       Exit Sub
      End If
     Next
    Next
   ' Total number of rows, total number of columns, total number of cells
     totalrows = (ManData.MSFlexGrid1.RowSel - ManData.MSFlexGrid1.Row) + 1
     totalcols = (ManData.MSFlexGrid1.ColSel - ManData.MSFlexGrid1.Col) + 1
     totalcells = totalrows * totalcols
     
     ReDim sortedmatrix(totalcells)
     'assigning the values of TextMatrix() array to a new array
     k = 1
     For j = ManData.MSFlexGrid1.Col To ManData.MSFlexGrid1.ColSel
     For i = ManData.MSFlexGrid1.Row To ManData.MSFlexGrid1.RowSel
        'MsgBox ManData.MSFlexGrid1.TextMatrix(i, j)
        sortedmatrix(k) = ManData.MSFlexGrid1.TextMatrix(i, j)
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
     If (totalcells Mod 2) <> 0 Then
        'perittos arithmos of cells
         Median = sortedmatrix((Int(k / 2) + 1)) 'ManData.MSFlexGrid1.TextMatrix(Round(totalrows / 2), Round(totalcols / 2))
      Else 'artios arithmos
        Median = (sortedmatrix(Int(k / 2)) + sortedmatrix((Int(k / 2) + 1))) / 2 '(ManData.MSFlexGrid1.TextMatrix(totalrows / 2, totalcols / 2) + ManData.MSFlexGrid1.TextMatrix(((totalrows / 2) + 1), ((totalcols / 2) + 1))) / 2
    End If
     resultlabel.Caption = Median
    On Error GoTo CalcError
    Exit Sub
 End If
 
 '********** IF Αρμονικός Μέσος Function *********
If List1.Text = "HarmonicMean()" Then
    'arg1 = Text1.Text
    inversetotalsum = 0
    For j = ManData.MSFlexGrid1.Col To ManData.MSFlexGrid1.ColSel
     For i = ManData.MSFlexGrid1.Row To ManData.MSFlexGrid1.RowSel
      If IsNumeric(ManData.MSFlexGrid1.TextMatrix(i, j)) Then
        inversetotalsum = inversetotalsum + (1 / ManData.MSFlexGrid1.TextMatrix(i, j))
      Else
       MsgBox ("Select the data")
       Exit Sub
      End If
     Next
    Next
   ' Total number of rows, total number of columns, total number of cells
     totalrows = (ManData.MSFlexGrid1.RowSel - ManData.MSFlexGrid1.Row) + 1
     totalcols = (ManData.MSFlexGrid1.ColSel - ManData.MSFlexGrid1.Col) + 1
     totalcells = totalrows * totalcols
     Harmonicmean = 1 / (inversetotalsum / totalcells)
     resultlabel.Caption = Harmonicmean
    On Error GoTo CalcError
    Exit Sub
 End If
 '********** IF Τυπική Απόκλιση Function *********
If List1.Text = "StandardDeviation()" Then
    'arg1 = Text1.Text
    totalsum = 0
    totalsquaresum = 0
    For j = ManData.MSFlexGrid1.Col To ManData.MSFlexGrid1.ColSel
     For i = ManData.MSFlexGrid1.Row To ManData.MSFlexGrid1.RowSel
      If IsNumeric(ManData.MSFlexGrid1.TextMatrix(i, j)) Then
        totalsum = totalsum + ManData.MSFlexGrid1.TextMatrix(i, j)
        totalsquaresum = totalsquaresum + (ManData.MSFlexGrid1.TextMatrix(i, j) * ManData.MSFlexGrid1.TextMatrix(i, j))
      Else
       MsgBox ("Select the data")
       Exit Sub
      End If
     Next
    Next
   ' Total number of rows, total number of columns, total number of cells
     totalrows = (ManData.MSFlexGrid1.RowSel - ManData.MSFlexGrid1.Row) + 1
     totalcols = (ManData.MSFlexGrid1.ColSel - ManData.MSFlexGrid1.Col) + 1
     totalcells = totalrows * totalcols
     stddev = Sqr(((totalcells * totalsquaresum) - (totalsum * totalsum)) / (totalcells * (totalcells - 1)))
     resultlabel.Caption = stddev
    On Error GoTo CalcError
    Exit Sub
 End If
 
'********** IF GeometricMean Function *********
If List1.Text = "GeometricMean()" Then
Set ExcelApp = CreateObject("Excel.Application")
If ExcelApp Is Nothing Then
    MsgBox ("Could not start Excel")
    Exit Sub
End If
totalproduct = 1
    For j = ManData.MSFlexGrid1.Col To ManData.MSFlexGrid1.ColSel
     For i = ManData.MSFlexGrid1.Row To ManData.MSFlexGrid1.RowSel
      If IsNumeric(ManData.MSFlexGrid1.TextMatrix(i, j)) Then
        totalproduct = totalproduct * ManData.MSFlexGrid1.TextMatrix(i, j)
      Else
       MsgBox ("Please select numeric data")
       Exit Sub
      End If
     Next
    Next
   ' Total number of rows, total number of columns, total number of cells
     totalrows = (ManData.MSFlexGrid1.RowSel - ManData.MSFlexGrid1.Row) + 1
     totalcols = (ManData.MSFlexGrid1.ColSel - ManData.MSFlexGrid1.Col) + 1
     totalcells = totalrows * totalcols
     ekthetis = 1 / totalcells
     'geometricmean = Power(totalproduct, ekthetis)
     'geometricmean = ExcelApp.Evaluate("Power(" & totalproduct & "," & ekthetis & ")")
     'resultlabel.Caption = Round(geometricmean, 2)
    
    On Error GoTo CalcError
    Exit Sub
End If

CalcError:
    MsgBox ("Error:" & vbCrLf & Err.Description)
End Sub

