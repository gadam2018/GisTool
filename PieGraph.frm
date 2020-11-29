VERSION 5.00
Begin VB.Form PieGraph 
   BackColor       =   &H00FFFFFF&
   Caption         =   "PieGraph"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton return 
      Caption         =   "Return"
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw pie"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   0
      Top             =   5520
      Width           =   1365
   End
End
Attribute VB_Name = "PieGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sortedmatrix() As Double
Dim unsortedmatrix() As Double
Dim k As Integer

Private Sub Command1_Click()
Dim piedata() As Double
  griddata
ReDim piedata(k) As Double
    PieGraph.Cls
    For i = 1 To k
        piedata(i) = unsortedmatrix(i) '+ Rnd() * 100
        Total = Total + piedata(i)
    Next
    
    PieGraph.DrawWidth = 2
    For i = 1 To k '9
        arc1 = arc2
        arc2 = arc1 + 6.28 * piedata(i) / Total
        'If Check1.Value Then
            PieGraph.FillStyle = 2 + (i Mod 5)
        'Else
            'PieGraph.FillStyle = 0
        'End If
       ' If Check2.Value Then
         '   PieGraph.FillColor = QBColor(8 + (i Mod 6))
        'Else
            PieGraph.FillColor = QBColor(7)
        'End If
        PieGraph.Circle (PieGraph.ScaleWidth / 2, PieGraph.ScaleHeight / 2), PieGraph.ScaleHeight / 2.5, , -arc1, -arc2
    Next
End Sub

Public Sub griddata()
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

Private Sub return_Click()
ManData.piegraphprocessing = 0
Unload Me
End Sub
