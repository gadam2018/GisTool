VERSION 5.00
Begin VB.MDIForm GisTool 
   BackColor       =   &H8000000C&
   Caption         =   "GisTool"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   Picture         =   "GisTool.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu CallImageProcForm 
      Caption         =   "Image Processing"
   End
   Begin VB.Menu CallManDataForm 
      Caption         =   "Database Processing"
   End
   Begin VB.Menu QuitGisTool 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "GisTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CallImageProcForm_Click()
Load ImageProc
ImageProc.Show
End Sub

Private Sub CallManDataForm_Click()
Load ManData
ManData.Show
End Sub

Private Sub QuitGisTool_Click()
Unload Me
End
End Sub
