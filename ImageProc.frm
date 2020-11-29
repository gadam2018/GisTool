VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ImageProc 
   Caption         =   "Image Processing"
   ClientHeight    =   6255
   ClientLeft      =   3090
   ClientTop       =   2430
   ClientWidth     =   6150
   LinkTopic       =   "GisTool"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   417
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   410
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6015
      Left            =   120
      Picture         =   "ImageProc.frx":0000
      ScaleHeight     =   397
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   390
      TabIndex        =   0
      Top             =   120
      Width           =   5910
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6045
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu FileMenu 
      Caption         =   "Image"
      Begin VB.Menu FileOpen 
         Caption         =   "Open Image"
      End
      Begin VB.Menu FileSave 
         Caption         =   "Save Image"
      End
      Begin VB.Menu Close 
         Caption         =   "Close Image"
      End
      Begin VB.Menu FileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Processing 
      Caption         =   "Image Processing"
      Begin VB.Menu ImageEnhancement 
         Caption         =   "Image Enhancement"
         Begin VB.Menu ProcessSmooth 
            Caption         =   "Smooth"
         End
         Begin VB.Menu ProcessSharpen 
            Caption         =   "Sharpen"
         End
         Begin VB.Menu Threshold 
            Caption         =   "Threshold"
         End
         Begin VB.Menu Diff 
            Caption         =   "Differentiate"
         End
      End
      Begin VB.Menu ImageIdentification 
         Caption         =   "ImageIdentification"
         Begin VB.Menu Invert 
            Caption         =   "Invert"
         End
         Begin VB.Menu Mag 
            Caption         =   "Magnification"
         End
         Begin VB.Menu Centroids 
            Caption         =   "Centroids"
         End
         Begin VB.Menu Rotate 
            Caption         =   "Rotate"
         End
      End
      Begin VB.Menu ImageFilters 
         Caption         =   "Filters"
         Begin VB.Menu Median 
            Caption         =   "Median"
         End
      End
      Begin VB.Menu ImageRestoration 
         Caption         =   "Image Restoration"
      End
   End
   Begin VB.Menu Links 
      Caption         =   "Links"
      Begin VB.Menu MSExcel 
         Caption         =   "MS Excel"
      End
      Begin VB.Menu SPSS 
         Caption         =   "SPSS"
      End
      Begin VB.Menu Arcview 
         Caption         =   "ESRI ArcView"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
   End
   Begin VB.Menu GisDataAnalysis 
      Caption         =   "GoToDatabaseProcessingForm"
   End
   Begin VB.Menu End 
      Caption         =   "End"
   End
End
Attribute VB_Name = "ImageProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Close_Click()
Unload Me
Load ImageProc
ImageProc.Show
End Sub

Private Sub End_Click()
Unload Me
End Sub

Private Sub FileOpen_Click()
Dim i As Integer, j As Integer
Dim red As Integer, green As Integer, blue As Integer
Dim pixel As Long
Dim PictureName As String
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "Images|*.BMP;*.GIF;*.JPG;*.DIB|All Files|*.*"
    CommonDialog1.Action = 1
    PictureName = CommonDialog1.FileName
    If PictureName = "" Then Exit Sub
    Picture1.Picture = LoadPicture(PictureName)
    ImageProc.refresh

    X = Picture1.ScaleWidth
    Y = Picture1.ScaleHeight
    If X > 1000 Or Y > 1000 Then
        MsgBox "Image too large to process. Please try loading a smaller image."
        X = 1000
        Y = 1000
       ' Exit Sub
    End If

    ImageProc.Width = ImageProc.ScaleX(Picture1.Width + 6, vbPixels, vbTwips)
    ImageProc.Height = ImageProc.ScaleY(Picture1.Height + 30, vbPixels, vbTwips)
    ImageProc.refresh
    ImgLoad.Show
    ImgLoad.refresh
    For i = 0 To Y - 1
        For j = 0 To X - 1
            pixel = ImageProc.Picture1.Point(j, i)
            red = pixel& Mod 256
            green = ((pixel And &HFF00) / 256&) Mod 256&
            blue = (pixel And &HFF0000) / 65536
            ImagePixels(0, i, j) = red
            ImagePixels(1, i, j) = green
            ImagePixels(2, i, j) = blue
        Next
        'Form3.ProgressBar1.Value = i * 100 / (Y - 1)
    Next
    ImgLoad.Hide
End Sub

Private Sub FileSave_Click()
Dim PictureName As String
    CommonDialog1.Action = 2
    PictureName = CommonDialog1.FileName
    SavePicture Picture1.Image, PictureName
End Sub

Private Sub Form_Load()
ImageProc.Width = 6270
ImageProc.Height = 6945
End Sub

Private Sub GisDataAnalysis_Click()
Load ManData
ManData.Show
ManData.SetFocus
End Sub

Private Sub Invert_Click()
Dim i As Integer, j As Integer
Dim red As Integer, green As Integer, blue As Integer
    
    For i = 1 To Y - 2
        For j = 1 To X - 2
            red = ImagePixels(0, i, j)
            green = ImagePixels(1, i, j)
            blue = ImagePixels(2, i, j)
            Picture1.PSet (i, j), RGB(red, green, blue)
        Next
        Picture1.refresh
    Next
End Sub

Private Sub ProcessSharpen_Click()
Dim i As Integer, j As Integer
Const Dx As Integer = 1
Const Dy As Integer = 1
Dim red As Integer, green As Integer, blue As Integer
    
    For i = 1 To Y - 2
        For j = 1 To X - 2
            red = ImagePixels(0, i, j) + 0.5 * (ImagePixels(0, i, j) - ImagePixels(0, i - Dx, j - Dy))
            green = ImagePixels(1, i, j) + 0.5 * (ImagePixels(1, i, j) - ImagePixels(1, i - Dx, j - Dy))
            blue = ImagePixels(2, i, j) + 0.5 * (ImagePixels(2, i, j) - ImagePixels(2, i - Dx, j - Dy))
            If red > 255 Then red = 255
            If red < 0 Then red = 0
            If green > 255 Then green = 255
            If green < 0 Then green = 0
            If blue > 255 Then blue = 255
            If blue < 0 Then blue = 0
            Picture1.PSet (j, i), RGB(red, green, blue)
        Next
        Picture1.refresh
    Next
End Sub

Private Sub ProcessSmooth_Click()
Dim i As Integer, j As Integer
Dim red As Integer, green As Integer, blue As Integer

    For i = 1 To Y - 2
        For j = 1 To X - 2
            red = ImagePixels(0, i - 1, j - 1) + ImagePixels(0, i - 1, j) + ImagePixels(0, i - 1, j + 1) + _
            ImagePixels(0, i, j - 1) + ImagePixels(0, i, j) + ImagePixels(0, i, j + 1) + _
            ImagePixels(0, i + 1, j - 1) + ImagePixels(0, i + 1, j) + ImagePixels(0, i + 1, j + 1)
            
            green = ImagePixels(1, i - 1, j - 1) + ImagePixels(1, i - 1, j) + ImagePixels(1, i - 1, j + 1) + _
            ImagePixels(1, i, j - 1) + ImagePixels(1, i, j) + ImagePixels(1, i, j + 1) + _
            ImagePixels(1, i + 1, j - 1) + ImagePixels(1, i + 1, j) + ImagePixels(1, i + 1, j + 1)
            
            blue = ImagePixels(2, i - 1, j - 1) + ImagePixels(2, i - 1, j) + ImagePixels(2, i - 1, j + 1) + _
            ImagePixels(2, i, j - 1) + ImagePixels(2, i, j) + ImagePixels(2, i, j + 1) + _
            ImagePixels(2, i + 1, j - 1) + ImagePixels(2, i + 1, j) + ImagePixels(2, i + 1, j + 1)

            Picture1.PSet (j, i), RGB(red / 9, green / 9, blue / 9)
        Next
        Picture1.refresh
    Next
End Sub

Private Sub Threshold_Click()
Dim i As Integer, j As Integer
Dim red As Integer, green As Integer, blue As Integer
For i = 1 To Y - 2
        For j = 1 To X - 2
            red = ImagePixels(0, i, j)
            green = ImagePixels(1, i, j)
            blue = ImagePixels(2, i, j)
            If red >= 128 Or green >= 128 Or blue >= 128 Then
               red = 255
               green = 255
               blue = 255
            Else
               red = 0
               green = 0
               blue = 0
            End If
            'If red < 128 Then red = 0
            'If green >= 128 Then green = 255
            'If green < 128 Then green = 0
            'If blue >= 128 Then blue = 255
            'If blue < 128 Then blue = 0
            Picture1.PSet (j, i), RGB(red, green, blue)
        Next
        Picture1.refresh
    Next
End Sub
Private Sub FileExit_Click()
    End
End Sub

Private Sub Tools_Click()

End Sub
