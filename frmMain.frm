VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gallery"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2400
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   2400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAlwaysOnTop 
      Caption         =   "On Top"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   2280
      Width           =   855
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2025
      ScaleWidth      =   2385
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdNext 
      Height          =   375
      Left            =   1080
      Picture         =   "frmMain.frx":10CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdPrev 
      Height          =   375
      Left            =   600
      Picture         =   "frmMain.frx":129B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cdFolder 
      Left            =   600
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSelectFolder 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   0
      Picture         =   "frmMain.frx":147F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   495
   End
   Begin VB.Timer tmrImage 
      Interval        =   5000
      Left            =   120
      Top             =   3120
   End
   Begin VB.Label lblName 
      Caption         =   "Select Image to start"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   2760
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ImageFiles As Collection
Private CurrentImageIndex As Integer
Private m_bAlwaysOnTop As Boolean
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Private Sub cmdSelectFolder_Click()
    Dim folderPath As String
    folderPath = BrowseForFolder("Select Image Folder")
    
    If folderPath <> "" Then
        LoadImageFiles folderPath
        If ImageFiles.Count > 0 Then
            CurrentImageIndex = 1
            DisplayCurrentImage
            tmrImage.Enabled = True
        Else
            MsgBox "No image files found in selected folder"
        End If
    End If
End Sub

Private Sub cmdNext_Click()
    If ImageFiles.Count = 0 Then Exit Sub

    CurrentImageIndex = CurrentImageIndex + 1
    If CurrentImageIndex > ImageFiles.Count Then
        CurrentImageIndex = 1 ' wrap to first image
    End If

    DisplayCurrentImage
End Sub

Private Sub cmdPrev_Click()
    If ImageFiles.Count = 0 Then Exit Sub

    CurrentImageIndex = CurrentImageIndex - 1
    If CurrentImageIndex < 1 Then
        CurrentImageIndex = ImageFiles.Count ' wrap to last image
    End If

    DisplayCurrentImage
End Sub

Private Sub cmdAlwaysOnTop_Click()
    m_bAlwaysOnTop = Not m_bAlwaysOnTop
    
    If m_bAlwaysOnTop Then
        ' Set window always on top
        SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
        cmdAlwaysOnTop.Caption = "Off Top"
    Else
        ' Remove always on top
        SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 3
        cmdAlwaysOnTop.Caption = "On Top"
    End If
End Sub

Private Sub Form_Load()
    Set ImageFiles = New Collection
    tmrImage.Enabled = False
End Sub

Private Sub tmrImage_Timer()
    If ImageFiles.Count > 0 Then
        CurrentImageIndex = CurrentImageIndex + 1
        If CurrentImageIndex > ImageFiles.Count Then
            CurrentImageIndex = 1
        End If
        DisplayCurrentImage
    End If
End Sub

Private Sub LoadImageFiles(folderPath As String)
    Set ImageFiles = New Collection
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    For Each file In folder.Files
        Select Case LCase(fso.GetExtensionName(file.Name))
            Case "jpg", "jpeg", "png", "gif", "bmp"
                ImageFiles.Add file.Path
        End Select
    Next file
End Sub

Private Sub DisplayCurrentImage()
    On Error GoTo ErrorHandler
    Dim img As StdPicture
    Dim imgWidth As Long, imgHeight As Long
    Dim destWidth As Long, destHeight As Long
    Dim offsetX As Long, offsetY As Long
    Dim ratio As Double

    If CurrentImageIndex >= 1 And CurrentImageIndex <= ImageFiles.Count Then
        Set img = LoadPicture(ImageFiles(CurrentImageIndex))
        If img Is Nothing Then Exit Sub
        
        ' Show current file name in label
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        lblName.Caption = fso.GetFileName(ImageFiles(CurrentImageIndex))

        ' Get original image dimensions (in HIMETRIC units ¡ú convert to twips)
        imgWidth = ScaleX(img.Width, vbHimetric, vbTwips)
        imgHeight = ScaleY(img.Height, vbHimetric, vbTwips)

        ' Calculate aspect ratio
        ratio = imgWidth / imgHeight

        ' Fit by height first
        destHeight = picImage.ScaleHeight
        destWidth = destHeight * ratio

        ' If too wide, fit by width instead
        If destWidth > picImage.ScaleWidth Then
            destWidth = picImage.ScaleWidth
            destHeight = destWidth / ratio
        End If

        ' Center aligned
        offsetX = (picImage.ScaleWidth - destWidth) / 2
        offsetY = (picImage.ScaleHeight - destHeight) / 2

        ' Draw scaled image to PictureBox canvas
        picImage.Cls
        picImage.AutoRedraw = True
        picImage.PaintPicture img, offsetX, offsetY, destWidth, destHeight
    End If
    Exit Sub

ErrorHandler:
    Resume Next
End Sub

Private Function BrowseForFolder(Optional Title As String = "Select Folder") As String
    ' You'll need to implement folder browser functionality
    ' This can be done using API calls or Shell.Application
    ' For simplicity, using CommonDialog for file selection
    cdFolder.DialogTitle = Title
    cdFolder.ShowOpen
    If cdFolder.FileName <> "" Then
        BrowseForFolder = Left(cdFolder.FileName, InStrRev(cdFolder.FileName, "\") - 1)
    Else
        BrowseForFolder = ""
    End If
End Function
