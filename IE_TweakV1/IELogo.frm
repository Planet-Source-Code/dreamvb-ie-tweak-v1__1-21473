VERSION 5.00
Begin VB.Form FrmLogoChanger 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open new IE Logo"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   5085
      TabIndex        =   11
      Top             =   3435
      Width           =   1155
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   360
      Left            =   3885
      TabIndex        =   10
      Top             =   3435
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Stop"
      Height          =   390
      Left            =   5235
      TabIndex        =   9
      Top             =   1560
      Width           =   900
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6645
      Top             =   3525
   End
   Begin VB.PictureBox Srcbitmap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   975
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   4140
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "&Play"
      Height          =   390
      Left            =   5235
      TabIndex        =   7
      Top             =   1080
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   5130
      TabIndex        =   3
      Top             =   -30
      Width           =   1155
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   1620
         Left            =   105
         ScaleHeight     =   108
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   62
         TabIndex        =   5
         Top             =   345
         Width           =   930
         Begin VB.PictureBox Preview 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   300
            ScaleHeight     =   330
            ScaleWidth      =   330
            TabIndex        =   6
            Top             =   105
            Width           =   330
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            X1              =   19
            X2              =   43
            Y1              =   29
            Y2              =   29
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            X1              =   42
            X2              =   42
            Y1              =   7
            Y2              =   29
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFFFFF&
            Height          =   360
            Left            =   285
            Top             =   90
            Width           =   360
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preview"
         Height          =   195
         Left            =   285
         TabIndex        =   4
         Top             =   120
         Width           =   570
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   0
      Pattern         =   "*.BMP"
      TabIndex        =   2
      Top             =   2130
      Width           =   6315
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   30
      TabIndex        =   1
      Top             =   420
      Width           =   5070
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4170
   End
   Begin VB.Label lblsize 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4305
      TabIndex        =   12
      Top             =   90
      Width           =   45
   End
End
Attribute VB_Name = "FrmLogoChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Const MaxWidth = 22
Dim ICount As Integer
Dim TBitmap As BITMAP

Function GetBitmapInfo()
Dim hFile As Long
    hFile = FreeFile
    On Error Resume Next
    Open IELogo_Picture For Binary As #hFile
        Get #hFile, 19, TBitmap.bmWidth
        Get #hFile, 23, TBitmap.bmHeight
    Close #1
    If Err Then MsgBox Err.Description, vbInformation
    
End Function
Private Sub cmdCancel_Click()
    IELogo_Picture = ""
    
End Sub

Private Sub cmdOpen_Click()

End Sub

Private Sub cmdOk_Click()
    If Len(IELogo_Picture) > 0 Then
        Form1.ttlogosm.Text = IELogo_Picture
        Unload FrmLogoChanger
    End If
    
End Sub

Private Sub cmdPrev_Click()
    ICount = 0
    Timer1.Enabled = True
    
End Sub

Private Sub Command1_Click()
    ICount = 0
    Timer1.Enabled = False
    
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    
End Sub

Private Sub Drive1_Change()
On Error Resume Next
    Dir1.Path = Drive1.Drive
    If Err Then MsgBox Err.Description, vbCritical
    
End Sub

Private Sub File1_Click()
    IELogo_Picture = AddSlash(File1.Path) & File1.Filename
    GetBitmapInfo
    If TBitmap.bmWidth <> MaxWidth Then
        Timer1.Enabled = False
        Preview.Picture = Nothing
        MsgBox "Inviald filename bitmap width must not be lower of higher than 22", vbCritical, "Error opening file"
        cmdPrev.Enabled = True
        cmdOk.Enabled = True
        IELogo_Picture = ""
        Exit Sub
    Else
        ICount = 0
        Srcbitmap.Picture = LoadPicture(IELogo_Picture)
        BitBlt Preview.hDC, 0, 0, MaxWidth, MaxWidth, Srcbitmap.hDC, 0, 0, SRCCOPY
        Preview.Refresh
        lblsize.Caption = TBitmap.bmWidth & " X " & TBitmap.bmHeight
        
    End If
    
End Sub

Private Sub Form_Load()
    Dir1.Path = App.Path & "\"
    
End Sub

Private Sub Timer1_Timer()
    ICount = ICount + 1
    If ICount = TBitmap.bmHeight / MaxWidth Then ICount = 0
    BitBlt Preview.hDC, 0, 0, MaxWidth, MaxWidth, Srcbitmap.hDC, 0, ICount * 22, SRCCOPY
    Preview.Refresh
    
    
End Sub
