VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM Internet Explorer 5 Tweak Kit Ver 1.1"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   8355
      TabIndex        =   14
      Top             =   0
      Width           =   8355
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DM Internet Explorer 5 Tweak Kit Ver 1.1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   75
         TabIndex        =   15
         Top             =   150
         Width           =   5730
      End
      Begin VB.Image Image4 
         Height          =   720
         Left            =   7500
         Picture         =   "Form1.frx":08CA
         Top             =   30
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tweak Options"
      Height          =   2325
      Left            =   45
      TabIndex        =   0
      Top             =   900
      Width           =   8250
      Begin VB.PictureBox pic5 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   1905
         Left            =   4410
         ScaleHeight     =   1905
         ScaleWidth      =   3735
         TabIndex        =   21
         Top             =   195
         Visible         =   0   'False
         Width           =   3735
         Begin VB.CommandButton Command4 
            Caption         =   "Disable Customize"
            Height          =   330
            Left            =   150
            TabIndex        =   23
            Top             =   570
            Width           =   1560
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Enable Customize"
            Height          =   330
            Left            =   150
            TabIndex        =   22
            Top             =   180
            Width           =   1560
         End
         Begin VB.Image Image6 
            Height          =   945
            Left            =   2130
            Picture         =   "Form1.frx":240C
            Top             =   825
            Width           =   1320
         End
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Disable IE Customize toolbar Option"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   240
         MouseIcon       =   "Form1.frx":3B4C
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   1695
         Width           =   2895
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Change your current download Folder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   240
         MouseIcon       =   "Form1.frx":4416
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   915
         Width           =   3360
      End
      Begin VB.PictureBox pic4 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   1830
         Left            =   4875
         ScaleHeight     =   1830
         ScaleWidth      =   3195
         TabIndex        =   16
         Top             =   5625
         Visible         =   0   'False
         Width           =   3195
         Begin VB.TextBox txtDownloadFolder 
            Height          =   330
            Left            =   135
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   315
            Width           =   2535
         End
         Begin VB.CommandButton Command2 
            Caption         =   "...."
            Height          =   330
            Left            =   2715
            TabIndex        =   17
            Top             =   330
            Width           =   375
         End
         Begin VB.Image Image5 
            Height          =   885
            Left            =   1380
            Picture         =   "Form1.frx":4CE0
            Top             =   870
            Width           =   1635
         End
      End
      Begin VB.CheckBox chkFullscreen 
         Caption         =   "Open IE in Full Screen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   240
         MouseIcon       =   "Form1.frx":6D78
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   1455
         Width           =   2895
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Change IE Title bar Caption"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   240
         MouseIcon       =   "Form1.frx":7642
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   1185
         Width           =   2895
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Change IE Toolbar Back Ground Picture"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   240
         MouseIcon       =   "Form1.frx":7F0C
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   630
         Width           =   3360
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Change IE Anmatied logo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   240
         MouseIcon       =   "Form1.frx":87D6
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   345
         Width           =   2895
      End
      Begin VB.PictureBox pic3 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   1560
         Left            =   4920
         ScaleHeight     =   1560
         ScaleWidth      =   3195
         TabIndex        =   7
         Top             =   4125
         Visible         =   0   'False
         Width           =   3195
         Begin VB.TextBox txtCaption 
            Height          =   330
            Left            =   30
            TabIndex        =   8
            Top             =   465
            Width           =   2955
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Enter your new Caption"
            Height          =   195
            Left            =   750
            TabIndex        =   9
            Top             =   165
            Width           =   1650
         End
         Begin VB.Image Image3 
            Height          =   390
            Left            =   45
            Picture         =   "Form1.frx":90A0
            Top             =   930
            Width           =   2910
         End
      End
      Begin VB.PictureBox pic2 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   1965
         Left            =   4920
         ScaleHeight     =   1965
         ScaleWidth      =   3195
         TabIndex        =   4
         Top             =   2400
         Visible         =   0   'False
         Width           =   3195
         Begin VB.CommandButton Command1 
            Caption         =   "...."
            Height          =   330
            Left            =   2565
            TabIndex        =   6
            Top             =   195
            Width           =   375
         End
         Begin VB.TextBox txtBackGround 
            Height          =   330
            Left            =   15
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   195
            Width           =   2535
         End
         Begin VB.Image Image2 
            Height          =   690
            Left            =   1380
            Picture         =   "Form1.frx":BB81
            Top             =   930
            Width           =   1485
         End
      End
      Begin VB.PictureBox pic1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   1965
         Left            =   4920
         ScaleHeight     =   1965
         ScaleWidth      =   3195
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   3195
         Begin VB.TextBox ttlogosm 
            Height          =   330
            Left            =   15
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   195
            Width           =   2535
         End
         Begin VB.CommandButton cmdLogo 
            Caption         =   "...."
            Height          =   330
            Left            =   2565
            TabIndex        =   2
            Top             =   195
            Width           =   375
         End
         Begin VB.Image Image1 
            Height          =   780
            Left            =   1890
            Picture         =   "Form1.frx":D3EA
            Top             =   1020
            Width           =   1020
         End
      End
   End
   Begin Project1.Button3D Button3D1 
      Height          =   510
      Left            =   5415
      TabIndex        =   25
      Top             =   3330
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   900
      BackColor       =   16777215
      Enabled         =   0   'False
      Picture         =   "Form1.frx":EF7E
   End
   Begin Project1.Button3D Button3D2 
      Height          =   510
      Left            =   6435
      TabIndex        =   26
      Top             =   3330
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   900
      BackColor       =   16777215
      Picture         =   "Form1.frx":F858
   End
   Begin Project1.Button3D Button3D3 
      Height          =   510
      Left            =   4410
      TabIndex        =   27
      Top             =   3330
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   900
      BackColor       =   16777215
      Picture         =   "Form1.frx":10132
   End
   Begin Project1.Button3D Button3D4 
      Height          =   510
      Left            =   7425
      TabIndex        =   28
      Top             =   3330
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   900
      BackColor       =   16777215
      Picture         =   "Form1.frx":10A0C
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Height          =   645
      Left            =   60
      TabIndex        =   24
      Top             =   3255
      Width           =   4785
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   0
      X2              =   1575
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   15
      X2              =   1590
      Y1              =   780
      Y2              =   780
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bar_Enable As Integer
Private Sub ShowDescription(Description As String)
    lblDescription.Caption = Description
    
End Sub
Private Sub Button3D1_Click()
    ans = MsgBox("Are you sure you want to save the new settings", _
    vbYesNo)
    If ans = vbNo Then
        Exit Sub
    Else
        
        SaveString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Toolbar\", "BackBitmap", IEConfig.IEBackBitmap
        SaveString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Toolbar\", "BrandBitmap", IEConfig.IElogo
        SaveString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Toolbar\", "SmBrandBitmap", IEConfig.IElogo
        SaveString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\", "Download Directory", IEConfig.IEDownloadFolder
        SaveString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main\", "FullScreen", IEConfig.IEOpenFullSceen
        SaveString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main\", "Window Title", IEConfig.IETitleBarCaption
        SaveDword HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoBandCustomize", IEConfig.IEToolBarCustomize
        MsgBox "You may have to restart your computer before all the new settings take effect", vbInformation
        
    End If
    

End Sub

Private Sub Button3D2_Click()
    Button3D1.Enabled = True
    IEConfig.IEBackBitmap = ""
    IEConfig.IElogo = ""
    IEConfig.IETitleBarCaption = "Internet Explorer"
    IEConfig.IEDownloadFolder = txtDownloadFolder
    IEConfig.IEOpenFullSceen = "no"
    IEConfig.IEToolBarCustomize = 0
    
    SaveString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Toolbar\", "BackBitmap", IEConfig.IEBackBitmap
    SaveString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Toolbar\", "BrandBitmap", IEConfig.IElogo
    SaveString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Toolbar\", "SmBrandBitmap", IEConfig.IElogo
    SaveString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\", "Download Directory", IEConfig.IEDownloadFolder
    SaveString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main\", "FullScreen", IEConfig.IEOpenFullSceen
    SaveString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main\", "Window Title", IEConfig.IETitleBarCaption
    MsgBox "Your old settings have been restored", vbInformation
    
End Sub

Private Sub Button3D3_Click()
    IEConfig.IEBackBitmap = txtBackGround
    IEConfig.IElogo = ttlogosm
    IEConfig.IETitleBarCaption = txtCaption
    IEConfig.IEDownloadFolder = txtDownloadFolder
    IEConfig.IEOpenFullSceen = chkFullscreen.Tag
    IEConfig.IEToolBarCustomize = Bar_Enable
    Button3D1.Enabled = True
    
    
End Sub

Private Sub Button3D4_Click()
    Unload Form1: End
    
End Sub

Private Sub Check1_Click()
    Check2.Value = 0
    Check3.Value = 0
    Check4.Value = 0
    Check5.Value = 0
    pic1.Visible = True
    pic2.Visible = False
    pic3.Visible = False
    pic4.Visible = False
    pic5.Visible = False
    pic1.Top = 240
    ShowDescription "Setting this option will let you change the small Animated " & vbCrLf & "logo that spins around in top right the corner"

    
End Sub

Private Sub Check2_Click()
    Check1.Value = 0
    Check3.Value = 0
    Check4.Value = 0
    Check5.Value = 0
    pic2.Visible = True
    pic1.Visible = False
    pic3.Visible = False
    pic4.Visible = False
    pic5.Visible = False
    pic2.Top = 240
    ShowDescription "This option will allow you to change add a picture" & vbCrLf & " as a background on the Internet Expolrer tool bar."
    

End Sub

Private Sub Check3_Click()
    Check1.Value = 0
    Check2.Value = 0
    Check4.Value = 0
    Check5.Value = 0
    pic3.Visible = True
    pic1.Visible = False
    pic2.Visible = False
    pic4.Visible = False
    pic5.Visible = False
    pic3.Top = 240
    ShowDescription "Setting this option will let you change the caption" & vbCrLf & "On IE title bar to what ever you like e.g." & vbCrLf & "maybe the name or your company."
    

End Sub

Private Sub Check4_Click()
    Check2.Value = 0
    Check3.Value = 0
    Check1.Value = 0
    Check5.Value = 0
    pic4.Visible = True
    pic2.Visible = False
    pic3.Visible = False
    pic1.Visible = False
    pic5.Visible = False
    pic4.Top = 240
    ShowDescription "This option will let you choose a new folder were your" & vbCrLf & "downloaded files are to be placed."

End Sub

Private Sub Check5_Click()
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    Check4.Value = 0
    pic1.Visible = False
    pic2.Visible = False
    pic3.Visible = False
    pic4.Visible = False
    pic5.Visible = True
    ShowDescription "Setting this option will allow you to enable or disable" & vbCrLf & "the customize option in IE very good if you don't want" & vbCrLf & " anyone changeing your toolbar around."
    
    
    
End Sub

Private Sub Check6_Click()

End Sub

Private Sub chkFullscreen_Click()
    If chkFullscreen Then
        chkFullscreen.Tag = "yes"
    Else
        chkFullscreen.Tag = "no"
    End If
    ShowDescription "You can use this option to let you start Inter Explorer in " & vbCrLf & "full screen mode."
    
    
End Sub

Private Sub cmdLogo_Click()
    FrmLogoChanger.Show
    
End Sub

Private Sub Command1_Click()
Dim Filename, FExt As String
    Filename = OpenFile
    FExt = UCase(Right(Filename, 3))
    If FExt <> "BMP" Then
        MsgBox "Invaild fileformat please use bitmap files names", vbInformation
        Exit Sub
    Else
        txtBackGround.Text = Filename
        Filename = ""
        FExt = ""
    End If
    
End Sub

Private Sub Command2_Click()
Dim FolName As String
    FolName = AddSlash(GetFolder(hWnd, "Please Select a new Dowload Folder"))
    If FolName = "\" Then
        txtDownloadFolder.Text = GetValues(HKEY_CURRENT_USER, REG_SZ, "Software\Microsoft\Internet Explorer\", "Download Directory")
    Else
        txtDownloadFolder.Text = FolName
        FolName = ""
    End If
    
End Sub

Private Sub Command3_Click()
    Bar_Enable = 0
    
End Sub

Private Sub Command4_Click()
    Bar_Enable = 1
    
End Sub

Private Sub Form_Load()
Dim FullScreen As String
Dim Toolbar As String
On Error Resume Next
    txtCaption.Text = GetValues(HKEY_CURRENT_USER, REG_SZ, "Software\Microsoft\Internet Explorer\Main\", "Window Title")
    txtDownloadFolder.Text = GetValues(HKEY_CURRENT_USER, REG_SZ, "Software\Microsoft\Internet Explorer\", "Download Directory")
    ttlogosm.Text = GetValues(HKEY_CURRENT_USER, REG_SZ, "Software\Microsoft\Internet Explorer\ToolBar\", "SmBrandBitmap")
    txtBackGround.Text = GetValues(HKEY_CURRENT_USER, REG_SZ, "Software\Microsoft\Internet Explorer\ToolBar\", "BackBitmap")
    FullScreen = UCase(GetValues(HKEY_CURRENT_USER, REG_SZ, "Software\Microsoft\Internet Explorer\Main\", "FullScreen"))
    Toolbar = UCase(GetValues(HKEY_CURRENT_USER, REG_DWORD, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoBandCustomize"))
    
    If Asc(Toolbar) = 0 Then
        Check5.Value = 0
    Else
        Check5.Value = 1
    End If
    
    If FullScreen = "NO" Then
        chkFullscreen.Value = 0
        chkFullscreen.Tag = "no"
    Else
        chkFullscreen.Value = 1
        chkFullscreen.Tag = "yes"
    End If
    
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Form_Resize()
    Line1(0).X2 = Form1.ScaleWidth
    Line1(1).X2 = Form1.ScaleWidth
    
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub toolbarebanle_Click()
 
End Sub

Private Sub optDisable_Click()
    Bar_Enable = 1
    
End Sub

Private Sub optEnable_Click()
    Bar_Enable = 0
    
End Sub

Private Sub Image4_Click()
    frmabout.Show
    
End Sub
