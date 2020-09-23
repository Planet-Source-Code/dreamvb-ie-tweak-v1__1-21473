VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Abour IE Logo Changer"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3075
      TabIndex        =   2
      Top             =   2040
      Width           =   720
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   3990
      TabIndex        =   0
      Top             =   0
      Width           =   3990
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   105
         Picture         =   "frmabout.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         Top             =   30
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DM Internet Explorer 5 Tweak Kit Ver 1.1"
         Height          =   195
         Left            =   870
         TabIndex        =   6
         Top             =   165
         Width           =   2910
      End
   End
   Begin VB.Image Image1 
      Height          =   75
      Left            =   -855
      Picture         =   "frmabout.frx":0C42
      Top             =   645
      Width           =   6195
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Website http://www.dreamvb.s5.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1080
      TabIndex        =   5
      Top             =   1545
      Width           =   2670
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Email Dreamvb@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1095
      TabIndex        =   4
      Top             =   1245
      Width           =   2010
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Writen by Ben Jones"
      Height          =   195
      Left            =   1200
      TabIndex        =   3
      Top             =   930
      Width           =   1470
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload frmabout
    
End Sub

