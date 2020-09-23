VERSION 5.00
Begin VB.UserControl Button3D 
   BackColor       =   &H00FF8080&
   ClientHeight    =   4050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4725
   ScaleHeight     =   4050
   ScaleWidth      =   4725
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00000000&
      FillColor       =   &H00C0C0C0&
      Height          =   15
      Left            =   240
      ScaleHeight     =   15
      ScaleWidth      =   4455
      TabIndex        =   3
      Top             =   3690
      Width           =   4455
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      FillColor       =   &H00C0C0C0&
      Height          =   3165
      Left            =   4560
      ScaleHeight     =   3165
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   675
      Width           =   15
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   2925
      Left            =   390
      ScaleHeight     =   2925
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   840
      Width           =   15
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   15
      Left            =   255
      ScaleHeight     =   15
      ScaleWidth      =   4470
      TabIndex        =   0
      Top             =   795
      Width           =   4470
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   600
      Top             =   1920
      Width           =   480
   End
End
Attribute VB_Name = "Button3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event Hide() 'MappingInfo=UserControl,UserControl,-1,Hide
Attribute Hide.VB_Description = "Occurs when the control's Visible property changes to False."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Private Sub TMouseUp()
    Picture1.BackColor = vbWhite
    Picture2.BackColor = vbWhite
    Picture3.BackColor = vbBlack
    Picture4.BackColor = vbBlack
End Sub
Private Sub TMouseDown()
Picture1.BackColor = vbBlack
Picture2.BackColor = vbBlack
Picture3.BackColor = vbWhite
Picture4.BackColor = vbWhite
    
End Sub

Private Sub Image1_Click()
RaiseEvent Click

End Sub

Private Sub Image1_DblClick()
RaiseEvent DblClick

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = 1 Then
        TMouseDown
    End If
    
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
  
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseUp(Button, Shift, X, Y)
If Button = 1 Then
    TMouseUp
End If

End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
If Button = 1 Then
    TMouseDown
End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
If Button = 1 Then
    TMouseUp
End If

End Sub

Private Sub UserControl_Resize()
Picture1.Move 0, 0, 20, UserControl.Height
Picture2.Move 0, 0, UserControl.Width, 20
Picture3.Move UserControl.ScaleWidth - 20, 0, 20, UserControl.Height
Picture4.Move 0, UserControl.ScaleHeight - 20, UserControl.Width, 20
'
Image1.Left = (UserControl.Width - Image1.Width) / 2
Image1.Top = (UserControl.Height - Image1.Height) / 2

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1,Image1,-1,Caption

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub UserControl_Hide()
    RaiseEvent Hide
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFF8080)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFF8080)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1,Image1,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = Image1.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Image1.Picture = New_Picture
    PropertyChanged "Picture"
End Property


