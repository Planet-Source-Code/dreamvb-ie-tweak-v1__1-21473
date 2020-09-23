Attribute VB_Name = "Module1"
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public IELogo_Picture As String
Public Const SRCCOPY = &HCC0020
Public Const REG_SZ = 1
Public Const REG_DWORD = 4
Public Const HKEY_CURRENT_USER = &H80000001

Private Const ERROR_SUCCESS = 0&
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_Value = &H1
Private Const KEY_SET_Value = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000


Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Const KEY_ALL_ACCESS = _
    ((STANDARD_RIGHTS_ALL Or _
    KEY_QUERY_Value Or _
    KEY_SET_Value Or _
    KEY_CREATE_SUB_KEY Or _
    KEY_ENUMERATE_SUB_KEYS Or _
    KEY_NOTIFY Or KEY_CREATE_LINK) And _
    (Not SYNCHRONIZE))

Private Type IE_Config
    IElogo As String
    IEBackBitmap As String
    IETitleBarCaption As String
    IEDownloadFolder As String
    IEOpenFullSceen As String
    IEToolBarCustomize As String
End Type

Public IEConfig As IE_Config
Function AddSlash(lzPath As String) As String
    If Right(lzPath, 1) = "\" Then AddSlash = lzPath Else AddSlash = lzPath & "\"
    
End Function

Function GetFolder(ByVal hWndOwner As Long, ByVal sTitle As String) As String
Dim bInf As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim RetPath As String
Dim Offset As Integer
    bInf.hOwner = hWndOwner
    bInf.lpszTitle = sTitle
    bInf.ulFlags = BIF_RETURNONLYFSDIRS
    PathID = SHBrowseForFolder(bInf)
    RetPath = Space$(512)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
    If RetVal Then
      Offset = InStr(RetPath, Chr$(0))
      GetFolder = Left$(RetPath, Offset - 1)
    End If
End Function

Function GetValues(hKey As Long, Value_type As Long, lzPath As String, strValue As String) As String
Dim Value As String
Dim StrLen As Long
On Error Resume Next
    ' The Key you want to open
    If RegOpenKeyEx(hKey, lzPath, 0&, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then
        Exit Function
    End If
    
    'Get the subkey's Value
    StrLen = 256
    Value = Space(StrLen)
    
    If RegQueryValueEx(hKey, strValue, 0&, Value_type, ByVal Value, StrLen) <> ERROR_SUCCESS Then
       Exit Function
    Else
        ' Remove all trailing null character
        Value = Left(Value, StrLen - 1)
        GetValues = Value
    End If
    
    ' Close the key.
    If RegCloseKey(hKey) <> ERROR_SUCCESS Then
        MsgBox "Error closing key." ' Show error if it happens
    End If
    
End Function
Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    'If lResult <> error_success Then Call errlog("SetDWORD", False)
    r = RegCloseKey(keyhand)
End Function
Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
Dim keyhand As Long
Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)

End Sub
Private Function RemoveNulls(lzString As String) As String
Dim XPos As Integer
    XPos = InStr(lzString, vbNullChar)
    If XPos > 0 Then
        lzString = Left(lzString, Len(lzString) - 1)
        RemoveNulls = lzString
    End If
    
End Function
Public Function OpenFile() As String
 Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hWndOwner = Form1.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "All Files(*.bmp Windows Pictures)" + Chr$(0) + "*.bmp"
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = App.Path & "\"
        ofn.lpstrTitle = "Open Bitmap"
        ofn.flags = 0
        
        a = GetOpenFileName(ofn)
        If (a) Then
                OpenFile = RemoveNulls(Trim(ofn.lpstrFile))
        End If
        
 End Function
