VERSION 5.00
Begin VB.Form frmservices 
   BackColor       =   &H80000004&
   Caption         =   "Service Creator"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdRemove 
      Caption         =   "Remove Service :"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txt_path_name 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txt_serv_name 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdcreate 
      Caption         =   "Install Service :"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblsername 
      Caption         =   "Enter service name :"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblpath 
      Caption         =   "Enter the path :"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmservices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 19-MAR-2002  Himanshu Dhami              Written by himansh_dhami@yahoo.com
' References  ::The idea for instsrv is borrowed from Microsoft KB article Q137890
               'Help in API for regsitry handling is borrowed from VBAPI.com
               'Need to put srvany in c:\srvany download from
               'http://bscw.gmd.de/Download/srvany.zip or from microsoft.
               'Make project reference to Windows script host object model
' ---------------------------------------------------------------------------


Public srv_name As String
Public path_name As String
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_WRITE = &H20006
Private Const REG_SZ = 1
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal _
    hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass _
    As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes _
    As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal _
    hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType _
    As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' *** Place the following code inside the form. ***
Private Sub ServiceCreator()
    Dim hKey As Long            ' receives handle to the registry key
    Dim secattr As SECURITY_ATTRIBUTES  ' security settings for the key
    Dim subkey As String        ' name of the subkey to create or open
    Dim neworused As Long       ' receives flag for if the key was created or opened
    Dim stringbuffer As String  ' the string to put into the registry
    Dim retval As Long          ' return value
    Dim test As String
    
    ' Set the name of the new key and the default security settings
    subkey = "System\CurrentControlSet\Services\" & srv_name & "\Parameters"
    secattr.nLength = Len(secattr)
    secattr.lpSecurityDescriptor = 0
    secattr.bInheritHandle = 1
    
    ' Create (or open) the registry key.
    retval = RegCreateKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, "", 0, KEY_WRITE, _
        secattr, hKey, neworused)
    If retval <> 0 Then
        Debug.Print "Error opening or creating registry key -- aborting."
        Exit Sub
        
    End If
    ' Write the string to the registry.  Note the use of ByVal in the second-to-last
    ' parameter because we are passing a string.
    stringbuffer = path_name & vbNullChar    ' the terminating null is necessary
    retval = RegSetValueEx(hKey, "Application", 0, REG_SZ, ByVal stringbuffer, _
             Len(stringbuffer))
    ' Close the registry key.
    retval = RegCloseKey(hKey)
End Sub
Private Sub cmdcreate_Click()
srv_name = txt_serv_name.Text
path_name = txt_path_name.Text
MsgBox "the sevice name is" & srv_name & " and pathname is " & path_name & " ", vbOKOnly
'End
Dim te As WshShell
Set te = New WshShell
te.Run ("c:\srvany\instsrv " & srv_name & " c:\srvany\srvany.exe")
Sleep 2000
ServiceCreator
MsgBox ("Successfully created string in the registry"), vbOKOnly
End
End Sub
Private Sub CmdRemove_Click()
srv_name = txt_serv_name.Text
path_name = txt_path_name.Text
MsgBox "the sevice name is" & srv_name & " and pathname is " & path_name & " ", vbOKOnly
Dim te As WshShell
Set te = New WshShell
te.Run ("c:\srvany\instsrv " & srv_name & " remove ")
Sleep 1000
MsgBox ("Successfully removed from the service"), vbOKOnly
End
End Sub


