VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_ARENA_TRASHED = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259
Private Const ERROR_SUCCESS& = 0&

Private Const KEY_ALL_ACCESS = &H3F

Private Const REG_OPTION_NON_VOLATILE = 0

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long


Private mvar_hKey As Long
Private mvar_bInit As Boolean

Public Function CreateKey( _
   KeyName As String, _
   Optional hKey As Long = HKEY_LOCAL_MACHINE _
) As Boolean

   Dim lret As Long

   lret = RegCreateKeyEx(hKey, KeyName, 0&, _
         vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
         0&, mvar_hKey, lret)
   If lret = ERROR_SUCCESS Then
      CreateKey = True
      mvar_bInit = True
   End If
End Function

Public Sub SetLong( _
   Value As String, _
   nData As Long _
)
   Dim lValue As Long, lret As Long
   CheckInit

   lValue = nData
   lret = RegSetValueExLong(mvar_hKey, Value, 0&, REG_DWORD, lValue, 4)
End Sub

Public Sub SetStringValue( _
   Value As String, _
   sData As String _
)
   Dim sValue As String, lret As Long
   CheckInit

   sValue = sData & Chr$(0)
   lret = RegSetValueExString(mvar_hKey, Value, 0&, REG_SZ, sValue, Len(sValue))
End Sub

Public Function GetLong( _
   Value As String, _
   Optional nDefault As Long = 0 _
) As Long

   Dim cch As Long
   Dim lrc As Long
   Dim lValue As Long
   Dim sValue As String
   Dim lType As Long

   CheckInit
   On Error GoTo GetLong_Error

   GetLong = nDefault
   ' Determine the size and type of data to be read
   lrc = RegQueryValueExNULL(mvar_hKey, Value, 0&, lType, 0&, cch)
   If lrc <> ERROR_NONE Then GoTo GetLong_Error

   If lType = REG_DWORD Then
      lrc = RegQueryValueExLong(mvar_hKey, Value, 0&, REG_DWORD, lValue, cch)
      If lrc = ERROR_NONE Then GetLong = lValue
   End If

GetLong_Exit:
   Exit Function

GetLong_Error:
   On Error GoTo 0
   GoTo GetLong_Exit

End Function

Public Function GetStringValue( _
   Value As String, _
   Optional sDefault As String = "" _
) As String

   Dim cch As Long
   Dim lrc As Long
   Dim lType As Long
   Dim lValue As Long
   Dim sValue As String

   CheckInit
   On Error GoTo GetStringValue_Error

   GetStringValue = sDefault

   ' Determine the size and type of data to be read
   lrc = RegQueryValueExNULL(mvar_hKey, Value, 0&, lType, 0&, cch)
   If lrc <> ERROR_NONE Then GoTo GetStringValue_Error

   If lType = REG_SZ Then
      ' For strings
      sValue = String(cch, 0)
      lrc = RegQueryValueExString(mvar_hKey, Value, 0&, REG_SZ, sValue, cch)
      If lrc = ERROR_NONE Then
         GetStringValue = Left$(sValue, cch - 1)
      End If
   End If


GetStringValue_Exit:
   Exit Function

GetStringValue_Error:
   On Error GoTo 0
   GoTo GetStringValue_Exit

End Function

Public Function OpenKey( _
   KeyName As String, _
   Optional hKey As Long = HKEY_LOCAL_MACHINE _
) As Boolean

   Dim lRetVal As Long

   lRetVal = RegOpenKeyEx(hKey, KeyName, 0, KEY_ALL_ACCESS, mvar_hKey)
   If lRetVal = ERROR_NONE Then
      OpenKey = True
      mvar_bInit = True
   End If
End Function

Private Sub Class_Terminate()
   If mvar_hKey <> 0 Then
      RegCloseKey mvar_hKey
   End If
End Sub

Private Sub CheckInit()
   If mvar_bInit <> True Then
      Err.Raise 1777, "Class cRegistry not Initialized."
   End If
End Sub



