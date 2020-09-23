Attribute VB_Name = "modRegister"
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, _
                ByVal lpValueName As String, ByRef lpcbValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, _
                ByRef lpData As Byte, ByRef lpcbData As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
                ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
                ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
              ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Public Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
              ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
              ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
              ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
              ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public sCateporyToDelete As String
Public Const KEY_QUERY_VALUE = &H1&
Public Const KEY_ENUMERATE_SUB_KEYS = &H8&
Public Const KEY_NOTIFY = &H10&
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = READ_CONTROL
Global Const KEY_ALL_ACCESS = &H3F
Public Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Public Const FilelistKey = "Software\VB and VBA Program Settings\The Weather Program\BookMark"
Public Const CityCodeValue = "Software\VB and VBA Program Settings\The Weather Program\City Information"
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const KEY_WRITE = &H20006
Public Const REG_DWORD = 4
Public Const REG_SZ = 1
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_MORE_DATA = 234&
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Public Security As SECURITY_ATTRIBUTES

Public Function DeleteRegisterValue(lPredefinedKey As Long, sKeyName As String, sValueName As String) As Long
  Dim lRetVal As Long         'result of the API functions
  Dim hKey As Long         'handle of opened key
  Dim vValue As Variant      'setting of queried value
  lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
  lRetVal = RegDeleteValue(hKey, sValueName)
  RegCloseKey (hKey)
End Function

Function EnumRegistryValues(ByVal hKey As String, ByVal KeyName As String) As Collection
  Dim handle As Long
  Dim Index As Long
  Dim valueType As Long
  Dim Name As String
  Dim nameLen As Long
  Dim resLong As Long
  Dim resString As String
  Dim Length As Long
  Dim valueInfo(0 To 1) As Variant
  Dim retval As Long
  Dim i As Integer
  Dim vTemp As Variant
  
  ' initialize the result
  Set EnumRegistryValues = New Collection
    
  ' Open the key, exit if not found.
  If Len(KeyName) Then
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
    ' in all cases, subsequent functions use hKey
    hKey = handle
  End If
    
  Do
    ' this is the max length for a key name
    nameLen = 260
    Name = Space$(nameLen)
    ' prepare the receiving buffer for the value
    Length = 4096
    ReDim resBinary(0 To Length - 1) As Byte
    
    ' read the value's name and data
    ' exit the loop if not found
    retval = RegEnumValue(hKey, Index, Name, nameLen, ByVal 0&, valueType, resBinary(0), Length)
    
    ' enlarge the buffer if you need more space
    If retval = ERROR_MORE_DATA Then
      ReDim resBinary(0 To Length - 1) As Byte
      retval = RegEnumValue(hKey, Index, Name, nameLen, ByVal 0&, valueType, resBinary(0), Length)
    End If
    ' exit the loop if any other error (typically, no more values)
    If retval Then Exit Do
        
    ' retrieve the value's name
    valueInfo(0) = Left$(Name, nameLen)
    
    ' copy everything but the trailing null char
    If Length <> 0 Then
      resString = Space$(Length - 1)
      CopyMemory ByVal resString, resBinary(0), Length - 1
      valueInfo(1) = resString
    Else
      valueInfo(1) = ""
    End If
    ' add the array to the result collection
    ' the element's key is the value's name
    EnumRegistryValues.Add valueInfo, valueInfo(0)
    Index = Index + 1
  Loop
  ' Close the key, if it was actually opened
  If handle Then RegCloseKey handle
End Function

Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
' Description:
'   This Function will return the data field of a value
'
' Syntax:
'   Variable = QueryValue(Location, KeyName, ValueName)
'
'   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   , HKEY_USERS
'
'   KeyName is the key that the value is under (example: "Software\Microsoft\Windows\CurrentVersion\Explorer")
'HKEY_CURRENT_USER\Software\VB and VBA Program Settings\The Weather Program\City Information
'   ValueName is the name of the value you want to access (example: "link")

  Dim lRetVal As Long         'result of the API functions
  Dim hKey As Long         'handle of opened key
  Dim vValue As Variant      'setting of queried value
  lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
  lRetVal = QueryValueEx(hKey, sValueName, vValue)
  QueryValue = vValue
  RegCloseKey (hKey)
End Function

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
  Dim cch As Long
  Dim lrc As Long
  Dim lType As Long
  Dim lValue As Long
  Dim sValue As String
  On Error GoTo QueryValueExError
  ' Determine the size and type of data to be read
  lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
  If lrc <> ERROR_NONE Then
    Error 5
  End If
  Select Case lType
    'For strings
    Case REG_SZ:
      sValue = String(cch, 0)
      lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
      If lrc = ERROR_NONE Then
        vValue = Left$(sValue, cch)
      Else
        vValue = Empty
      End If
      'For multi strings
    Case REG_MULTI_SZ:
      sValue = String(cch, 0)
      lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
      If lrc = ERROR_NONE Then
        vValue = Left$(sValue, cch)
      Else
        vValue = Empty
      End If
      ' For DWORDS
    Case REG_DWORD:
      lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
      If lrc = ERROR_NONE Then
        vValue = lValue
      End If
    Case Else
      'all other data types not supported
      lrc = -1
  End Select
QueryValueExExit:
  QueryValueEx = lrc
  Exit Function
QueryValueExError:
  Resume QueryValueExExit
End Function
