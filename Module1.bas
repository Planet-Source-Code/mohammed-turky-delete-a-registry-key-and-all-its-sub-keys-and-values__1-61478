Attribute VB_Name = "Module1"
'Option Explicit
'Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
'Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
'
''Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
'Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
'
'Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
'Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
'Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
'
'Public Const ERROR_SUCCESS = 0&
'Public Const ERROR_BADKEY = 1010&
'Public Const ACCESS_DENIED = 5
'Public Const ERROR_NO_MORE_ITEMS = 259&
'Private Const READ_CONTROL = &H20000
'Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
'Private Const STANDARD_RIGHTS_READ = READ_CONTROL
'Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL
'Private Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
'Private Const STANDARD_RIGHTS_ALL = &H1F0000
'Private Const KEY_NOTIFY = &H10&
'Private Const KEY_ENUMERATE_SUB_KEYS = &H8 'Permission to enumerate subkeys.
'Private Const KEY_QUERY_VALUE = &H1&
'Private Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
'
'Public Const KEY_ALL_ACCESS = &H3F
'
'Private Const MAX_KEY_SIZE = 255
'Private Const MAX_VALUE_SIZE = 4096
'
'Public Type KeyInfo       ' Type to store information about a key
'    Subkeys As Long       ' number of subkeys
'    LenSubkeys As Long    ' length of the longest subkey names
'    Values As Long        ' number of value entries in the key
'    LenValueNames As Long ' length of the longest value names
'    LenValues As Long     ' length (bytes) of the longest value
'End Type
'
'Private Type FILETIME
'  dwLowDateTime As Long
'  dwHighDateTime As Long
'End Type
'
'Public Enum ROOT_KEYS
'     HKEY_ALL = &H0&
'     HKEY_CLASSES_ROOT = &H80000000
'     HKEY_CURRENT_USER = &H80000001
'     HKEY_LOCAL_MACHINE = &H80000002
'     HKEY_USERS = &H80000003
'     HKEY_PERFORMANCE_DATA = &H80000004
'     HKEY_CURRENT_CONFIG = &H80000005
'     HKEY_DYN_DATA = &H80000006
'End Enum
'Public Enum KeyType
'    REG_BINARY = 3 'A non-text sequence of bytes
'    REG_DWORD = 4  'A 32-bit integer...visual basic data type of Long
'    REG_SZ = 1     'A string terminated by a null character
'    REG_MULTI_SZ = 7
'    REG_EXPAND_SZ = 2
'    REG_NONE = 0
'    REG_DWORD_LITTLE_ENDIAN = 4
'    REG_DWORD_BIG_ENDIAN = 5
'    REG_LINK = 6
'    REG_RESOURCE_LIST = 8
'    REG_FULL_RESOURCE_DESCRIPTOR = 9
'    REG_RESOURCE_REQUIREMENTS_LIST = 10
'End Enum
'Public Type REG_VALUE
'    sValues(2) As String   '0=sKyPath,1=sKeyName,2=sKeyValue
'    iSerachType As Integer 'for automatically selecting the type of serach from the above array without if statments in scanning procedure...
'End Type
'Public Type ENUM_TYPES
'    sEnumAllKeys() As String       'the paths that will be enum by all keys recursivly('search by all keys under this key and subkeys).
'    sEnumKeys() As String          'the paths that will be enum to get all its child keys ('search by all keys under this key only).
'    sEnumAllKeysValues() As String 'the paths that will ne enum by All keys recursivly until we reach the values ('Search by all keys and thier values under this key)
'    sEnumAllValues() As String     'the paths that will be enum by values only('Serach by all values under this key and all subkeys)
'    sEnumValues() As String        'the paths that will be enum by values only('Serach by all values under this key only)
'End Type
'Public Enum SEARCH_VALUE_TYPE 'to serach if a string is exist in any value or value name under a key
'    Value_Value = 1
'    Value_Name = 2
'    Value_Both = 3
'End Enum
'Private hKey      As Long
'Private sKey      As String
'Private mainKey   As Long
''Private Enum rcRegType
''    REG_NONE = 0
''    REG_SZ = 1
''    REG_EXPAND_SZ = 2
''    REG_BINARY = 3
''    REG_DWORD = 4
''    REG_DWORD_LITTLE_ENDIAN = 4
''    REG_DWORD_BIG_ENDIAN = 5
''    REG_LINK = 6
''    REG_MULTI_SZ = 7
''    REG_RESOURCE_LIST = 8
''    REG_FULL_RESOURCE_DESCRIPTOR = 9
''    REG_RESOURCE_REQUIREMENTS_LIST = 10
''End Enum
'
'Private Function GetRegData(ByVal lType As Long, abData() As Byte) As String
'   Dim lData As Long, i As Long
'   Dim sTemp As String
'   sTemp = ""
'   Select Case lType
'        Case REG_SZ, REG_MULTI_SZ, REG_EXPAND_SZ
'             GetRegData = StripNulls(StrConv(abData, vbUnicode))
'        Case REG_DWORD
'             CopyMem lData, abData(0), 4&
'             GetRegData = "0x" & Format(Hex(lData), "00000000") ' & "(" & lData & ")"
'        Case REG_BINARY
'             For i = 0 To UBound(abData)
'                 sTemp = sTemp & Right("00" & Hex(abData(i)), 2) & " "
'             Next i
'             GetRegData = Left(sTemp, Len(sTemp) - 1)
'        Case Else
'             GetRegData = "Temporary unsupported"
'   End Select
'End Function
''Function to opan a Reg key given its path(like HKEY_USERS\???\???), and returns a handle to it.
'Public Function RegGetKeyHandle(sKeyPath As String, OpenMode As Long, ByRef hRetKey As Long) As Long
'Dim i As Integer, index As Long
'Dim lKeyRoot As ROOT_KEYS
'Dim sSubkey As String
'Dim lType As Long, valueName As String, ValueValue As String, curidx As Long
'Dim arrData() As Byte, cbDataSize As Long
'
'On Error GoTo Loc_Error
'
'sKeyPath = Replace(sKeyPath, "HKCU", "HKEY_CURRENT_USER", , , vbTextCompare)
'sKeyPath = Replace(sKeyPath, "HKLM", "HKEY_LOCAL_MACHINE", , , vbTextCompare)
'sKeyPath = Replace(sKeyPath, "HKCR", "HKEY_CLASSES_ROOT", , , vbTextCompare)
'sKeyPath = Replace(sKeyPath, "HKUS", "HKEY_USERS", , , vbTextCompare)
'sKeyPath = Replace(sKeyPath, "HKPD", "HKEY_PERFORMANCE_DATA", , , vbTextCompare)
'sKeyPath = Replace(sKeyPath, "HKDD", "HKEY_DYN_DATA", , , vbTextCompare)
'sKeyPath = Replace(sKeyPath, "HKCC", "HKEY_CURRENT_CONFIG", , , vbTextCompare)
'
'ReDim rRet(0)
'RegGetKeyHandle = ERROR_BADKEY
'    If InStr(sKeyPath, "HKEY_CURRENT_USER") Then
'        sSubkey = Mid(sKeyPath, 19): lKeyRoot = HKEY_CURRENT_USER
'    ElseIf InStr(sKeyPath, "HKEY_CLASSES_ROOT") Then
'        sSubkey = Mid(sKeyPath, 19): lKeyRoot = HKEY_CLASSES_ROOT
'    ElseIf InStr(sKeyPath, "HKEY_LOCAL_MACHINE") Then
'        sSubkey = Mid(sKeyPath, 20): lKeyRoot = HKEY_LOCAL_MACHINE
'    ElseIf InStr(sKeyPath, "HKEY_USERS") Then
'        sSubkey = Mid(sKeyPath, 12): lKeyRoot = HKEY_USERS
'    End If
'   RegGetKeyHandle = RegOpenKeyEx(lKeyRoot, sSubkey, 0, OpenMode, hRetKey)
'Exit Function
'Loc_Error:
'LogError.LoggError Err, "RegGetKeyHandle", "Unable to open a registry key"
'End Function
'Public Function GetKeyNameValue(keyPath As String, keyName As String, Optional ByRef lType As Long) As String
'  ' routine to get the registry key value and convert to a string
'  On Error Resume Next
'  Dim ReturnValue As Long
'  Dim KeyType As KeyType
'  Dim MyBuffer As String
'  Dim MyBufferSize As Long
'  Dim hKey As Long
'On Error GoTo Loc_Error
'
'  'get registry key information
'  ReturnValue = RegGetKeyHandle(keyPath, KEY_ALL_ACCESS, hKey) 'get the key handle
'  ReturnValue = RegQueryValueEx(hKey, keyName, 0, KeyType, ByVal 0, MyBufferSize)
'  If ReturnValue = 0 Then ' no error encountered
'    ' determine what the KeyType is
'    Select Case KeyType
'      Case REG_SZ
'        ' create a buffer
'        MyBuffer = String(MyBufferSize, Chr$(0))
'        ' retrieve the key's content
'        ReturnValue = RegQueryValueEx(hKey, keyName, 0, 0, ByVal MyBuffer, MyBufferSize)
'        If ReturnValue = 0 Then
'          ' remove the unnecessary chr$(0)'s
'          GetKeyNameValue = Left$(MyBuffer, InStr(1, MyBuffer, Chr$(0)) - 1)
'        End If
'      Case Else 'REG_DWORD or REG_BINARY
'        Dim MyNewBuffer As Long
'        ' retrieve the key's value
'        ReturnValue = RegQueryValueEx(hKey, keyName, 0, 0, MyNewBuffer, MyBufferSize)
'        If ReturnValue = 0 Then ' no error encountered
'          GetKeyNameValue = MyNewBuffer
'        End If
'    End Select
'  End If
'  lType = KeyType
'Exit Function
'Loc_Error:
'
'LogError.LoggError Err, GetKeyNameValue, "Unable to get a key value"
'End Function
'
''Function to check a Reg key OR reg value existence.Pass KeyValueName if you need to check a key value
'Public Function IsRegKey_ValueExist(sKey As String, Optional KeyValueName As String) As Long
'    Dim lRetKey As Long      'Result for opening a key
'    Dim lRetVal As Long      'Result for opening a key value
'    Dim hKey As Long         'handle of opened key
'    Dim vValue As Variant    'setting of queried value
'    Dim KeyType As Long
'    Dim SubKeyName As String
'    Dim MyBufferSize As Long
'    On Error GoTo localError
'
'     If Right(sKey, 1) = "\" Then sKey = Mid(sKey, 1, Len(sKey) - 1)
'
'   lRetKey = RegGetKeyHandle(sKey, KEY_ALL_ACCESS, hKey)
'
'    If Not IsMissing(KeyValueName) And KeyValueName <> "" And lRetKey = ERROR_SUCCESS Then 'Check regsirty value for the opned key.
'            lRetVal = RegQueryValueEx(hKey, KeyValueName, 0, KeyType, ByVal 0, MyBufferSize)
'            RegCloseKey (hKey)
'            IsRegKey_ValueExist = lRetVal
'    ElseIf lRetKey = ERROR_SUCCESS Then 'Check Regsitry key
'            RegCloseKey (hKey)
'            IsRegKey_ValueExist = lRetKey
'    Else
'            IsRegKey_ValueExist = lRetKey
'    End If
'
'Exit Function
'
'localError:
'    LogError.LoggError Err, "IsRegKeyExist", "Unable to check Reg key or value existence"
'End Function
''function to enumerate all subkeys under a registry key
''rRet() is an array to append to , rRet must have at least on value(even empty)
''bEnumValues to indicate if we will enumerate all values under all enumerated keys.
'Public Function EnumerateRegKey(sKeyPaths() As String, ByRef rRet() As REG_VALUE)
'  On Error GoTo Loc_Error
'  Dim hKey As Long  ' receives a handle to the opened registry key
'  Dim ReturnValue As Long  ' return value
'  Dim Counter As Long
'  Dim MyBuffer As String
'  Dim MyBufferSize As Long
'  Dim ClassNameBuffer As String
'  Dim ClassNameBufferSize As Long
'  Dim LastWrite As FILETIME
'  Dim i As Integer, j As Long
'  Dim n(0) As String
'  DoEvents
'  ' open the registry key
'For i = LBound(sKeyPaths()) To UBound(sKeyPaths())
'  If RegGetKeyHandle(sKeyPaths(i), KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then GoTo NextLoop 'key is not exist or access denied
'  Counter = 0
'  ' loop until no more registry keys
'  Do Until ReturnValue <> 0
'        MyBuffer = Space(255)
'        ClassNameBuffer = Space(255)
'        MyBufferSize = 255
'        ClassNameBufferSize = 255
'        ReturnValue = RegEnumKeyEx(hKey, Counter, MyBuffer, MyBufferSize, ByVal 0, ClassNameBuffer, ClassNameBufferSize, LastWrite)
'      If ReturnValue = 0 Then
'            MyBuffer = Left$(MyBuffer, MyBufferSize)
'            ClassNameBuffer = Left$(ClassNameBuffer, ClassNameBufferSize)
'            j = UBound(rRet()) + 1
'            ReDim Preserve rRet(j)
'            rRet(j).sValues(0) = MyBuffer
'            rRet(j).iSerachType = 0
'       End If
'    Counter = Counter + 1
'  Loop
'  ' close the registry key
'  ReturnValue = RegCloseKey(hKey)
'NextLoop: Next i
'
'Exit Function
'Loc_Error:
'LogError.LoggError Err, "EnumerateRegKey"
'Resume Next 'go to the next key
'End Function
''function to enumerate all subkeys and their values under a registry key
''rRet() is an array to append to, rRet must have at least on value(even empty)
'Public Function EnumerateAllRegKeys_Values(sKeyPaths() As String, ByRef rRet() As REG_VALUE)
'  On Error GoTo Loc_Error
'  Dim hKey As Long  ' receives a handle to the opened registry key
'  Dim ReturnValue As Long  ' return value
'  Dim Counter As Long
'  Dim MyBuffer As String
'  Dim MyBufferSize As Long
'  Dim ClassNameBuffer As String
'  Dim ClassNameBufferSize As Long
'  Dim LastWrite As FILETIME
'  Dim i As Integer, j As Long
'  Dim n(0) As String
'  DoEvents
'  ' open the registry key
'For i = LBound(sKeyPaths()) To UBound(sKeyPaths())
'  If RegGetKeyHandle(sKeyPaths(i), KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then GoTo NextLoop 'key is not exist or access denied
'  Counter = 0
'  ' loop until no more registry keys
'  Do Until ReturnValue <> 0
'        MyBuffer = Space(255)
'        ClassNameBuffer = Space(255)
'        MyBufferSize = 255
'        ClassNameBufferSize = 255
'        ReturnValue = RegEnumKeyEx(hKey, Counter, MyBuffer, MyBufferSize, ByVal 0, ClassNameBuffer, ClassNameBufferSize, LastWrite)
'      If ReturnValue = 0 Then
'            MyBuffer = Left$(MyBuffer, MyBufferSize)
'            ClassNameBuffer = Left$(ClassNameBuffer, ClassNameBufferSize)
'            j = UBound(rRet()) + 1
'            ReDim Preserve rRet(j)
'            rRet(j).sValues(0) = sKeyPaths(i) + "\" + MyBuffer
'            rRet(j).iSerachType = 0
'            n(0) = rRet(j).sValues(0)
'            EnumRegValues n, rRet()
'            EnumerateAllRegKeys_Values n, rRet() 'Recursive: it slolws down the processing ???????????????????
'       End If
'    Counter = Counter + 1
'  Loop
'  ' close the registry key
'  ReturnValue = RegCloseKey(hKey)
'NextLoop: Next i
'
'Exit Function
'Loc_Error:
'LogError.LoggError Err, "EnumerateAllRegKeys_Values"
'Resume Next 'go to the next key
'End Function
''function to enumerate all keys and subkeys under a specific key
''rRet() is an array to append to, rRet must have at least on value(even empty)
'Public Function EnumerateAllRegKeys(sKeyPaths() As String, ByRef rRet() As REG_VALUE)
'  On Error GoTo Loc_Error
'  Dim hKey As Long  ' receives a handle to the opened registry key
'  Dim ReturnValue As Long  ' return value
'  Dim Counter As Long
'  Dim MyBuffer As String
'  Dim MyBufferSize As Long
'  Dim ClassNameBuffer As String
'  Dim ClassNameBufferSize As Long
'  Dim LastWrite As FILETIME
'  Dim i As Integer, j As Long
'  Dim n(0) As String
'  DoEvents
'  ' open the registry key
'For i = LBound(sKeyPaths()) To UBound(sKeyPaths())
'  If RegGetKeyHandle(sKeyPaths(i), KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then GoTo NextLoop 'key is not exist or access denied
'  Counter = 0
'  ' loop until no more registry keys
'  Do Until ReturnValue <> 0
'        MyBuffer = Space(255)
'        ClassNameBuffer = Space(255)
'        MyBufferSize = 255
'        ClassNameBufferSize = 255
'        ReturnValue = RegEnumKeyEx(hKey, Counter, MyBuffer, MyBufferSize, ByVal 0, ClassNameBuffer, ClassNameBufferSize, LastWrite)
'      If ReturnValue = 0 Then
'            MyBuffer = Left$(MyBuffer, MyBufferSize)
'            ClassNameBuffer = Left$(ClassNameBuffer, ClassNameBufferSize)
'            j = UBound(rRet()) + 1
'            ReDim Preserve rRet(j)
'            rRet(j).sValues(0) = sKeyPaths(i) + "\" + MyBuffer
'            rRet(j).iSerachType = 0
'            n(0) = rRet(j).sValues(0)
'            EnumerateAllRegKeys n, rRet()
'       End If
'    Counter = Counter + 1
'  Loop
'  ' close the registry key
'  ReturnValue = RegCloseKey(hKey)
'NextLoop: Next i
'
'Exit Function
'Loc_Error:
'LogError.LoggError Err, "EnumerateAllRegKeys"
'Resume Next 'go to the next key
'End Function
''function to enumerate all values under a registry key and all sub keys
''rRet() is an array to append to , rRet must have at least on value(even empty)
'Public Function EnumerateAllRegValues(sKeyPaths() As String, ByRef rRet() As REG_VALUE)
'  On Error GoTo Loc_Error
'  Dim hKey As Long  ' receives a handle to the opened registry key
'  Dim ReturnValue As Long  ' return value
'  Dim Counter As Long
'  Dim MyBuffer As String
'  Dim MyBufferSize As Long
'  Dim ClassNameBuffer As String
'  Dim ClassNameBufferSize As Long
'  Dim LastWrite As FILETIME
'  Dim i As Integer
'  Dim n(0) As String
'  DoEvents
'  ' open the registry key
'For i = LBound(sKeyPaths()) To UBound(sKeyPaths())
'  If RegGetKeyHandle(sKeyPaths(i), KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then GoTo NextLoop 'key is not exist or access denied
'  Counter = 0
'  ' loop until no more registry keys
'  Do Until ReturnValue <> 0
'        MyBuffer = Space(255)
'        ClassNameBuffer = Space(255)
'        MyBufferSize = 255
'        ClassNameBufferSize = 255
'        ReturnValue = RegEnumKeyEx(hKey, Counter, MyBuffer, MyBufferSize, ByVal 0, ClassNameBuffer, ClassNameBufferSize, LastWrite)
'      If ReturnValue = 0 Then
'            MyBuffer = Left$(MyBuffer, MyBufferSize)
'            ClassNameBuffer = Left$(ClassNameBuffer, ClassNameBufferSize)
'            n(0) = sKeyPaths(i) + "\" + MyBuffer
'            EnumRegValues n, rRet()
'            EnumerateAllRegValues n, rRet() 'Recursuve: it slolws down the processing ???????????????????
'       End If
'    Counter = Counter + 1
'  Loop
'  ' close the registry key
'  ReturnValue = RegCloseKey(hKey)
'NextLoop: Next i
'
'Exit Function
'Loc_Error:
'LogError.LoggError Err, "EnumerateAllRegValues"
'Resume Next 'go to the next key
'End Function
''Function to enumurate all reg values under a specific key....
''rRet must have at least on value(even empty)
'Public Function EnumRegValues(sPaths() As String, ByRef rRet() As REG_VALUE)
'Dim i As Integer, hKey As Long, index As Long
'Dim lKeyRoot As ROOT_KEYS
'Dim sSubkey As String
'Dim lType As Long, valueName As String, ValueValue As String, curidx As Long
'Dim arrData() As Byte, cbDataSize As Long
'
'On Error GoTo Loc_Error
'DoEvents
'For i = LBound(sPaths()) To UBound(sPaths())
'   If RegGetKeyHandle(sPaths(i), KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then GoTo NextLoopI 'the key is not exist..
'
'    curidx = 0
'    Do
'        cbDataSize = MAX_VALUE_SIZE
'        ReDim arrData(cbDataSize)
'        valueName = Space$(MAX_KEY_SIZE)
'        If RegEnumValue(hKey, curidx, valueName, Len(valueName), 0&, lType, arrData(0), cbDataSize) <> ERROR_SUCCESS Then Exit Do '
'        If cbDataSize < 1 Then cbDataSize = 1
'        ReDim Preserve arrData(cbDataSize - 1)
'        valueName = StripNulls(valueName)
'        ValueValue = StripNulls(GetRegData(lType, arrData))
'        index = UBound(rRet()) + 1
'        ReDim Preserve rRet(index)
'        rRet(index).sValues(0) = sPaths(i)
'        rRet(index).sValues(1) = valueName
'        rRet(index).sValues(2) = ValueValue
'        rRet(index).iSerachType = 2
'        curidx = curidx + 1
'        index = index + 1
'    Loop
'NextLoopI: Next i
'
'Exit Function
'Loc_Error:
'LogError.LoggError Err, "EnumRegValues", "Unable to enum all the common registry keys of:" + sPaths(i)
'Resume Next 'go to the next key
'End Function
'
'Public Function EnumRegItems(sPaths As ENUM_TYPES) As REG_VALUE()
'Dim Ret() As REG_VALUE
'ReDim rRet(0)
' EnumRegValues sPaths.sEnumValues(), Ret()
' EnumerateAllRegKeys_Values sPaths.sEnumAllKeysValues(), Ret()
' EnumerateRegKey sPaths.sEnumKeys(), Ret()
' EnumerateAllRegValues sPaths.sEnumAllValues(), Ret()
' EnumerateAllRegKeys sPaths.sEnumAllKeys(), Ret()
' EnumRegItems = Ret()
'End Function
'
''Delete all Subkeys with it's values for given Key
'Public Function DeleteKey(hRootKey As ROOT_KEYS, Optional sSubkey = "") As Long  '12-27-2003
'    Dim kiInfo As KeyInfo
'    Dim hSubkey As Long
'    Dim sPath As String
'    Dim sKeyName As String
'    Dim lErrorCode As Long, x As Integer
'
'    On Error GoTo localError
'
'    sPath = sSubkey
'    lErrorCode = RegDeleteKey(hRootKey, sSubkey)
'    If lErrorCode = ACCESS_DENIED Then
'
'    ' The key probably has subkeys. Windows NT will not allow a key to be deleted if
'    ' that key has subkeys. Another possibility is that the SAM is incorrect for the
'    ' key. Attempting to open the key will KEY_ALL_ACCESS will indicate which is the
'    ' case.
'        lErrorCode = RegOpenKeyEx(hRootKey, sPath, 0&, KEY_ALL_ACCESS, hSubkey)
'        If lErrorCode = ACCESS_DENIED Then
'           LogError.LoggError Err, "DeleteKey", "Access denied to delete registry key."
'        Else
'            GetKeyInfo hSubkey, kiInfo
'            For x = kiInfo.Subkeys To 1 Step -1
'                sKeyName = Space(500)
'                GetKeyName hSubkey, x - 1, sKeyName
'                DeleteKey hSubkey, sKeyName
'            Next x
'            DeleteKey hSubkey
'        End If
'        RegCloseKey hSubkey
'    End If
'    DeleteKey = lErrorCode
'
'    Exit Function
'localError:
'LogError.LoggError Err, "DeleteKey"
'End Function
'' DeleteValue from Registry
'Public Function DeleteRegKeyValue(KeyRoot As ROOT_KEYS, keyPath As String, Optional valueName As String = "") As Boolean
'    ' routine to delete a value from a key (but not the key) in the registry
'    On Error Resume Next
'    Dim hKey As Long  ' handle to the open registry key
'    Dim ReturnValue As Long  ' return value
'
'    ' First, open up the registry key which holds the value to delete.
'    ReturnValue = RegOpenKeyEx(KeyRoot, keyPath, 0, KEY_ALL_ACCESS, hKey)
'    If ReturnValue <> 0 Then
'        ' error encountered
'        DeleteRegKeyValue = False
'        ReturnValue = RegCloseKey(hKey)
'        Exit Function
'    End If
'
'    ' check to see if we are deleting a subkey or primary key
'    If valueName = "" Then valueName = keyPath
'
'    ' successfully opened registry key so delete the desired value from the key.
'    ReturnValue = RegDeleteValue(hKey, valueName)
'    If ReturnValue = 0 Then
'        ' no error encountered
'        DeleteRegKeyValue = True
'    Else
'        ' error encountered
'        DeleteRegKeyValue = False
'    End If
'    ' close the registry key
'    ReturnValue = RegCloseKey(hKey)
'End Function
'Public Function GetKeyInfo(hKey As Long, ByRef kiKeyInfo As KeyInfo) As Long
'Dim lErrorCode As Long
'' These variables are declared merely to make the function work. Windows 9x
'' does not use the values that are passed in these variables in calls to
'' RegQueryInfoKey.
'    Dim sClassName As String   ' class name of the key
'    Dim lLenClassName As Long  ' length of the key's class name
'    Dim lMaxLenClass As Long   ' length of the longest class name of the key's
'' subkeys
'    Dim lDescriptor As Long     ' security descriptor
'    Dim ftWriteTime As FILETIME ' last time this key was written
'
'    lErrorCode = RegQueryInfoKey(hKey, sClassName, lLenClassName, 0, kiKeyInfo.Subkeys, kiKeyInfo.LenSubkeys, lMaxLenClass, kiKeyInfo.Values, kiKeyInfo.LenValueNames, kiKeyInfo.LenValues, lDescriptor, ftWriteTime)
'GetKeyInfo = lErrorCode
'End Function
'
'Public Function GetKeyName(hKey As Long, ByVal lKeyNum As Long, sKeyName As String) As Long
'    Dim sClassName As String
'    Dim lLenClassName As Long
'    Dim ftWriteTime As FILETIME
'    Dim lNameLen As Long
'    Dim lErrorCode  As Long
'    lNameLen = Len(sKeyName)
'    lErrorCode = RegEnumKeyEx(hKey, lKeyNum, sKeyName, lNameLen, 0, sClassName, lLenClassName, ftWriteTime)
'    sKeyName = Left(sKeyName, lNameLen)
'    GetKeyName = lErrorCode
'End Function
''Function to check if a value is exist in a key values-values or values-names and deletes it
'Public Function DelValueIfExistInKeyValues(sRequiredValue As String, sFullKeyPath As String, iSearchType As SEARCH_VALUE_TYPE)
'Dim sPath(0) As String, rRet() As REG_VALUE
'Dim i  As Long
'Dim RemoveTmp As New RemoveEngine
'On Error GoTo Loc_Error
'
'sPath(0) = sFullKeyPath
'ReDim rRet(0)
'EnumRegValues sPath(), rRet()
''0=sKyPath,1=sKeyName,2=sKeyValue
'Select Case iSearchType
'    Case Value_Value:
'10:
'            For i = LBound(rRet()) To UBound(rRet())
'             If LCase(sRequiredValue) = LCase(rRet(i).sValues(2)) Then
'               RemoveTmp.RemoveRegKeyValue sFullKeyPath, rRet(i).sValues(1)
'             End If
'            Next i
'    Case Value_Name:
'20:
'            For i = LBound(rRet()) To UBound(rRet())
'             If LCase(sRequiredValue) = LCase(rRet(i).sValues(1)) Then
'               RemoveTmp.RemoveRegKeyValue sFullKeyPath, rRet(i).sValues(1)
'             End If
'            Next i
'    Case Value_Both:
'          GoTo 10
'          GoTo 20
'Set RemoveTmp = Nothing
'End Select
'
'Exit Function
'Loc_Error:
'LogError.LoggError Err, "DelValueIfExistInKeyValues"
'End Function
'
''----[ EnumKeys ]----'
'Public Function EnumKeys(ByVal sPath As String, Key() As String) As Long
'    Dim sName As String, RetVal As Long
'   DoEvents
'    hKey = GetKeys(sPath, sKey)
'
'    If (RegOpenKey(hKey, sKey, mainKey) = ERROR_SUCCESS) Then
'
'        EnumKeys = 0
'        sName = Space(255)
'        RetVal = Len(sName)
'
'        While RegEnumKeyEx(mainKey, EnumKeys, sName, RetVal, ByVal 0&, _
'                           vbNullString, ByVal 0&, ByVal 0&) <> ERROR_NO_MORE_ITEMS
'
'            ReDim Preserve Key(EnumKeys)
'
'            Key(EnumKeys) = Left$(sName, RetVal)
'
'            EnumKeys = EnumKeys + 1
'            sName = Space(255)
'            RetVal = Len(sName)
'
'        Wend
'
'        RegCloseKey mainKey
'    Else
'        EnumKeys = -1
'    End If
'
'End Function
'
''----[ EnumValues ]----'
'Private Function EnumValues(ByVal sPath As String, sName() As String, _
'                sValue() As Variant, Optional OnlyType As KeyType = -1) As Long
'    Dim mainKey As Long, rName As String, Cnt As Long
'    Dim rData As String, rType As Long, RetData As Long, RetVal As Long
'    DoEvents
'    hKey = GetKeys(sPath, sKey)
'
'    If RegOpenKey(hKey, sKey, mainKey) = ERROR_SUCCESS Then
'
'        Cnt = 0
'        rName = Space(255)
'        rData = Space(255)
'        RetVal = 255
'        RetData = 255
'
'        While RegEnumValue(mainKey, Cnt, rName, RetVal, 0, _
'                           rType, ByVal rData, RetData) <> ERROR_NO_MORE_ITEMS
'            'If RetData > 0 Then
'
'                If (OnlyType = -1) Or (OnlyType = rType) Then '
'
'                    ReDim Preserve sName(EnumValues) As String
'                    ReDim Preserve sValue(EnumValues) As Variant
'
'                    sName(EnumValues) = Left$(rName, RetVal)
'
'                    If (rType = REG_BINARY) Then '
'                        sValue(EnumValues) = ReadBinary(sPath, sName(EnumValues))
'                    ElseIf (rType = REG_DWORD) Then
'                        sValue(EnumValues) = ReadDWORD(sPath, sName(EnumValues))
'                    ElseIf (rType = REG_SZ) Then
'                        sValue(EnumValues) = ReadString(sPath, sName(EnumValues), "")
'                    End If
'
'                    EnumValues = EnumValues + 1
'
'                End If
'
'                Cnt = Cnt + 1
'                rName = Space(255)
'                rData = Space(255)
'                RetVal = 255
'                RetData = 255
'
'            'End If
'
'        Wend
'
'        RegCloseKey hKey
'    Else
'        EnumValues = -1
'    End If
'
'End Function
'
'Public Function ReadString(ByVal sPath As String, ByVal sName As String, _
'                           Optional sDefault As String = vbNullChar) As String
'    Dim sData As String, lDuz As Long
'
'    hKey = GetKeys(sPath, sKey)
'
'    If (RegOpenKeyEx(hKey, sKey, 0, KEY_READ, mainKey) = ERROR_SUCCESS) Then
'        sData = Space(255)
'        lDuz = Len(sData)
'        If (RegQueryValueEx(mainKey, sName, 0, REG_SZ, sData, lDuz) = ERROR_SUCCESS) Then
'            RegCloseKey mainKey
'            sData = Trim$(sData)
'            ReadString = Left$(sData, Len(sData) - 1)
'        Else
'                ReadString = sDefault
'        End If
'    Else
'        ReadString = sDefault
'    End If
'
'End Function
'Private Function GetKeys(sPath As String, sKey As String) As ROOT_KEYS
'Dim pos As Integer, mk As String
'
'    sPath = Replace$(sPath, "HKEY_CURRENT_USER", "HKCU", , , vbTextCompare)
'    sPath = Replace$(sPath, "HKEY_LOCAL_MACHINE", "HKLM", , , vbTextCompare)
'    sPath = Replace$(sPath, "HKEY_CLASSES_ROOT", "HKCR", , , vbTextCompare)
'    sPath = Replace$(sPath, "HKEY_USERS", "HKUS", , , vbTextCompare)
'    sPath = Replace$(sPath, "HKEY_PERFORMANCE_DATA", "HKPD", , , vbTextCompare)
'    sPath = Replace$(sPath, "HKEY_DYN_DATA", "HKDD", , , vbTextCompare)
'    sPath = Replace$(sPath, "HKEY_CURRENT_CONFIG", "HKCC", , , vbTextCompare)
'
'    pos = InStr(1, sPath, "\")
'
'    If (pos = 0) Then
'        mk = UCase$(sPath)
'        sKey = ""
'    Else
'        mk = UCase$(Left$(sPath, 4))
'        sKey = Right$(sPath, Len(sPath) - pos)
'    End If
'
'    Select Case mk
'        Case "HKCU": GetKeys = HKEY_CURRENT_USER
'        Case "HKLM": GetKeys = HKEY_LOCAL_MACHINE
'        Case "HKCR": GetKeys = HKEY_CLASSES_ROOT
'        Case "HKUS": GetKeys = HKEY_USERS
'        Case "HKPD": GetKeys = HKEY_PERFORMANCE_DATA
'        Case "HKDD": GetKeys = HKEY_DYN_DATA
'        Case "HKCC": GetKeys = HKEY_CURRENT_CONFIG
'    End Select
'
'End Function
'
'Public Function ReadBinary(ByVal sPath As String, ByVal sName As String, _
'                           Optional sDefault As String = vbNullString) As String
'
'    Dim lDuz As Long, sData As String
'
'    hKey = GetKeys(sPath, sKey)
'
'    If (RegOpenKeyEx(hKey, sKey, 0, KEY_READ, mainKey) = ERROR_SUCCESS) Then
'        lDuz = 1
'        RegQueryValueEx mainKey, sName, 0, REG_BINARY, 0, lDuz
'        sData = Space(lDuz)
'        If (RegQueryValueEx(mainKey, sName, 0, REG_BINARY, sData, lDuz) = ERROR_SUCCESS) Then
'            RegCloseKey mainKey
'            ReadBinary = Trim$(BinToStr(sData))
'        Else
'            ReadBinary = sDefault
'        End If
'    Else
'        ReadBinary = sDefault
'    End If
'
'End Function
'
'Public Function ReadDWORD(ByVal sPath As String, ByVal sName As String, _
'                         Optional lDefault As Long = -1) As Long
'    Dim lData As Long
'
'    hKey = GetKeys(sPath, sKey)
'
'    If (RegOpenKeyEx(hKey, sKey, 0, KEY_READ, mainKey) = ERROR_SUCCESS) Then
'        If (RegQueryValueEx(mainKey, sName, 0, REG_DWORD, lData, 4) = ERROR_SUCCESS) Then
'            RegCloseKey mainKey
'            ReadDWORD = lData
'        Else
'            ReadDWORD = lDefault
'        End If
'    Else
'        ReadDWORD = lDefault
'    End If
'End Function
'Public Function BinToStr(sStr As String) As String
'    'prebacuje npr. "¾>ÿ«" u "BE 3E FF AB" - koristi se za binary vrednost
'    Dim bs As String, Ret As String, q As Integer, tStr As String
'
'    Ret = vbNullString
'    For q = 1 To Len(sStr)
'        bs = Mid$(sStr, q, 1)
'        If bs = vbNullChar Then tStr = "00" Else tStr = CStr(Hex(Asc(bs)))
'        If (Len(tStr) = 1) Then tStr = tStr & "0"
'        Ret = Ret & " " & tStr
'    Next
'    BinToStr = Ret
'End Function
'
'Public Function GetKeyType(k As String) As ROOT_KEYS
'Dim i As ROOT_KEYS
'Select Case UCase(Mid(k, 1, InStr(k, "\") - 1))
'Case "HKEY_ALL": GetKeyType = HKEY_ALL
'Case "HKEY_CLASSES_ROOT": GetKeyType = HKEY_CLASSES_ROOT
'Case "HKEY_CURRENT_CONFIG": GetKeyType = HKEY_CURRENT_CONFIG
'Case "HKEY_CURRENT_USER": GetKeyType = HKEY_CURRENT_USER
'Case "HKEY_DYN_DATA": GetKeyType = HKEY_DYN_DATA
'Case "HKEY_LOCAL_MACHINE": GetKeyType = HKEY_LOCAL_MACHINE
'Case "HKEY_PERFORMANCE_DATA": GetKeyType = HKEY_PERFORMANCE_DATA
'Case "HKEY_USERS": GetKeyType = HKEY_USERS
'End Select
'End Function
'
''Function to export reg item, it Should be called by ExportRegPath
'Public Function WriteReg(ByVal sKeyName As String, ByVal sKeyValue As String, ByVal fn As Integer, rKeyType As KeyType) As Boolean
'    On Error Resume Next
'    Dim h As Long
'    Dim Temp() As String: Dim ii As Integer
'    Dim keyName() As String, aName() As String, aValue() As Variant, x As Integer
'    Dim u As Long, tmp As String, opened As Boolean, L As Long
'    Dim hasOutput As Boolean, nuls As String
'    Dim lpKey As Long
'
' Select Case rKeyType
'   Case REG_SZ
'       Print #fn, Chr$(34) & Replace(sKeyName, "\", "\\") & Chr$(34) & " = " & "" _
'                           ; Chr$(34) & Replace(sKeyValue, "\", "\\") & Chr$(34)
'   Case REG_BINARY
'       Print #fn, Chr$(34) & Replace(sKeyName, "\", "\\") & "=hex:" & _
'                                               Replace(Trim$(sKeyValue), " ", ",")
'   Case REG_DWORD
'       tmp = Hex$(sKeyValue)
'            If (Len(tmp) < 8) Then
'                nuls = ""
'                For x = 1 To 8 - Len(tmp)
'                    nuls = nuls & "0"
'                Next
'                tmp = nuls & tmp
'            End If
'
'         Print #fn, Chr$(34) & Replace(sKeyName, "\", "\\") & Chr$(34) & "=" & _
'               Chr$(34) & "=dword:" & tmp
'
'        Print #fn, ""
' End Select
''Close #fn
'End Function
''function to export a registry path to a .reg file to be recovered later by launching it.
''rRet() is an array to append to
'Public Function ExportRegPath(sKeyPath As String, sRegFilePath As String, _
'                              Optional bkeyValue As Boolean, Optional sKeyValueName As String)
'  On Error GoTo Loc_Error
'  Dim hKey As Long  ' receives a handle to the opened registry key
'  Dim ReturnValue As Long  ' return value
'  Dim Counter As Long, arrData() As Byte
'  Dim MyBuffer As String
'  Dim MyBufferSize As Long
'  Dim ClassNameBuffer As String
'  Dim ClassNameBufferSize As Long
'  Dim LastWrite As FILETIME
'  Dim i As Integer, j As Long
'  Dim n(0) As String, cbDataSize As Long
'  Dim fn As Integer, curidx As Long
'  Dim rRet() As REG_VALUE, lType As Long
'  Dim ValueValue As String, valueName As String
'  Dim sBuf1 As String, lBuf1Size As Long
'
'  ' open the registry key
'  If RegGetKeyHandle(sKeyPath, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then GoTo Loc_Error 'key is not exist or access denied
'  Counter = 0
'  ' loop until no more registry keys
'    fn = FreeFile
'    Open sRegFilePath For Append As #fn 'Append
'
'    Print #fn, "[" & sKeyPath & "]"
'    If (ReadString(sKeyPath, "") <> vbNullChar) And Not bkeyValue Then '(Default)
'        Print #fn, "@=" & Chr$(34) & ReadString(sKeyPath, "", "") & Chr$(34)
'    End If
'
'    If bkeyValue Then
'      RegCloseKey hKey
'      lBuf1Size = 255
'      sBuf1 = Space$(lBuf1Size)
'      sBuf1 = GetKeyNameValue(sKeyPath, sKeyValueName, lType)
'      If sBuf1 <> "" Then
'        WriteReg sKeyValueName, StripNulls(sBuf1), fn, lType
'      End If
'      Close #fn
'    Exit Function
'   End If
'
'  Do Until ReturnValue <> 0
'        MyBuffer = Space(255)
'        ClassNameBuffer = Space(255)
'        MyBufferSize = 255
'        ClassNameBufferSize = 255
'        ReturnValue = RegEnumKeyEx(hKey, Counter, MyBuffer, MyBufferSize, ByVal 0, ClassNameBuffer, ClassNameBufferSize, LastWrite)
'      If ReturnValue = 0 Then
'            MyBuffer = Left$(MyBuffer, MyBufferSize)
'            ClassNameBuffer = Left$(ClassNameBuffer, ClassNameBufferSize)
'            j = UBound(rRet()) + 1
'            ReDim Preserve rRet(j)
'            rRet(j).sValues(0) = sKeyPath + "\" + MyBuffer
'            rRet(j).iSerachType = 0
'            n(0) = rRet(j).sValues(0)
'
'            ExportRegPath rRet(j).sValues(0), sRegFilePath, False, "" 'Recursive: it slolws down the processing ???????
'       Else
'          RegCloseKey hKey
'        If RegGetKeyHandle(sKeyPath, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then GoTo Loc_Error 'the key is not exist..
'
'            curidx = 0
'            Do
'                cbDataSize = MAX_VALUE_SIZE
'                ReDim arrData(cbDataSize)
'                valueName = Space$(MAX_KEY_SIZE)
'                If RegEnumValue(hKey, curidx, valueName, Len(valueName), 0&, lType, arrData(0), cbDataSize) <> ERROR_SUCCESS Then Exit Do '
'                If cbDataSize < 1 Then cbDataSize = 1
'                ReDim Preserve arrData(cbDataSize - 1)
'                valueName = StripNulls(valueName)
'                ValueValue = StripNulls(GetRegData(lType, arrData))
'                WriteReg valueName, ValueValue, fn, lType
'                curidx = curidx + 1
'            Loop
'            Close #fn
'       End If
'    Counter = Counter + 1
'  Loop
'  ' close the registry key
'  ReturnValue = RegCloseKey(hKey)
'Close #fn
'
'Exit Function
'Loc_Error:
'LogError.LoggError Err, "EnumerateAllRegKeys_Values"
'Resume Next 'go to the next key
'End Function
'
'
'
