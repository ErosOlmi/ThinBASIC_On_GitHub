  'MESSAGE http://www.powerbasic.com/support/forums/Forum7/HTML/000913.html
  'FORUM:  Source Code
  'TOPIC:  Registry.inc - easy general purpose registry access
  'NAME:   Wayne Diamond, Member
  'DATE:   March 01, 2001 12:01 PM
  
  'REGISTRY.INC
  'General Purpose Registry Access
  'Uses: win32api.inc
  
  'Functions:
  ' _DelRegValue(lpKey As Long,ByVal cMainkey As String, ByVal Key As String) As Long
  ' _GetRegValue(lpKey As Long,ByVal cMainkey As String, ByVal Key As String) As String
  ' _SetRegValue(lpKey As Long,ByVal cMainkey As String, ByVal Key As String, ByVal Setting As String) As Long
  ' _DelRegKey(lpKey As Long, Key As String)  As Long
  
'  FUNCTION xDelRegValue(lpKey AS LONG,BYVAL cMainkey AS STRING, BYVAL Key AS STRING) AS LONG
'    ON ERROR RESUME NEXT
'    LOCAL RetCode AS LONG
'    LOCAL hKey AS LONG
'    DIM acMainkey AS ASCIIZ * 300
'
'    acMainkey = cMainkey
'    RetCode = RegOpenKeyEx(lpKey, acMainkey, 0&, %KEY_ALL_ACCESS, hKey)
'    IF RetCode = %ERROR_SUCCESS THEN
'      IF Key$ = "*" THEN Key$ = CHR$(0,0)
'      RetCode = RegDeleteValue(hKey, BYVAL STRPTR(Key$))
'    END IF
'    RegCloseKey hKey
'    FUNCTION = RetCode
'  END FUNCTION
  
'  FUNCTION xGetRegValue(lpKey AS LONG,BYVAL cMainkey AS STRING, BYVAL Key AS STRING) AS STRING
'    ON ERROR RESUME NEXT
'    DIM RetCode AS LONG
'    DIM hKey AS LONG
'    DIM KeyNameA AS ASCIIZ * 256
'    LOCAL zTmp AS ASCIIZ * 256
'    DIM acMainKey AS ASCIIZ * 300
'    LOCAL ZZZ AS STRING
'    LOCAL cbData AS LONG
'    LOCAL KeyType AS LONG
'
'    acMainKey = cMainKey
'    RetCode = RegOpenKeyEx(lpKey, acMainkey, 0&, %KEY_ALL_ACCESS, hKey)
'    IF RetCode = %ERROR_SUCCESS THEN
'      IF Key$ = "*" THEN Key$ = CHR$(0,0)
'      'szdat=256
'      DIM zbuffer AS ASCIIZ*256
'      KeyNameA = Key
'      cbData = SIZEOF(zTmp)
'      RetCode = RegQueryValueEx(BYVAL hKey, KeyNameA, BYVAL 0, KeyType, zTmp, cbData)
'      ZZZ = zTmp
'      FUNCTION = ZZZ
'      EXIT FUNCTION
'    END IF
'    FUNCTION = ""
'  END FUNCTION
  
'  FUNCTION xSetRegValue(lpKey AS LONG,BYVAL cMainkey AS STRING, BYVAL Key AS STRING, BYVAL Setting AS STRING) AS LONG
'    ON ERROR RESUME NEXT
'    LOCAL hKey   AS LONG
'    LOCAL Result AS LONG
'    LOCAL zText  AS ASCIIZ * 2048
'
'    IF Key$ = "*" THEN
'      Key$ = CHR$(0,0)
'    END IF
'    IF RegCreateKeyEx(lpKey, cMainKey + CHR$(0),0, "", %REG_OPTION_NON_VOLATILE, _
'                       %KEY_ALL_ACCESS, BYVAL %NULL, hKey, Result) <> %ERROR_SUCCESS THEN
'      FUNCTION = 0
'      EXIT FUNCTION
'    END IF
'    zText = Setting
'    IF LEN(Setting) THEN
'      RegSetValueEx hKey, Key+CHR$(0), 0, %REG_SZ, zText, LEN(Setting)+1
'    ELSE
'      RegSetValueEx hKey, Key+CHR$(0), 0, %REG_SZ, zText, 1
'    END IF
'    RegCloseKey hKey
'    FUNCTION = 0
'  END FUNCTION
  
'  FUNCTION xDelRegKey(lpKey AS LONG, Key AS STRING)  AS LONG
'    ON ERROR RESUME NEXT
'    LOCAL zStrKey AS ASCIIZ * 255, hKey AS LONG, Result AS LONG
'
'    zStrKey = Key
'    IF RegOpenKeyEx(BYVAL lpKey,zStrKey, 0, %KEY_ALL_ACCESS, hKey) <> %ERROR_SUCCESS THEN
'      EXIT FUNCTION
'    END IF
'    Result = RegDeleteKey(lpKey, zStrKey)
'    RegCloseKey hKey
'    FUNCTION = Result
'  END FUNCTION

'  SUB RerunAfterReboot(WhatToDo AS LONG)
'    LOCAL hKey AS LONG, result AS LONG
'    LOCAL szApplication AS ASCIIZ * %MAX_PATH, szKeyName AS ASCIIZ * 18
'  
'    ' create a registry key (if it already exists, it will be opened)
'    IF RegCreateKeyEx(%HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", _
'                      0&, "", %REG_OPTION_NON_VOLATILE, %KEY_ALL_ACCESS, BYVAL %NULL, _
'                      hKey, result) <> %ERROR_SUCCESS THEN
'      EXIT SUB                                         ' quit when the key cannot be created/opened
'    END IF
'  
'    SELECT CASE result
'      CASE %REG_CREATED_NEW_KEY
'        'MSGBOX "Key created", 64, "Info"
'      CASE %REG_OPENED_EXISTING_KEY
'        'MSGBOX "Key opened (it already exists)", 64, "Info"
'    END SELECT
'  
'    szKeyName = "thinBasic " + DATE$ + TIME$            ' use a unique name for the key
'    IF ISFALSE WhatToDo THEN                           ' when closing your app. always take care...
'      RegDeleteValue hKey, szKeyName                   ' ... to delete the registry entry
'    ELSE
'      GetModuleFileName 0, szApplication, %MAX_PATH    ' get path and name of this app.
'      RegSetValueEx hKey, szKeyName, 0&, %REG_SZ, szApplication, SIZEOF(szApplication)
'    END IF
'    RegCloseKey hKey
'  END SUB

%KEY_WOW64_32KEY = &H0200???
%KEY_WOW64_64KEY = &H0100???

  FUNCTION GetRegValues(hKey AS DWORD) AS STRING
    DIM lReturn            AS LONG
    DIM KeyName            AS ASCIIZ * 255
    DIM KeyType            AS DWORD
    DIM KeyValue           AS ASCIIZ * 255
    DIM Reserved           AS DWORD
    DIM lIndex             AS DWORD
    DIM lNameLen           AS DWORD
    DIM lValueLen          AS DWORD
    DIM TextOut3           AS STRING
    DIM A                  AS STRING
    DIM lSIZE              AS LONG
    DIM OutText(200)       AS STRING
    DIM Count              AS INTEGER
    DIM I                  AS INTEGER
  
    lIndex   = 0
    Textout3 = ""
    lReturn  = 0
    Count    = 0
    
    Do While lReturn <> 259 '%ERROR_NO_MORE_ITEMS  '(259)
       KeyName   = STRING$(255,32)
       KeyValue  = STRING$(255,32)
       lNameLen  = 255 'LEN(KeyName)
       lValueLen = 255 'LEN(KeyValue)
  
       Reserved=%NULL
       LReturn = RegEnumValue(hKey, lIndex, KeyName, BYVAL VARPTR(lNameLen), BYVAL Reserved , BYVAL VARPTR (KeyType), KeyValue, BYVAL VARPTR(lValueLen))
       IF lReturn = 87 THEN
          FUNCTION = "Error code 87 in RegEnumValue" + $CRLF
          EXIT FUNCTION
       END IF
       IF (lReturn <> %ERROR_NO_MORE_ITEMS) THEN
          IF lReturn = %ERROR_SUCCESS THEN
              lSIZE = lValueLen
              IF KeyType = %REG_SZ THEN
                 A= LEFT$(KeyValue, lSIZE - 1)
              ELSEIF KeyType = %REG_DWORD THEN
                 A= STR$(CVL(LEFT$(KeyValue, lSIZE)))
              ELSEIF KeyType = %REG_BINARY AND lSIZE=1 THEN
                 A= STR$(ASC(LEFT$(KeyValue, lSIZE)))
              ELSEIF KeyType = %REG_MULTI_SZ THEN
                 A= LEFT$(KeyValue, lSIZE - 1)
                 REPLACE ANY CHR$(0) WITH "," IN A
              ELSE
                 A= LEFT$(KeyValue, lSIZE)
                 REPLACE ANY CHR$(0, 1, 2, 3, 4, 5, 6, 7, 8, 9) WITH "*" IN A
              END IF
              count = count + 1
              OutText(Count) = KeyName + CHR$(255) + A
          ELSE
          END IF
       END IF
       INCR lIndex
    LOOP
    IF count > 0 THEN
       ARRAY SORT OutText(1) FOR count, FROM 1 TO 15
       FOR I = 1 TO count
           KeyName = EXTRACT$(OutText(I), CHR$(255))
           A = REMAIN$(OutText(I), CHR$(255))
           TextOut3 = TextOut3 + KeyName + "=" + A + $CRLF
       NEXT I
    END IF
    FUNCTION = TextOut3
  END FUNCTION
  
  Function GetAllKeys(ByVal lHKey As Dword, lHive As Asciiz, ByVal sSep As String) As String 
    DIM lReturn            AS LONG
    Dim szSubKeyName       As Asciiz * 1024
    Dim azClass            As Asciiz * 1024
    DIM Reserved           AS DWORD
    DIM LastWriteTime      AS FILETIME
    DIM lIndex             AS DWORD
    DIM lNameLen           AS DWORD
    DIM lClassLen          AS LONG
    DIM hKey               AS DWORD
    DIM A                  AS STRING
    DIM strTmp             AS STRING
  
    lNamelen = 255
    Reserved = 0
    lIndex   = 0
  
    lReturn = RegOpenKeyEx(lHKey, lHive, 0, BYVAL %KEY_READ, hKey)

'MsgBox Str$(lReturn)   & Str$(%ERROR_SUCCESS)

    IF lReturn <> %ERROR_SUCCESS THEN
        IF lReturn = %ERROR_FILE_NOT_FOUND THEN
           FUNCTION = ""
           EXIT FUNCTION
        END IF
        A = "Open Key Fails for " + lhive + " Error=" + STR$(lReturn) + $CRLF
        IF lReturn = %ERROR_ACCESS_DENIED THEN
           a = a + " Access Denied to Registry"
        END IF
        MSGBOX " " + A, , "Error"
        FUNCTION = A + $CRLF
        EXIT FUNCTION
    END IF
  
    strTmp = strTmp + GetRegValues(hKey)

'MsgBox strTmp
  
    ' Enumerate all the subkeys for the current key.
    '...............................................
    Do While  lReturn <> %ERROR_NO_MORE_ITEMS
       If lReturn = %ERROR_SUCCESS Then
          lNameLen = SizeOf(szSubKeyName)
          lClassLen = SizeOf(azClass)
          lReturn = RegEnumKeyEx(hKey, lIndex, szSubKeyName, ByVal VarPtr(lNameLen), Reserved, ByVal VarPtr(azClass), ByVal VarPtr(lClasslen), LastWriteTime)
'MsgBox Str$(lReturn)   & Str$(%ERROR_NO_MORE_ITEMS)
          If (lReturn <> %ERROR_NO_MORE_ITEMS) Then
             If lReturn = %ERROR_SUCCESS Then
                strTmp += iif$(len(strTmp), sSep, "") + szSubKeyName '+ sSep '$LIST_SEP
'MsgBox strTmp
             End If
          End If
       End If
       lIndex = lIndex + 1
'MsgBox "Index" & Str$(lIndex)       
'MsgBox Str$(lReturn)   & Str$(%ERROR_NO_MORE_ITEMS)
    Loop

'MsgBox "Before Close"
    RegCloseKey hKey
'MsgBox "After Close"

'    ' Enumerate all the subkeys for the current key.
'    '...............................................
'    Do While  lReturn = %ERROR_SUCCESS
'      If lReturn = %ERROR_SUCCESS Then
'        lNameLen = SizeOf(szSubKeyName)
'        lClassLen = SizeOf(azClass)
'        lReturn = RegEnumKeyEx(hKey, lIndex, szSubKeyName, ByVal VarPtr(lNameLen), Reserved, ByVal VarPtr(azClass), ByVal VarPtr(lClasslen), LastWriteTime)
''MsgBox Str$(lReturn)   & Str$(%ERROR_NO_MORE_ITEMS)
'        If lReturn = %ERROR_SUCCESS Then
'          strTmp += szSubKeyName + sSep '$LIST_SEP
''MsgBox strTmp
''MsgBox "Index" & Str$(lIndex)       
'          lIndex = lIndex + 1
'        End If
'      End If
''MsgBox Str$(lReturn)   & Str$(%ERROR_NO_MORE_ITEMS)
'    Loop

'MsgBox "Before Close"
    RegCloseKey hKey
'MsgBox "After Close"

    Function = strTmp

  End Function

    
'  '############################ REPLY ############################
'  Wayne Diamond, Member
'  posted March 01, 2001 12:03 PM
'  Example usage
'  #COMPILE EXE  'compiles to just 8704 bytes - registry.inc doesn't add much overhead
'  #INCLUDE "win32api.inc"
'  #INCLUDE "registry.inc"
'
'  FUNCTION PBMAIN() AS LONG
'  'Create new value in new key
'   IF SetRegValue(%HKEY_LOCAL_MACHINE, "SOFTWARE\RegistryTest", "NewSetting", "Data in key") = 0 THEN
'      MSGBOX "Set registry setting!"
'   ELSE
'      MSGBOX "Couldnt set registry setting!"
'   END IF
'  
'   'Read new value
'   DIM RegData AS STRING
'   RegData = GetRegValue(%HKEY_LOCAL_MACHINE, "SOFTWARE\RegistryTest", "NewSetting")
'   MSGBOX "Data = " & RegData
'  
'   'Delete the new value
'   IF DelRegValue(%HKEY_LOCAL_MACHINE, "SOFTWARE\RegistryTest", "NewSetting") = 0 THEN
'      MSGBOX "Deleted the value!"
'   ELSE
'      MSGBOX "Couldn't delete the value!"
'   END IF
'  
'   'Delete the new key
'   IF DelRegKey(%HKEY_LOCAL_MACHINE, "SOFTWARE\RegistryTest") = 0 THEN
'      MSGBOX "Deleted the key!"
'   ELSE
'      MSGBOX "Couldn't delete the key!"
'   END IF
'  END FUNCTION
'  OR IN it's simplest usage form:
'  
'  #COMPILE EXE "regtiny.exe"
'  #INCLUDE "win32api.inc"
'  #INCLUDE "registry.inc"
'  FUNCTION PBMAIN() AS LONG
'   SetRegValue %HKEY_LOCAL_MACHINE, "SOFTWARE\RegistryTest", "NewSetting", "Data in key"  'write
'   MSGBOX "You added: " & GetRegValue(%HKEY_LOCAL_MACHINE, "SOFTWARE\RegistryTest", "NewSetting") 'read
'  END FUNCTION 
                                             

'------------------------------------------------------------------------------
'REGISTRY.INC
'General Purpose Registry Access

'DECLARE FUNCTION SetRegValue        (BYVAL hKey AS LONG, zSubKey AS ASCIIZ, zValueName AS ASCIIZ, zValue AS ASCIIZ) AS LONG
'DECLARE FUNCTION SetRegDwordValue   (BYVAL hKey AS LONG, zSubKey AS ASCIIZ, zValueName AS ASCIIZ, BYVAL dwValue AS DWORD) AS LONG
'DECLARE FUNCTION SetRegBinValue     (BYVAL hKey AS LONG, zSubKey AS ASCIIZ, zValueName AS ASCIIZ, BYVAL BinValuePtr AS LONG, BYVAL lBufferSize AS LONG) AS LONG
'DECLARE FUNCTION SetRegTxtNumValue  (BYVAL hKey AS LONG, zSubKey AS ASCIIZ, zValueName AS ASCIIZ, BYVAL lValue AS LONG) AS LONG
'DECLARE FUNCTION SetRegTxtBoolValue (BYVAL hKey AS LONG, zSubKey AS ASCIIZ, zValueName AS ASCIIZ, BYVAL lValue AS LONG) AS LONG
'
'DECLARE FUNCTION GetRegValue        (BYVAL hKey AS LONG, zSubKey AS ASCIIZ, zValueName AS ASCIIZ) AS STRING
'DECLARE FUNCTION GetRegDwordValue   (BYVAL hKey AS LONG, zSubKey AS ASCIIZ, zValueName AS ASCIIZ) AS DWORD
'DECLARE FUNCTION GetRegBinValue     (BYVAL hKey AS LONG, zSubKey AS ASCIIZ, zValueName AS ASCIIZ, BYVAL BinValuePtr AS LONG, BYVAL lBufferSize AS LONG) AS LONG
'DECLARE FUNCTION GetRegTxtNumValue  (BYVAL hKey AS LONG, zSubKey AS ASCIIZ, zValueName AS ASCIIZ, BYVAL lDefVal AS LONG) AS LONG
'DECLARE FUNCTION GetRegTxtBoolValue (BYVAL hKey AS LONG, zSubKey AS ASCIIZ, zValueName AS ASCIIZ, BYVAL lDefVal AS LONG) AS LONG
'
'DECLARE FUNCTION DelRegValue        (BYVAL hKey AS LONG, zSubKey AS ASCIIZ, zValueName AS ASCIIZ) AS LONG
'DECLARE FUNCTION DelRegKey          (BYVAL hKey AS LONG, zSubKey AS ASCIIZ) AS LONG

' **************************************************************************
' * SetRegValue
' * Set a string registry value
' *
' * hKey        - One of the following equates:
' *                %HKEY_CLASSES_ROOT,
' *                %HKEY_CURRENT_CONFIG
' *                %HKEY_CURRENT_USER
' *                %HKEY_LOCAL_MACHINE
' *                %HKEY_USERS
' * szSubKey    - Name of a subkey. This parameter cannot be NULL
' * szValueName - Name of the value to set
' * szValue     - Data to be stored
' **************************************************************************
FUNCTION SetRegValue(BYVAL hKey AS LONG, szSubKey AS ASCIIZ, szValueName AS ASCIIZ, szValue AS ASCIIZ) AS LONG
  LOCAL lRet      AS LONG
  LOCAL hOpenKey  AS LONG

  IF RegCreateKeyEx(hKey, szSubKey, %NULL, "", %REG_OPTION_NON_VOLATILE, _
                    %KEY_WRITE, BYVAL %NULL, hOpenKey, lRet) <> %ERROR_SUCCESS THEN
    EXIT FUNCTION
  END IF

  lRet = RegSetValueEx(hOpenKey, szValueName, %NULL, %REG_SZ, szValue, LEN(szValue) + 1)
  CALL RegCloseKey(hOpenKey)

  IF lRet = %ERROR_SUCCESS THEN
    FUNCTION = %TRUE
  END IF
END FUNCTION

' **************************************************************************
' * SetRegDwordValue
' * Set a binary registry value - very usefull to save UDTs to the registry
' *
' * hKey         - One of the following equates:
' *                %HKEY_CLASSES_ROOT,
' *                %HKEY_CURRENT_CONFIG
' *                %HKEY_CURRENT_USER
' *                %HKEY_LOCAL_MACHINE
' *                %HKEY_USERS
' * szSubKey     - Name of a subkey. This parameter cannot be NULL
' * szValueName  - Name of the value to set
' * dwValue      - Dword value to store
' **************************************************************************
FUNCTION SetRegDwordValue(BYVAL hKey AS LONG, szSubKey AS ASCIIZ, szValueName AS ASCIIZ, BYVAL dwValue AS DWORD) AS LONG
  LOCAL lRet     AS LONG
  LOCAL hOpenKey    AS LONG

  IF RegCreateKeyEx( hKey, szSubKey, %NULL, "", %REG_OPTION_NON_VOLATILE, _
                    %KEY_WRITE, BYVAL %NULL, hOpenKey, lRet) <> %ERROR_SUCCESS THEN
    EXIT FUNCTION
  END IF

  lRet = RegSetValueEx(hOpenKey, szValueName, %NULL, %REG_DWORD, dwValue, SIZEOF(dwValue))
  CALL RegCloseKey(hOpenKey)

  IF lRet = %ERROR_SUCCESS THEN
    FUNCTION = %TRUE
  END IF
END FUNCTION

' **************************************************************************
' * SetRegBinValue
' * Set a binary registry value - very usefull to save UDTs to the registry
' *
' * hKey        - One of the following equates:
' *                %HKEY_CLASSES_ROOT,
' *                %HKEY_CURRENT_CONFIG
' *                %HKEY_CURRENT_USER
' *                %HKEY_LOCAL_MACHINE
' *                %HKEY_USERS
' * zSubKey     - Name of a subkey. This parameter cannot be NULL
' * zValueName  - Name of the value to set
' * BinValuePtr - Pointer to a buffer containing the data to be stored
' * lBufferSize - Specifies the size, in bytes, of the buffer
' **************************************************************************
FUNCTION SetRegBinValue(BYVAL hKey AS LONG, szSubKey AS ASCIIZ, szValueName AS ASCIIZ, BYVAL BinValuePtr AS LONG, BYVAL lBufferSize AS LONG) AS LONG
  LOCAL lRet      AS LONG
  LOCAL hOpenKey  AS LONG

  IF RegCreateKeyEx(hKey, szSubKey, %NULL, "", %REG_OPTION_NON_VOLATILE, _
                    %KEY_WRITE, BYVAL %NULL, hOpenKey, lRet) <> %ERROR_SUCCESS THEN
    EXIT FUNCTION
  END IF

  lRet = RegSetValueEx(hOpenKey, szValueName, %NULL, %REG_BINARY, BYVAL BinValuePtr, lBufferSize)
  CALL RegCloseKey(hOpenKey)

  IF lRet = %ERROR_SUCCESS THEN
    FUNCTION = %TRUE
  END IF
END FUNCTION

' **************************************************************************
' * SetRegTxtNumValue
' * Set a registry value
' *
' * hKey        - One of the following equates:
' *                %HKEY_CLASSES_ROOT,
' *                %HKEY_CURRENT_CONFIG
' *                %HKEY_CURRENT_USER
' *                %HKEY_LOCAL_MACHINE
' *                %HKEY_USERS
' * zSubKey     - Name of a subkey. This parameter cannot be NULL
' * zValueName  - Name of the value to query
' **************************************************************************
FUNCTION SetRegTxtNumValue(BYVAL hKey AS LONG, szSubKey AS ASCIIZ, szValueName AS ASCIIZ, BYVAL lValue AS LONG) AS LONG
  LOCAL lRet        AS LONG
  LOCAL hOpenKey    AS LONG
  LOCAL szTmp       AS ASCIIZ * 64

  FUNCTION = %FALSE
  IF RegCreateKeyEx( hKey, szSubKey, %NULL, "", %REG_OPTION_NON_VOLATILE, _
                    %KEY_WRITE, BYVAL %NULL, hOpenKey, lRet) <> %ERROR_SUCCESS THEN
    EXIT FUNCTION
  END IF

  szTmp = FORMAT$(lValue)
  lRet = RegSetValueEx(hOpenKey, szValueName, %NULL, %REG_SZ, szTmp, LEN(szTmp)+1)
  CALL RegCloseKey(hOpenKey)

  IF lRet = %ERROR_SUCCESS THEN
    FUNCTION = %TRUE
  END IF
END FUNCTION

' **************************************************************************
' * SetRegTxtBoolValue
' * Set a registry value
' *
' * hKey        - One of the following equates:
' *                %HKEY_CLASSES_ROOT,
' *                %HKEY_CURRENT_CONFIG
' *                %HKEY_CURRENT_USER
' *                %HKEY_LOCAL_MACHINE
' *                %HKEY_USERS
' * zSubKey     - Name of a subkey. This parameter cannot be NULL
' * zValueName  - Name of the value to query
' **************************************************************************
FUNCTION SetRegTxtBoolValue(BYVAL hKey AS LONG, szSubKey AS ASCIIZ, szValueName AS ASCIIZ, BYVAL lValue AS LONG) AS LONG
  LOCAL lRet     AS LONG
  LOCAL hOpenKey    AS LONG
  LOCAL szTmp       AS ASCIIZ * 64

  FUNCTION = %FALSE
  IF RegCreateKeyEx( hKey, szSubKey, %NULL, "", %REG_OPTION_NON_VOLATILE, _
                    %KEY_WRITE, BYVAL %NULL, hOpenKey, lRet) <> %ERROR_SUCCESS THEN
    EXIT FUNCTION
  END IF

  IF lValue = %TRUE THEN
    szTmp = "Yes"
  ELSE
    szTmp = "No"
  END IF
  lRet = RegSetValueEx(hOpenKey, szValueName, %NULL, %REG_SZ, szTmp, LEN(szTmp)+1)
  CALL RegCloseKey(hOpenKey)

  IF lRet = %ERROR_SUCCESS THEN
    FUNCTION = %TRUE
  END IF
END FUNCTION

' **************************************************************************
' * GetRegValue
' * Get a string registry value
' *
' * hKey       - One of the following equates:
' *               %HKEY_CLASSES_ROOT,
' *               %HKEY_CURRENT_CONFIG
' *               %HKEY_CURRENT_USER
' *               %HKEY_LOCAL_MACHINE
' *               %HKEY_USERS
' * zSubKey    - Name of a subkey. This parameter cannot be NULL
' * zValueName - Name of the value to query
' **************************************************************************
'Function GetRegValue(ByVal hKey As Long, szSubKey As Asciiz, szValueName As Asciiz) As String
'Function GetRegValue(ByVal hKey As Long, szSubKey As Asciiz, szValueName As Asciiz) As String
Function GetRegValue( _
                                ByVal hLocation   As Dword, _
                                ByVal sSubKeys    As String, _
                                ByVal sValueName  As String, _
                      Optional ByVal sDefault     As String _
                     ) As String
'  Local lRet        As Long
'  LOCAL hOpenKey    AS LONG
'  Local sBuffer     As Asciiz * 32000
'  Local lBufferSize As Long
'  Local szKey       As Asciiz * 32000


'Function FF_GetRegistryString( ByVal hLocation As Dword, _
'                               ByVal sSubKeys As String, _
'                               ByVal sValueName As String, _
'                               ByVal sDefault As String _
'                               ) As String
  
  Local hKey    As Dword
  Local zRegVal As Asciiz * 1024
  Local dwType  As Dword
  Local dwSize  As Dword
  
  If hLocation = 0 Then hLocation = %HKEY_CURRENT_USER
  
  zRegVal = sDefault
  If (RegOpenKeyEx(hLocation, Trim$(sSubKeys, "\"), 0, %KEY_READ, hKey) = %ERROR_SUCCESS) Then
     dwType = %REG_SZ
     dwSize = SizeOf(zRegVal)
     Call RegQueryValueEx(hKey, ByCopy sValueName, 0, dwType, zRegVal, dwSize)
     RegCloseKey hKey
  End If
  Function = zRegVal
'End Function

'  
'  If Right$(szSubKey, 1) <> "\" Then szSubKey += "\"
'  szKey = szSubKey & szValueName
'  'lRet = RegOpenKeyEx(hKey, szSubKey, %Null, %KEY_QUERY_VALUE, hOpenKey)
'  'lRet = RegOpenKeyEx(hKey, szKey, %Null, %KEY_QUERY_VALUE Or %KEY_WOW64_32KEY , hOpenKey) %KEY_ALL_ACCESS
'  lRet = RegOpenKeyEx(hKey, szKey, %Null, %KEY_ALL_ACCESS Or %KEY_WOW64_64KEY , hOpenKey) 
'  
'MsgBox szKey & " --" & Str$(lRet) & Str$(getlasterror) & Str$(%ERROR_SUCCESS)
'  If lRet = %ERROR_SUCCESS Then
'MsgBox szValueName
'    'lRet = RegQueryValueEx(hOpenKey, szValueName, %NULL, %REG_SZ, BYVAL %NULL, lBufferSize)
'    lBufferSize = 32000
'    sBuffer = String$(lBufferSize, %Null)
''MsgBox Str$(Len(sBuffer))'          %REG_EXPAND_SZ
'    'lRet = RegQueryValueEx(hOpenKey, szValueName, %Null, %REG_SZ, ByVal StrPtr(sBuffer), lBufferSize)
'    'lRet = RegQueryValueEx(hOpenKey, szValueName, %Null, %REG_EXPAND_SZ, ByVal StrPtr(sBuffer), lBufferSize)
'    lRet = RegEnumValue (hOpenKey, 0, szValueName, lBufferSize, ByVal %Null, %REG_SZ, sBuffer, lBufferSize)
''MsgBox Str$(lBufferSize) & Str$(lRet)
''MsgBox RTrim$(sBuffer, Chr$(0))
'
''     If hr <> 0 Or sSubkey = "" Then Exit Function
''     Local szValueName As Asciiz * %MAX_PATH
''     Local KeyType As Dword
''     Local szKeyValue As Asciiz * %MAX_PATH
''     Local cValueName As Dword
''     Local cbData As Dword
''     ' win32 node
''     dwIdx = 0
''     cValueName = %MAX_PATH
''     cbData = %MAX_PATH
''     szKey = sSubkey & "\" & "win32"
''MsgBox szKey
''     hr = RegOpenKeyEx (%HKEY_CLASSES_ROOT, szKey, 0, %KEY_READ, hKey)
''     If hr <> %ERROR_SUCCESS Then Exit Function
''     hr = RegEnumValue (hKey, dwIdx, szValueName, cValueName, ByVal %Null, KeyType, szKeyValue, cbData)
''     RegCloseKey hKey
''     Function = szKeyValue
'
'    CALL RegCloseKey(hOpenKey)
'    FUNCTION = RTRIM$(sBuffer, CHR$(0))
'  END IF
END FUNCTION

' **************************************************************************
' * GetRegExists
' * Check if registry key exists
' *
' * hKey       - One of the following equates:
' *               %HKEY_CLASSES_ROOT,
' *               %HKEY_CURRENT_CONFIG
' *               %HKEY_CURRENT_USER
' *               %HKEY_LOCAL_MACHINE
' *               %HKEY_USERS
' * zSubKey    - Name of a subkey. This parameter cannot be NULL
' * zValueName - Name of the value to query
' **************************************************************************
Function GetRegExists(ByVal hKey As Long, szSubKey As Asciiz, szValueName As Asciiz) As Long
  Local lRet        As Long
  Local hOpenKey    As Long
  Local sBuffer     As String
  Local lBufferSize As Long

  lRet = RegOpenKeyEx(hKey, szSubKey, %Null, %KEY_QUERY_VALUE, hOpenKey)
  If lRet = %ERROR_SUCCESS Then
    lRet = RegQueryValueEx(hOpenKey, szValueName, %Null, %REG_SZ, ByVal %Null, lBufferSize)
    If lRet = %ERROR_SUCCESS Then
      'sBuffer = String$(lBufferSize, %Null)
      'lRet = RegQueryValueEx(hOpenKey, szValueName, %Null, %REG_SZ, ByVal StrPtr(sBuffer), lBufferSize)
      Function = %TRUE
    Else
      Function = -2&
    End If
    Call RegCloseKey(hOpenKey)
  Else
    Function = -1&
  End If
End Function

' **************************************************************************
' * GetRegExists
' * Check if registry key exists
' *
' * hKey       - One of the following equates:
' *               %HKEY_CLASSES_ROOT,
' *               %HKEY_CURRENT_CONFIG
' *               %HKEY_CURRENT_USER
' *               %HKEY_LOCAL_MACHINE
' *               %HKEY_USERS
' * zSubKey    - Name of a subkey. This parameter cannot be NULL
' * zValueName - Name of the value to query
' **************************************************************************
Function GetRegPathExists(ByVal hKey As Long, szSubKey As Asciiz) As Long
  Local lRet        As Long
  Local hOpenKey    As Long
  Local sBuffer     As String
  Local lBufferSize As Long

  lRet = RegOpenKeyEx(hKey, szSubKey, %Null, %KEY_QUERY_VALUE, hOpenKey)
  If lRet = %ERROR_SUCCESS Then
    Function = %TRUE
    Call RegCloseKey(hOpenKey)
  End If
End Function







' **************************************************************************
' * GetRegDwordValue
' * Get a Dword registry value
' *
' * hKey       - One of the following equates:
' *               %HKEY_CLASSES_ROOT,
' *               %HKEY_CURRENT_CONFIG
' *               %HKEY_CURRENT_USER
' *               %HKEY_LOCAL_MACHINE
' *               %HKEY_USERS
' * zSubKey    - Name of a subkey. This parameter cannot be NULL
' * zValueName - Name of the value to query
' **************************************************************************
FUNCTION GetRegDwordValue(BYVAL hKey AS LONG, szSubKey AS ASCIIZ, szValueName AS ASCIIZ) AS DWORD
  LOCAL lRet        AS LONG
  LOCAL hOpenKey    AS LONG
  LOCAL dwTmp       AS DWORD
  LOCAL lBufferSize AS LONG

  lRet = RegOpenKeyEx(hKey, szSubKey, %NULL, %KEY_QUERY_VALUE, hOpenKey)
  IF lRet = %ERROR_SUCCESS THEN
    lBufferSize = SIZEOF(dwTmp)
    lRet = RegQueryValueEx(hOpenKey, szValueName, %NULL, %REG_DWORD, dwTmp, lBufferSize)
    CALL RegCloseKey(hOpenKey)
    FUNCTION = dwTmp
  END IF

END FUNCTION

' **************************************************************************
' * GetRegBinValue
' * Get a binary registry value - very usefull to get UDTs from the registry
' *
' * hKey        - One of the following equates:
' *                %HKEY_CLASSES_ROOT,
' *                %HKEY_CURRENT_CONFIG
' *                %HKEY_CURRENT_USER
' *                %HKEY_LOCAL_MACHINE
' *                %HKEY_USERS
' * zSubKey     - Name of a subkey. This parameter cannot be NULL
' * zValueName  - Name of the value to query
' * BinValuePtr - Pointer to a buffer that receives the value's data
' * lBufferSize - Specifies the size, in bytes, of the buffer
' **************************************************************************
FUNCTION GetRegBinValue(BYVAL hKey AS LONG, szSubKey AS ASCIIZ, szValueName AS ASCIIZ, BYVAL BinValuePtr AS LONG, BYVAL lBufferSize AS LONG) AS LONG
  LOCAL lRet     AS LONG
  LOCAL hOpenKey    AS LONG

  FUNCTION = %FALSE

  lRet = RegOpenKeyEx(hKey, szSubKey, %NULL, %KEY_QUERY_VALUE, hOpenKey)
  IF lRet = %ERROR_SUCCESS THEN
    lRet = RegQueryValueEx( hOpenKey, szValueName, %NULL, %REG_BINARY, BYVAL BinValuePtr, lBufferSize)
    IF lRet = %ERROR_SUCCESS THEN
      FUNCTION = %TRUE
    END IF
    CALL RegCloseKey(hOpenKey)
  END IF
END FUNCTION

' **************************************************************************
' * GetRegTxtNumValue
' * Get a number from a registry value stored as text
' *
' * hKey        - One of the following equates:
' *                %HKEY_CLASSES_ROOT,
' *                %HKEY_CURRENT_CONFIG
' *                %HKEY_CURRENT_USER
' *                %HKEY_LOCAL_MACHINE
' *                %HKEY_USERS
' * zSubKey     - Name of a subkey. This parameter cannot be NULL
' * zValueName  - Name of the value to query
' **************************************************************************
FUNCTION GetRegTxtNumValue(BYVAL hKey AS LONG, zSubKey AS ASCIIZ, zValueName AS ASCIIZ, BYVAL lDefVal AS LONG) AS LONG
  LOCAL lRet     AS LONG
  LOCAL hOpenKey    AS LONG
  LOCAL szTmp       AS ASCIIZ * 64
  LOCAL lBufferSize AS LONG

  lRet = RegOpenKeyEx(hKey, zSubKey, %NULL, %KEY_QUERY_VALUE, hOpenKey)
  IF lRet = %ERROR_SUCCESS THEN
    lBufferSize = SIZEOF(szTmp)
    lRet = RegQueryValueEx(hOpenKey, zValueName, %NULL, %REG_SZ, szTmp, lBufferSize)
    IF LEN(szTmp) > 0 THEN
      FUNCTION = VAL(szTmp)
    ELSE
      FUNCTION = lDefVal
    END IF
    CALL RegCloseKey(hOpenKey)
  ELSE
    FUNCTION = lDefVal
  END IF

END FUNCTION

' **************************************************************************
' * GetRegTxtBoolValue
' * Get a number from a registry value stored as text
' *
' * hKey        - One of the following equates:
' *                %HKEY_CLASSES_ROOT,
' *                %HKEY_CURRENT_CONFIG
' *                %HKEY_CURRENT_USER
' *                %HKEY_LOCAL_MACHINE
' *                %HKEY_USERS
' * zSubKey     - Name of a subkey. This parameter cannot be NULL
' * zValueName  - Name of the value to query
' **************************************************************************
FUNCTION GetRegTxtBoolValue(BYVAL hKey AS LONG, zSubKey AS ASCIIZ, zValueName AS ASCIIZ, BYVAL lDefVal AS LONG) AS LONG
  LOCAL RetCode     AS LONG
  LOCAL hOpenKey    AS LONG
  LOCAL zText       AS ASCIIZ * 2048
  LOCAL lBufferSize AS LONG

  RetCode = RegOpenKeyEx(hKey, zSubKey, %NULL, %KEY_QUERY_VALUE, hOpenKey)
  IF RetCode = %ERROR_SUCCESS THEN
    lBufferSize = SIZEOF(zText)
    RetCode = RegQueryValueEx( hOpenKey, zValueName, %NULL, %REG_SZ, zText, lBufferSize)
    IF LEN(zText) > 0 THEN
      IF UCASE$(zText) = "YES" OR UCASE$(zText) = "TRUE" THEN
        FUNCTION = %TRUE
      END IF
    ELSE
      FUNCTION = lDefVal
    END IF
  ELSE
    FUNCTION = lDefVal
  END IF

  RegCloseKey hOpenKey

END FUNCTION


' **************************************************************************
' * DelRegValue
' * Delete a registry value
' *
' * hKey        - One of the following equates:
' *                %HKEY_CLASSES_ROOT,
' *                %HKEY_CURRENT_CONFIG
' *                %HKEY_CURRENT_USER
' *                %HKEY_LOCAL_MACHINE
' *                %HKEY_USERS
' * zSubKey     - Name of a subkey. This parameter cannot be NULL
' * zValueName  - Name of the value to remove
' **************************************************************************
FUNCTION DelRegValue (BYVAL hKey AS LONG, zSubKey AS ASCIIZ, zValueName AS ASCIIZ) AS LONG
  LOCAL RetCode     AS LONG
  LOCAL hOpenKey    AS LONG

  FUNCTION = %FALSE

  RetCode = RegOpenKeyEx(hKey, zSubKey, %NULL, %KEY_ALL_ACCESS, hOpenKey)
  IF RetCode = %ERROR_SUCCESS THEN
    RetCode = RegDeleteValue(hOpenKey, zValueName)
  END IF

  RegCloseKey hOpenKey

  IF RetCode = %ERROR_SUCCESS THEN
    FUNCTION = %TRUE
  END IF
END FUNCTION

' **************************************************************************
' * DelRegValue
' * Delete a registry key
' *
' * hKey        - One of the following equates:
' *                %HKEY_CLASSES_ROOT,
' *                %HKEY_CURRENT_CONFIG
' *                %HKEY_CURRENT_USER
' *                %HKEY_LOCAL_MACHINE
' *                %HKEY_USERS
' * zSubKey     - Name of a subkey to remove. This parameter cannot be NULL
' **************************************************************************
FUNCTION DelRegKey(BYVAL hKey AS LONG, zSubKey AS ASCIIZ) AS LONG
  LOCAL RetCode   AS LONG
  LOCAL hOpenKey  AS LONG

  FUNCTION = %FALSE

  IF RegOpenKeyEx(hKey, zSubKey, %NULL, %KEY_ALL_ACCESS, hOpenKey) <> %ERROR_SUCCESS THEN  EXIT FUNCTION
  RetCode = RegDeleteKey(hKey, zSubKey)
  RegCloseKey hOpenKey

  IF RetCode = %ERROR_SUCCESS THEN
    FUNCTION = %TRUE
  END IF
END FUNCTION
'------------------------------------------------------------------------------

