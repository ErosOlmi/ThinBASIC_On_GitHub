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
  DIM COUNT              AS INTEGER
  DIM I                  AS INTEGER

  lIndex   = 0
  Textout3 = ""
  lReturn  = 0
  COUNT    = 0

  DO WHILE lReturn <> 259 '%ERROR_NO_MORE_ITEMS  '(259)
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
            COUNT = COUNT + 1
            OutText(COUNT) = KeyName + CHR$(255) + A
        ELSE
        END IF
     END IF
     INCR lIndex
  LOOP
  IF COUNT > 0 THEN
     ARRAY SORT OutText(1) FOR COUNT, FROM 1 TO 15
     FOR I = 1 TO COUNT
         KeyName = EXTRACT$(OutText(I), CHR$(255))
         A = REMAIN$(OutText(I), CHR$(255))
         TextOut3 = TextOut3 + KeyName + "=" + A + $CRLF
     NEXT I
  END IF
  FUNCTION = TextOut3
END FUNCTION

FUNCTION GetAllKeys(BYVAL lHKey AS DWORD, lHive AS ASCIIZ, BYVAL sSep AS STRING) AS STRING
  DIM lReturn            AS LONG
  DIM szSubKeyName       AS ASCIIZ * 1024
  DIM azClass            AS ASCIIZ * 1024
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

  ' Enumerate all the subkeys for the current key.
  '...............................................
  DO WHILE  lReturn <> %ERROR_NO_MORE_ITEMS
     IF lReturn = %ERROR_SUCCESS THEN
        lNameLen = SIZEOF(szSubKeyName)
        lClassLen = SIZEOF(azClass)
        lReturn = RegEnumKeyEx(hKey, lIndex, szSubKeyName, BYVAL VARPTR(lNameLen), Reserved, BYVAL VARPTR(azClass), BYVAL VARPTR(lClasslen), LastWriteTime)

        IF (lReturn <> %ERROR_NO_MORE_ITEMS) THEN
           IF lReturn = %ERROR_SUCCESS THEN
              strTmp += IIF$(LEN(strTmp), sSep, "") + szSubKeyName '+ sSep '$LIST_SEP
           END IF
        END IF
     END IF
     lIndex = lIndex + 1
  LOOP

  RegCloseKey hKey

  FUNCTION = strTmp

END FUNCTION

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

  szTmp = IIF$(lValue = %TRUE, "Yes", "No")
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
FUNCTION GetRegValue( _
                                BYVAL hLocation   AS DWORD, _
                                BYVAL sSubKeys    AS STRING, _
                                BYVAL sValueName  AS STRING, _
                      OPTIONAL BYVAL sDefault     AS STRING _
                     ) AS STRING

  LOCAL hKey    AS DWORD
  LOCAL zRegVal AS ASCIIZ * 1024
  LOCAL dwType  AS DWORD
  LOCAL dwSize  AS DWORD

  IF hLocation = 0 THEN hLocation = %HKEY_CURRENT_USER

  zRegVal = sDefault
  IF (RegOpenKeyEx(hLocation, TRIM$(sSubKeys, "\"), 0, %KEY_READ, hKey) = %ERROR_SUCCESS) THEN
     dwType = %REG_SZ
     dwSize = SIZEOF(zRegVal)
     CALL RegQueryValueEx(hKey, BYCOPY sValueName, 0, dwType, zRegVal, dwSize)
     RegCloseKey hKey
  END IF
  FUNCTION = zRegVal

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
FUNCTION GetRegExists(BYVAL hKey AS LONG, szSubKey AS ASCIIZ, szValueName AS ASCIIZ) AS LONG
  LOCAL lRet        AS LONG
  LOCAL hOpenKey    AS LONG
  LOCAL sBuffer     AS STRING
  LOCAL lBufferSize AS LONG

  lRet = RegOpenKeyEx(hKey, szSubKey, %Null, %KEY_QUERY_VALUE, hOpenKey)
  IF lRet = %ERROR_SUCCESS THEN
    lRet = RegQueryValueEx(hOpenKey, szValueName, %Null, %REG_SZ, BYVAL %Null, lBufferSize)
    IF lRet = %ERROR_SUCCESS THEN
      'sBuffer = String$(lBufferSize, %Null)
      'lRet = RegQueryValueEx(hOpenKey, szValueName, %Null, %REG_SZ, ByVal StrPtr(sBuffer), lBufferSize)
      FUNCTION = %TRUE
    ELSE
      FUNCTION = -2&
    END IF
    CALL RegCloseKey(hOpenKey)
  ELSE
    FUNCTION = -1&
  END IF
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
FUNCTION GetRegPathExists(BYVAL hKey AS LONG, szSubKey AS ASCIIZ) AS LONG
  LOCAL lRet        AS LONG
  LOCAL hOpenKey    AS LONG
  LOCAL sBuffer     AS STRING
  LOCAL lBufferSize AS LONG

  lRet = RegOpenKeyEx(hKey, szSubKey, %Null, %KEY_QUERY_VALUE, hOpenKey)
  IF lRet = %ERROR_SUCCESS THEN
    FUNCTION = %TRUE
    CALL RegCloseKey(hOpenKey)
  END IF
END FUNCTION

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
