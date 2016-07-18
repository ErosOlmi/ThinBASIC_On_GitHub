%KEY_WOW64_32KEY = &H0200???
%KEY_WOW64_64KEY = &H0100???

FUNCTION GetRegValues(hKey AS DWORD) AS STRING
  DIM lReturn            AS LONG
  DIM azKeyName          AS ASCIIZ * 32767   ' up to 32 KB according to MSDN
  DIM KeyType            AS DWORD
  DIM azKeyValue         AS ASCIIZ * 1048576 ' up to 1 MB according to MSDN
  DIM Reserved           AS DWORD
  DIM lIndex             AS DWORD
  DIM azKeyNameLen       AS DWORD
  DIM azKeyValueLen      AS DWORD

  DIM sKeyValue          AS STRING
  DIM items(1 TO 1024)   AS STRING
  DIM itemCount          AS LONG

  DO WHILE lReturn <> %ERROR_NO_MORE_ITEMS
    azKeyNameLen  = SIZEOF(azKeyName)
    azKeyValueLen = SIZEOF(azKeyValue)

    Reserved = %NULL
    lReturn  = RegEnumValue(hKey, lIndex, azKeyName, BYVAL VARPTR(azKeyNameLen), BYVAL Reserved , BYVAL VARPTR(KeyType), azKeyValue, BYVAL VARPTR(azKeyValueLen))
    INCR lIndex

    IF lReturn = %ERROR_SUCCESS THEN
      SELECT CASE keyType
        CASE %REG_DWORD
          sKeyValue = FORMAT$(CVDWD(azKeyValue))

        CASE %REG_QWORD
          sKeyValue = FORMAT$(CVQ(azKeyValue))

        CASE ELSE
          sKeyValue = LEFT$(azKeyValue, azKeyValueLen)
      END SELECT

      INCR itemCount
      IF itemCount > UBOUND(items) THEN
        REDIM PRESERVE items(1 TO UBOUND(items) * 2)
      END IF
      items(itemCount) = TRIM$(azKeyName) + "=" + sKeyValue
    ELSEIF lReturn <> %ERROR_NO_MORE_ITEMS THEN
        FUNCTION = "Error code " + FORMAT$(lReturn) + " in RegEnumValue"
        EXIT FUNCTION
    END IF
  LOOP

  IF itemCount THEN
    REDIM PRESERVE items(1 TO itemCount)
    ARRAY SORT items()

    FUNCTION = JOIN$(items(), $CRLF)
  END IF
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

  strTmp = GetRegValues(hKey)

  ' Enumerate all the subkeys for the current key.
  '...............................................
  DO WHILE lReturn <> %ERROR_NO_MORE_ITEMS
     IF lReturn = %ERROR_SUCCESS THEN
        lNameLen = SIZEOF(szSubKeyName)
        lClassLen = SIZEOF(azClass)
        lReturn = RegEnumKeyEx(hKey, lIndex, szSubKeyName, BYVAL VARPTR(lNameLen), Reserved, BYVAL VARPTR(azClass), BYVAL VARPTR(lClasslen), LastWriteTime)

        IF (lReturn <> %ERROR_NO_MORE_ITEMS) THEN
           IF lReturn = %ERROR_SUCCESS THEN
              strTmp += IIF$(LEN(strTmp), sSep, "") + szSubKeyName
           END IF
        END IF
     END IF
     INCR lIndex
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
