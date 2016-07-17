#IF 0
  =============================================================================
   Program NAME: BINT_MyLib.bas
   Author      : Eros Olmi
   Date        : 08/02/2004
   Version     : 
   Description : DLL to test BINT loading library
  =============================================================================
  'COPYRIGHT AND PERMISSION NOTICE
  '============================================================================
  'Copyright (c) 2003 - 2004, Eros Olmi, <eros.olmi@autoapfp.com>
  '   
  'ALL rights reserved.
  '   
  'Permission TO use AND copy this software FOR ANY non commercial purpose
  'WITH OR without fee is hereby granted, provided that the above copyright
  'notice AND this permission notice appear IN ALL copies.
  '
  'FOR commercial purpose, contact the copyright holder.
  ' 
  'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT OF THIRD PARTY RIGHTS.
  'IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
  'DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
  'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE
  'USE OR OTHER DEALINGS IN THE SOFTWARE.
  '----------------------------------------------------------------------------
  'Note: BINT32.DLL is dependant FROM MDICT.DLL
  '      MDICT.DLL is copyrighted by:
  '                Copyright © Radbit GmbH, Florent Heyworth, 2000
  =============================================================================
#ENDIF

  #COMPILE DLL
  #REGISTER NONE
  #DIM ALL

  #RESOURCE "thinBasic_Registry.PBR"
  
  #INCLUDE "WIN32API.INC"
  #INCLUDE "..\thinCore.INC"

  '------------------------------------------------------------------
  #INCLUDE ".\thinRegistry.INC"
  '------------------------------------------------------------------


  Function Registry_ConvertHKey(ByVal sHKey As String) As Dword
    SELECT CASE UCASE$(TRIM$(sHKey))
      CASE "HKEYCR"
        FUNCTION = %HKEY_CLASSES_ROOT
      CASE "HKEYCU"
        FUNCTION = %HKEY_CURRENT_USER
      CASE "HKEYLM"
        FUNCTION = %HKEY_LOCAL_MACHINE
      CASE "HKEYU"
        FUNCTION = %HKEY_USERS
      CASE "HKEYCC"
        FUNCTION = %HKEY_CURRENT_CONFIG
    END SELECT
  END FUNCTION
                                               
  '------------------------------------------------------------------------------
  FUNCTION Exec_Registry_GetValue() AS STRING
  '------------------------------------------------------------------------------
  'Syntax: String = REGISTRY_GetValue(HKey, MainKey, ValueName)
  '------------------------------------------------------------------------------
    LOCAL lHKey       AS STRING
    LOCAL lMainKey    AS STRING
    LOCAL lValueName  AS STRING
    LOCAL lTmp        AS STRING
    LOCAL tmpHKey     AS DWORD
  
    IF thinBasic_CheckOpenParens() THEN
      thinBasic_ParseString lHKey
      IF thinBasic_CheckComma() THEN
        thinBasic_ParseString lMainKey
        IF thinBasic_CheckComma() THEN
          thinBasic_ParseString lValueName
          IF thinBasic_CheckCloseParens() THEN 
            tmpHKey = Registry_ConvertHKey(lHKey)
            If tmpHKey <> 0 Then
'MsgBox FuncName$ 
              lTmp = GetRegValue(tmpHKey, BYCOPY lMainKey, BYCOPY lValueName)
              FUNCTION = lTmp
            END IF
          END IF
        END IF
      END IF
    END IF
  END FUNCTION


  '------------------------------------------------------------------------------
  Function Exec_Registry_KeyExists() As Ext
  '------------------------------------------------------------------------------
  'Syntax: String = REGISTRY_KeyExists(HKey, MainKey, ValueName)
  '------------------------------------------------------------------------------
    Local lHKey       As String
    Local lMainKey    As String
    Local lValueName  As String
    Local lret        As Long
    Local tmpHKey     As Dword
  
    If thinBasic_CheckOpenParens() Then
      thinBasic_ParseString lHKey
      If thinBasic_CheckComma() Then
        thinBasic_ParseString lMainKey
        If thinBasic_CheckComma() Then
          thinBasic_ParseString lValueName
          If thinBasic_CheckCloseParens() Then 
            tmpHKey = Registry_ConvertHKey(lHKey)
            If tmpHKey <> 0 Then 
              lret = GetRegExists(tmpHKey, ByCopy lMainKey, ByCopy lValueName)
              Function = lret
            End If
          End If
        End If
      End If
    End If
  End Function

  '------------------------------------------------------------------------------
  Function Exec_Registry_PathExists() As Ext
  '------------------------------------------------------------------------------
  'Syntax: String = REGISTRY_PathExists(HKey, MainKey)
  '------------------------------------------------------------------------------
    Local lHKey       As String
    Local lMainKey    As String
    Local lret        As Long
    Local tmpHKey     As Dword
  
    If thinBasic_CheckOpenParens() Then
      thinBasic_ParseString lHKey
      If thinBasic_CheckComma() Then
        thinBasic_ParseString lMainKey
        If thinBasic_CheckCloseParens() Then 
          tmpHKey = Registry_ConvertHKey(lHKey)
          If tmpHKey <> 0 Then 
            lret = GetRegPathExists(tmpHKey, ByCopy lMainKey)
            Function = lret
          End If
        End If
      End If
    End If
  End Function

  '------------------------------------------------------------------------------
  FUNCTION Exec_Registry_GetDWord() AS EXT
  '------------------------------------------------------------------------------
  'Syntax: String = REGISTRY_GetDWord(HKey, MainKey, ValueName)
  '------------------------------------------------------------------------------
    LOCAL lHKey       AS STRING
    LOCAL lMainKey    AS STRING
    LOCAL lValueName  AS STRING
    LOCAL lTmp        AS DWORD
    LOCAL tmpHKey     AS DWORD
  
    IF thinBasic_CheckOpenParens() THEN
      thinBasic_ParseString lHKey
      IF thinBasic_CheckComma() THEN
        thinBasic_ParseString lMainKey
        IF thinBasic_CheckComma() THEN
          thinBasic_ParseString lValueName
          IF thinBasic_CheckCloseParens() THEN 
            tmpHKey = Registry_ConvertHKey(lHKey)
            IF tmpHKey <> 0 THEN 
              lTmp = GetRegDwordValue(tmpHKey, BYCOPY lMainKey, BYCOPY lValueName)
              FUNCTION = lTmp
            END IF
          END IF
        END IF
      END IF
    END IF
  END FUNCTION

  '------------------------------------------------------------------------------
  FUNCTION Exec_Registry_GetTxtNum() AS EXT
  '------------------------------------------------------------------------------
  'Syntax: String = REGISTRY_GetTxtNum(HKey, MainKey, ValueName, DefaultValue)
  '------------------------------------------------------------------------------
    LOCAL lHKey         AS STRING
    LOCAL lMainKey      AS STRING
    LOCAL lValueName    AS STRING
    LOCAL lDefaultValue AS EXT
    LOCAL lTmp          AS DWORD
    LOCAL tmpHKey       AS DWORD
  
    IF thinBasic_CheckOpenParens() THEN
      thinBasic_ParseString lHKey
      IF thinBasic_CheckComma() THEN
        thinBasic_ParseString lMainKey
        IF thinBasic_CheckComma() THEN
          thinBasic_ParseString lValueName
          IF thinBasic_CheckComma() THEN
            thinBasic_ParseNumber lDefaultValue           
            IF thinBasic_CheckCloseParens() THEN 
              tmpHKey = Registry_ConvertHKey(lHKey)
              IF tmpHKey <> 0 THEN 
                lTmp = GetRegTxtNumValue(tmpHKey, BYCOPY lMainKey, BYCOPY lValueName, BYCOPY lDefaultValue)
                FUNCTION = lTmp
              END IF
            END IF
          END IF
        END IF
      END IF
    END IF
  END FUNCTION

  '------------------------------------------------------------------------------
  FUNCTION Exec_Registry_GetTxtBool() AS EXT
  '------------------------------------------------------------------------------
  'Syntax: String = REGISTRY_GetTxtBool(HKey, MainKey, ValueName, DefaultValue)
  '------------------------------------------------------------------------------
    LOCAL lHKey         AS STRING
    LOCAL lMainKey      AS STRING
    LOCAL lValueName    AS STRING
    LOCAL lDefaultValue AS EXT
    LOCAL lTmp          AS DWORD
    LOCAL tmpHKey       AS DWORD
  
    IF thinBasic_CheckOpenParens() THEN
      thinBasic_ParseString lHKey
      IF thinBasic_CheckComma() THEN
        thinBasic_ParseString lMainKey
        IF thinBasic_CheckComma() THEN
          thinBasic_ParseString lValueName
          IF thinBasic_CheckComma() THEN
            thinBasic_ParseNumber lDefaultValue           
            IF thinBasic_CheckCloseParens() THEN 
              tmpHKey = Registry_ConvertHKey(lHKey)
              IF tmpHKey <> 0 THEN 
                lTmp = GetRegTxtBoolValue(tmpHKey, BYCOPY lMainKey, BYCOPY lValueName, BYCOPY lDefaultValue)
                FUNCTION = lTmp
              END IF
            END IF
          END IF
        END IF
      END IF
    END IF
  END FUNCTION

  '------------------------------------------------------------------------------
  FUNCTION Exec_Registry_SetValue() AS EXT
  '------------------------------------------------------------------------------
  'Syntax: Number = REGISTRY_SetValue(HKey, MainKey, Key, Value)
  '------------------------------------------------------------------------------
    LOCAL lHKey     AS STRING
    LOCAL lMainKey  AS STRING
    LOCAL lKey      AS STRING
    LOCAL lTmp      AS EXT
    LOCAL lValue    AS STRING
    LOCAL tmpHKey   AS DWORD
  
    IF thinBasic_CheckOpenParens() THEN
      thinBasic_ParseString lHKey
      IF thinBasic_CheckComma() THEN
        thinBasic_ParseString lMainKey
        IF thinBasic_CheckComma() THEN
          thinBasic_ParseString lKey
          IF thinBasic_CheckComma() THEN
            thinBasic_ParseString lValue
            IF thinBasic_CheckCloseParens() THEN 
              tmpHKey = Registry_ConvertHKey(lHKey)
              IF tmpHKey <> 0 THEN 
                lTmp = SetRegValue(tmpHKey, BYCOPY lMainKey, BYCOPY lKey, BYCOPY lValue)
                FUNCTION = lTmp
              END IF
            END IF
          END IF
        END IF
      END IF
    END IF
  END FUNCTION

  '------------------------------------------------------------------------------
  FUNCTION Exec_Registry_SetDWord() AS EXT
  '------------------------------------------------------------------------------
  'Syntax: Number = REGISTRY_SetValue(HKey, MainKey, Key, Value)
  '------------------------------------------------------------------------------
    LOCAL lHKey     AS STRING
    LOCAL lMainKey  AS STRING
    LOCAL lKey      AS STRING
    LOCAL lTmp      AS EXT
    LOCAL lValue    AS EXT
    LOCAL tmpHKey   AS DWORD
  
    IF thinBasic_CheckOpenParens() THEN
      thinBasic_ParseString lHKey
      IF thinBasic_CheckComma() THEN
        thinBasic_ParseString lMainKey
        IF thinBasic_CheckComma() THEN
          thinBasic_ParseString lKey
          IF thinBasic_CheckComma() THEN
            thinBasic_ParseNumber lValue
            IF thinBasic_CheckCloseParens() THEN 
              tmpHKey = Registry_ConvertHKey(lHKey)
              IF tmpHKey <> 0 THEN 
                lTmp = SetRegDwordValue(tmpHKey, BYCOPY lMainKey, BYCOPY lKey, BYCOPY lValue)
                FUNCTION = lTmp
              END IF
            END IF
          END IF
        END IF
      END IF
    END IF
  END FUNCTION

  '------------------------------------------------------------------------------
  FUNCTION Exec_Registry_SetTxtNum() AS EXT
  '------------------------------------------------------------------------------
  'Syntax: Number = REGISTRY_SetTxtNum(HKey, MainKey, Key, Value)
  '------------------------------------------------------------------------------
    LOCAL lHKey     AS STRING
    LOCAL lMainKey  AS STRING
    LOCAL lKey      AS STRING
    LOCAL lTmp      AS EXT
    LOCAL lValue    AS EXT
    LOCAL tmpHKey   AS DWORD
  
    IF thinBasic_CheckOpenParens() THEN
      thinBasic_ParseString lHKey
      IF thinBasic_CheckComma() THEN
        thinBasic_ParseString lMainKey
        IF thinBasic_CheckComma() THEN
          thinBasic_ParseString lKey
          IF thinBasic_CheckComma() THEN
            thinBasic_ParseNumber lValue
            IF thinBasic_CheckCloseParens() THEN 
              tmpHKey = Registry_ConvertHKey(lHKey)
              IF tmpHKey <> 0 THEN 
                lTmp = SetRegTxtNumValue(tmpHKey, BYCOPY lMainKey, BYCOPY lKey, BYCOPY lValue)
                FUNCTION = lTmp
              END IF
            END IF
          END IF
        END IF
      END IF
    END IF
  END FUNCTION

  '------------------------------------------------------------------------------
  FUNCTION Exec_Registry_SetTxtBool() AS EXT
  '------------------------------------------------------------------------------
  'Syntax: Number = REGISTRY_SetTxtBool(HKey, MainKey, Key, Value)
  '------------------------------------------------------------------------------
    LOCAL lHKey     AS STRING
    LOCAL lMainKey  AS STRING
    LOCAL lKey      AS STRING
    LOCAL lTmp      AS EXT
    LOCAL lValue    AS EXT
    LOCAL tmpHKey   AS DWORD
  
    IF thinBasic_CheckOpenParens() THEN
      thinBasic_ParseString lHKey
      IF thinBasic_CheckComma() THEN
        thinBasic_ParseString lMainKey
        IF thinBasic_CheckComma() THEN
          thinBasic_ParseString lKey
          IF thinBasic_CheckComma() THEN
            thinBasic_ParseNumber lValue
            IF thinBasic_CheckCloseParens() THEN 
              tmpHKey = Registry_ConvertHKey(lHKey)
              IF tmpHKey <> 0 THEN 
                lTmp = SetRegTxtBoolValue(tmpHKey, BYCOPY lMainKey, BYCOPY lKey, BYCOPY lValue)
                FUNCTION = lTmp
              END IF
            END IF
          END IF
        END IF
      END IF
    END IF
  END FUNCTION

  '------------------------------------------------------------------------------
  FUNCTION Exec_Registry_GetAllKeys() AS STRING
  '------------------------------------------------------------------------------
  'Syntax: List = REGISTRY_GetAllKeys(HKey, MainKey, Separator)
  '------------------------------------------------------------------------------
    LOCAL sHKey       AS STRING
    LOCAL sMainKey    AS STRING
    LOCAL sSeparator  AS STRING

    LOCAL azMainKey AS ASCIIZ * 1024
    LOCAL dwHKey    AS DWORD

    LOCAL parensPresent AS BYTE

    parensPresent = thinBasic_CheckOpenParens_Optional

    thinBasic_ParseString sHKey
    IF thinBasic_CheckComma_Mandatory THEN
      thinBasic_ParseString sMainKey
      IF LEFT$(sMainKey, 1) = "\" THEN sMainKey = LTRIM$(sMainKey, "\")

      IF thinBasic_CheckComma_Optional THEN
        thinBasic_ParseString sSeparator
      ELSE
        sSeparator = $CRLF
      END IF

      dwHKey = Registry_ConvertHKey(sHKey)
      IF thinBasic_ErrorFree() THEN
        azMainKey = sMainKey
        FUNCTION = GetAllKeys(dwHKey, azMainKey, sSeparator)
      END IF
    END IF

    IF parensPresent THEN thinBasic_CheckCloseParens_Mandatory

  END FUNCTION

  '------------------------------------------------------------------------------
  FUNCTION Exec_Registry_DelValue() AS EXT
  '------------------------------------------------------------------------------
  'Syntax: Number = REGISTRY_DelValue(HKey, Key, ValueName)
  '------------------------------------------------------------------------------
    LOCAL lHKey       AS STRING
    LOCAL lKey        AS STRING
    LOCAL lTmp        AS EXT
    LOCAL lValueName  AS STRING
    LOCAL tmpHKey     AS DWORD
  
    IF thinBasic_CheckOpenParens() THEN
      thinBasic_ParseString lHKey
      IF thinBasic_CheckComma() THEN
        thinBasic_ParseString lKey
        IF thinBasic_CheckComma() THEN
          thinBasic_ParseString lValueName
          IF thinBasic_CheckCloseParens() THEN 
            tmpHKey = Registry_ConvertHKey(lHKey)
            IF tmpHKey <> 0 THEN 
              lTmp = DelRegValue(tmpHKey, BYCOPY lKey, BYCOPY lValueName)
              FUNCTION = lTmp
            END IF
          END IF
        END IF
      END IF
    END IF
  END FUNCTION

  '------------------------------------------------------------------------------
  FUNCTION Exec_Registry_DelKey() AS EXT
  '------------------------------------------------------------------------------
  'Syntax: Number = REGISTRY_DelKey(HKey, SubKey)
  '------------------------------------------------------------------------------
    LOCAL lHKey       AS STRING
    LOCAL lSubKey     AS STRING
    LOCAL lTmp        AS EXT
    LOCAL lValueName  AS STRING
    LOCAL tmpHKey     AS DWORD
  
    IF thinBasic_CheckOpenParens() THEN
      thinBasic_ParseString lHKey
      IF thinBasic_CheckComma() THEN
        thinBasic_ParseString lSubKey
        IF thinBasic_CheckCloseParens() THEN 
          tmpHKey = Registry_ConvertHKey(lHKey)
          IF tmpHKey <> 0 THEN 
            lTmp = DelRegKey(tmpHKey, BYCOPY lSubKey)
            FUNCTION = lTmp
          END IF
        END IF
      END IF
    END IF
  END FUNCTION

 
 
  '----------------------------------------------------------------------------
  FUNCTION LoadLocalSymbols ALIAS "LoadLocalSymbols" (OPTIONAL BYVAL sPath AS STRING) EXPORT AS LONG
  ' This function is automatically called by thinCore whenever this DLL is loaded.
  ' This function MUST be present in every external DLL you want to use
  ' with thinBasic
  ' Use this function to initialize every variable you need and for loading the
  ' new symbol (read Keyword) you have created.
  '----------------------------------------------------------------------------

    thinBasic_LoadSymbol "Registry_GetAllKeys"  , %thinBasic_ReturnString , CODEPTR(Exec_Registry_GetAllKeys    ),  %thinBasic_ForceOverWrite
    thinBasic_LoadSymbol "Registry_GetValue"    , %thinBasic_ReturnString , CODEPTR(Exec_Registry_GetValue      ),  %thinBasic_ForceOverWrite
    thinBasic_LoadSymbol "Registry_GetDWord"    , %thinBasic_ReturnNumber , CODEPTR(Exec_Registry_GetDWord      ),  %thinBasic_ForceOverWrite
    thinBasic_LoadSymbol "Registry_GetTxtNum"   , %thinBasic_ReturnNumber , CODEPTR(Exec_Registry_GetTxtNum     ),  %thinBasic_ForceOverWrite
    thinBasic_LoadSymbol "Registry_GetTxtBool"  , %thinBasic_ReturnNumber , CODEPTR(Exec_Registry_GetTxtBool    ),  %thinBasic_ForceOverWrite

    thinBasic_LoadSymbol "Registry_SetValue"    , %thinBasic_ReturnNumber , CODEPTR(Exec_Registry_SetValue      ),  %thinBasic_ForceOverWrite
    thinBasic_LoadSymbol "Registry_SetDWord"    , %thinBasic_ReturnNumber , CODEPTR(Exec_Registry_SetDWord      ),  %thinBasic_ForceOverWrite
    thinBasic_LoadSymbol "Registry_SetTxtNum"   , %thinBasic_ReturnNumber , CODEPTR(Exec_Registry_SetTxtNum     ),  %thinBasic_ForceOverWrite
    thinBasic_LoadSymbol "Registry_SetTxtBool"  , %thinBasic_ReturnNumber , CODEPTR(Exec_Registry_SetTxtBool    ),  %thinBasic_ForceOverWrite

    thinBasic_LoadSymbol "Registry_DelValue"    , %thinBasic_ReturnNumber , CODEPTR(Exec_Registry_DelValue      ),  %thinBasic_ForceOverWrite
    thinBasic_LoadSymbol "Registry_DelKey"      , %thinBasic_ReturnNumber , CODEPTR(Exec_Registry_DelKey        ),  %thinBasic_ForceOverWrite

    thinBasic_LoadSymbol "Registry_KeyExists"   , %thinBasic_ReturnNumber , CodePtr(Exec_Registry_KeyExists     ),  %thinBasic_ForceOverWrite
    thinBasic_LoadSymbol "Registry_PathExists"  , %thinBasic_ReturnNumber , CodePtr(Exec_Registry_PathExists    ),  %thinBasic_ForceOverWrite

    FUNCTION = 0&
  END FUNCTION

  '----------------------------------------------------------------------------
  FUNCTION UnLoadLocalSymbols ALIAS "UnLoadLocalSymbols" () EXPORT AS LONG
  ' This function is automatically called by thinCore whenever this DLL is unloaded.
  ' This function CAN be present but it is not necessary.
  ' Use this function to perform uninitialize process, if needed.
  '----------------------------------------------------------------------------


    FUNCTION = 0&
  END FUNCTION


  FUNCTION LIBMAIN ALIAS "LibMain" (BYVAL hInstance   AS LONG, _
                                    BYVAL fwdReason   AS LONG, _
                                    BYVAL lpvReserved AS LONG) EXPORT AS LONG
    SELECT CASE fwdReason
      CASE %DLL_PROCESS_ATTACH

        FUNCTION = 1
        EXIT FUNCTION
      CASE %DLL_PROCESS_DETACH

        FUNCTION = 1
        EXIT FUNCTION
      CASE %DLL_THREAD_ATTACH
      
        FUNCTION = 1
        EXIT FUNCTION
      CASE %DLL_THREAD_DETACH
      
        FUNCTION = 1
        EXIT FUNCTION
    END SELECT 
    
  END FUNCTION
