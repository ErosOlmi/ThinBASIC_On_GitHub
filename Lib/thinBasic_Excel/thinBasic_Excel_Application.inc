''
'' Excel_Application class wrapper
''
'
'' -----------------------------------------------------------------------------
'' <Class>_Register functions handle the related module initialization
'' -----------------------------------------------------------------------------
'SUB Excel_Application_Register()
'
'  LOCAL pExcel_Application, retCode AS LONG
'
'  ' -- We register class and its function
'  pExcel_Application = thinBasic_Class_Add("Excel_Application", CODEPTR(Excel_Application_ClassFunction))
'
'  ' -- In case of success a constructor is registered, the other stuff is handled in class function
'  IF pExcel_Application THEN
'
'    ' -- Constructor wrapper function needs to be linked in as _Create
'    retCode = thinBasic_Class_AddMethod   (pExcel_Application, "_Create"      , %thinBasic_ReturnNone       , CODEPTR(Excel_Application_Create))
'
'  END IF
'
'END SUB
'
'' -----------------------------------------------------------------------------
'' Constructor
'' -----------------------------------------------------------------------------
'SUB Excel_Application_Create( BYVAL pObject AS LONG )
'
'  REDIM oObject(1& TO 1&) AS Excel_Application AT pObject
'
'  oObject(1&) = ANYCOM $PROGID_Excel_Application
'
'  IF ISNOTHING(oObject(1&)) THEN
'    thinbasic_RunTimeError ( %ERR__MODULE_SPECIFIC, "Failed to instantiate Excel_Application" )
'  END IF
'
'END SUB
'
'' -----------------------------------------------------------------------------
'' Class function to manage the method calling, even in chained way
'' -----------------------------------------------------------------------------
'FUNCTION Excel_Application_ClassFunction( _
'                                                        BYVAL pObject     AS LONG       , _
'                                                        BYVAL lAction     AS LONG       , _
'                                              OPTIONAL  BYVAL RetValNum   AS EXT PTR    , _
'                                                        BYVAL RetValStr   AS STRING PTR , _
'                                                        BYVAL lRecursive  AS LONG         _
'                                            ) AS LONG
'
'    LOCAL sName       AS STRING
'    LOCAL lEqualType  AS LONG
'    LOCAL nChars      AS LONG
'    LOCAL eStart      AS EXT
'    LOCAL eLen        AS EXT
'    LOCAL MatchString AS STRING
'    LOCAL NewString   AS STRING
'
'    ' -- Mapping the object
'    REDIM oObject(1& TO 1&) AS Excel_Application AT pObject
'
'    sName = thinBasic_GetUnknownToken
'
'    ' -- Processing the token, UCASE way
'    SELECT CASE CONST$ sName
'
'      '---------------
'      ' -- Methods
'      '---------------
'      CASE "WORKBOOKS", "ACTIVEWINDOW"
'        FUNCTION = %thinBasic_ReturnString
'
'        ' -- If all went fine ...
'        IF thinBasic_ErrorFree THEN
'          IF lRecursive THEN
'            thinbasic_RunTimeError ( %ERR__MODULE_SPECIFIC, "You cannot use "+sName+" at this level" )
'          ELSE
'            @RetValStr = sName
'          END IF
'        END IF
'
'      CASE "ADD"
'        ' -- If all went fine ...
'        IF thinBasic_CheckOpenParens_Mandatory AND thinBasic_ErrorFree THEN
'          IF thinBasic_CheckCloseParens_Mandatory AND thinBasic_ErrorFree THEN
'            IF lRecursive THEN
'              SELECT CASE CONST$ @RetValStr
'                CASE "WORKBOOKS"
'                  FUNCTION = %thinBasic_ReturnNumber
'
'                  GLOBAL oExcelWorkBook AS Excel_Workbook
'                  OBJECT CALL oObject(1&).Workbooks.Add TO oExcelWorkBook
'
'                  DIM t AS IUNKNOWN
'                  t = oExcelWorkBook
'                  t.AddRef
'                  ' -- Returning the object pointer
'                  @RetValNum = OBJPTR(oExcelWorkBook)
'
'                CASE ELSE
'                  thinbasic_RunTimeError ( %ERR__MODULE_SPECIFIC, @RetValStr+" not recognized" )
'
'              END SELECT
'            ELSE
'              thinbasic_RunTimeError ( %ERR__MODULE_SPECIFIC, "You cannot use Add at this level" )
'            END IF
'          END IF
'        END IF
'
'      CASE "CLOSE"
'        '---If all went fine ...
'        IF thinBasic_CheckOpenParens_Mandatory AND thinBasic_ErrorFree THEN
'          IF thinBasic_CheckCloseParens_Mandatory AND thinBasic_ErrorFree THEN
'
'            IF lRecursive THEN
'              SELECT CASE CONST$ @RetValStr
'                CASE "ACTIVEWINDOW"
'                  FUNCTION = %thinBasic_ReturnNone
'
'                  OBJECT CALL oObject(1&).ActiveWindow.Close
'
'                CASE ELSE
'                  thinbasic_RunTimeError ( %ERR__MODULE_SPECIFIC, @RetValStr+" not recognized" )
'              END SELECT
'            ELSE
'              thinbasic_RunTimeError ( %ERR__MODULE_SPECIFIC, "You cannot use Close at this level" )
'            END IF
'          END IF
'        END IF
'
'      CASE "QUIT"
'        ' -- If all went fine ...
'        IF thinBasic_CheckOpenParens_Mandatory AND thinBasic_ErrorFree THEN
'          IF thinBasic_CheckCloseParens_Mandatory AND thinBasic_ErrorFree THEN
'
'            IF lRecursive THEN
'              thinbasic_RunTimeError ( %ERR__MODULE_SPECIFIC, "You cannot use Quit at this level" )
'            ELSE
'              OBJECT CALL oObject(1&).Quit
'            END IF
'
'          END IF
'        END IF
'
'
'
'      CASE ELSE
'        thinBasic_RunTimeError(%ERR__CLASS_METHODPROPERTY_NOTFOUND, sName)
'    END SELECT
'
'    IF thinBasic_CheckPoint_Optional THEN
'      FUNCTION = Excel_Application_ClassFunction ( _
'                                                  pObject   , _
'                                                  lAction   , _
'                                                  RetValNum , _
'                                                  RetValStr , _
'                                                  %TRUE       _
'                                                )
'    END IF
'
'END FUNCTION

  '----------------------------------------------------------------------------
  ' Class cExcel_Application implements ...
  '----------------------------------------------------------------------------
  Class cExcel_Application

    '----------------------------------------------------------------------------
    '---Class Private variables
    '----------------------------------------------------------------------------
    Instance pExcel_Application As Excel_Application
    
    '----------------------------------------------------------------------------
    ' We need to define an interface that will contain public methods
    '----------------------------------------------------------------------------
    Interface iExcel_Application

      Inherit IUnknown

      '----------------------------------------------------------------------------
      ' Creator
      '----------------------------------------------------------------------------
      Method cExcel_Application_Create() As Dword
        pExcel_Application = ANYCOM $PROGID_Excel_Application
        If IsNothing(pExcel_Application) Then Exit Method
      End Method

      '----------------------------------------------------------------------------
      ' Destructor
      '----------------------------------------------------------------------------
      Method cExcel_Application_Destroy() As Dword 
        Set pExcel_Application = Nothing
      End Method

      '----------------------------------------------------------------------------
      ' GetClassObject
      '----------------------------------------------------------------------------
      Method cExcel_Application_GetClassObject() As Dword 
        Method = VarPtr(pExcel_Application)
      End Method

'      '----------------------------------------------------------------------------
'      ' Open
'      '----------------------------------------------------------------------------
'      Method cADODB_Connection_Open(ByVal sConnect As String) As Dword
'        pConnection.Open(sConnect)
'        Method = pConnection.State
'      End Method
'
'
'      '----------------------------------------------------------------------------
'      ' Close
'      '----------------------------------------------------------------------------
'      Method cADODB_Connection_Close() As Dword
'        pConnection.Close
'        Method = pConnection.State
'      End Method
'
'      '----------------------------------------------------------------------------
'      ' GET ConnectionString
'      '----------------------------------------------------------------------------
'      Method cADODB_Connection_GET_ConnectionString() As String
'        Method = pConnection.ConnectionString
'      End Method
'
'      '----------------------------------------------------------------------------
'      ' SET ConnectionString
'      '----------------------------------------------------------------------------
'      Method cADODB_Connection_SET_ConnectionString(ByVal lConnString As String) As String
'        pConnection.ConnectionString = lConnString
'        Method = pConnection.ConnectionString
'      End Method
'
'      '----------------------------------------------------------------------------
'      ' GET Version
'      '----------------------------------------------------------------------------
'      Method cExcel_Application_GET_Version() As String
'        Object Get pExcel_Application.Version To dummyVariant
'        Method = Variant$(dummyVariant)
'      End Method
'
'      '----------------------------------------------------------------------------
'      ' GET Visible
'      '----------------------------------------------------------------------------
'      Method cExcel_Application_GET_Visible() As Long
'        Object Get pExcel_Application.Visible To dummyVariant
'        Method = Variant#(dummyVariant)
'      End Method
'
'      '----------------------------------------------------------------------------
'      ' SET Visible
'      '----------------------------------------------------------------------------
'      Method cExcel_Application_SET_Visible(ByVal lVisible As Long) As Long
'        'Object Get Interface.Member[.Member.] [([[paramname =] param1 [, ...]])] To ResultVar
'        'Object Let Interface.Member[.Member.] [([[paramname =] param1 [, ...]])] = ValueVar
'        'Object Set Interface.Member[.Member.] [([[paramname =] param1 [, ...]])] = ValueVar
'        dummyVariant = lVisible
'        Object Let pExcel_Application.Visible = dummyVariant
'        Method = lVisible
'      End Method
'
'      '----------------------------------------------------------------------------
'      ' GET CommandTimeOut
'      '----------------------------------------------------------------------------
'      Method cADODB_Connection_GET_CommandTimeOut() As Long
'        Method = pConnection.CommandTimeOut
'      End Method
'
'      '----------------------------------------------------------------------------
'      ' SET CommandTimeout
'      '----------------------------------------------------------------------------
'      Method cADODB_Connection_SET_CommandTimeOut(ByVal lTimeOut As Long) As Long
'        pConnection.CommandTimeOut = lTimeOut
'        Method = pConnection.CommandTimeOut
'      End Method

    End Interface

  End Class
  '----------------------------------------------------------------------------
  '----------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  Sub cExcel_Application_Create( ByVal pObject As Long )
  '----------------------------------------------------------------------------
'    Local sConnect  As String
'    Local pp        As Byte

    '---Optionally checks if there is a (
'    LOCAL eSize           AS EXT
'    Local KeyToUCase      As Ext
'    
'    pp = thinBasic_CheckOpenParens_Optional
'    thinBasic_ParseString sConnect
'    If thinBasic_CheckComma_Optional Then
'      thinBasic_ParseNumber KeyToUCase
'    End If
'    If pp Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then

      'Here is the magic !!!!!!!!!!!!
      'Use the magic REDIM ... AT and the trick is done
      'We are defining a new class interface variable at the pointer location allocated by thinBasic and passed to all methods
      ReDim oObject(1& To 1&) As iExcel_Application At pObject

      '---We instantiate a new class
      oObject(1&) = Class "cExcel_Application"
      '---We call sSet in order to assign initial value
      oObject(1&).cExcel_Application_Create()


    'else
      '---emit some error
    End If
  End Sub


  '----------------------------------------------------------------------------
  Sub cExcel_Application_Destroy(ByVal pObject As Long)
  '----------------------------------------------------------------------------
    'If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Application At pObject
      oObject(1&).cExcel_Application_Destroy()
    End If
  End Sub

  '----------------------------------------------------------------------------
  Function cExcel_Application_GetClassObject(ByVal pObject As Long) As Ext
  '----------------------------------------------------------------------------
    'If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Application At pObject
      Function = oObject(1&).cExcel_Application_GetClassObject()
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Application_Method_Quit(ByVal pObject As Long) As Long
  '----------------------------------------------------------------------------
    'If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Application At pObject
      ReDim lExcel_Application(1& To 1&) As Excel_Application At oObject(1&).cExcel_Application_GetClassObject
     
      Object Call lExcel_Application(1&).ActiveWindow.Close
      Object Call lExcel_Application(1&).Quit

    End If
  End Function

  '--------------------------------------------------------------------------------------
  Function cExcel_Application_Property_Visible(ByVal pObject As Long, ByVal lAction As Long) As Long
  '--------------------------------------------------------------------------------------
  '  Properties functions has 2 params:
  '   pObject   is a pointer to module class or any other memory area handled by a pseudo class
  '   lAction   can be %Class_Action_Get, %Class_Action_Set, %Class_Action_None
  '             depending on how the property is used in the script
  '             %Class_Action_Get   when property is used inside an expression
  '             %Class_Action_Set   when property is used at the beginning of a line
  '             %Class_Action_None  SHOULD NEVER HAPPEN
  '--------------------------------------------------------------------------------------
  ' Syntax:
  '   
  '--------------------------------------------------------------------------------------
    Local lEqualType    As Long
    Local lDummyVariant As variant
    
    ReDim oObject(1& To 1&) As iExcel_Application At pObject
    ReDim lExcel_Application(1& To 1&) As Excel_Application At oObject(1&).cExcel_Application_GetClassObject

    Select Case Long lAction
      '---GET action happen usually when property is used inside an expression
      Case %Class_Action_Get                                                  
        '---Return value
        Object Get lExcel_Application(1&).Visible To lDummyVariant
        Function = Variant#(lDummyVariant)
        
      '---SET action happen usually when property is used directly as first token in a source code line
      Case %Class_Action_Set 
        '---Here programmer has to parse equal (=) sign and than a numeric or string expression
        '---But remember: equal sign can also be += &= -= *= /= \=
        lEqualType = thinBasic_CheckEqualType_Mandatory
        If thinBasic_ErrorFree Then
          Select Case Long lEqualType

            '---Assignment
            Case  %Equal_EQ
              lDummyVariant = thinBasic_Parse1Number
              Object Let lExcel_Application(1&).Visible = lDummyVariant
              'Function = oObject(1&).cExcel_Application_SET_Visible(thinBasic_Parse1Number)

'            '---Assignment with concatenation      
'            Case  %Equal_PLUSEQUAL, _
'                  %Equal_CONCEQUAL, _
'                  %Equal_POINTEQUAL
'
            Case Else
              '---Not supperted for strings 
              MsgBox FuncName$ & Str$(lEqualType)
              
          End Select
        End If
                
      Case %Class_Action_None
        '---Should not happen but ... I will leave in any case
    End Select

  End Function

  '--------------------------------------------------------------------------------------
  Function cExcel_Application_Property_AlertBeforeOverwriting(ByVal pObject As Long, ByVal lAction As Long) As Long
  '--------------------------------------------------------------------------------------
  '  Properties functions has 2 params:
  '   pObject   is a pointer to module class or any other memory area handled by a pseudo class
  '   lAction   can be %Class_Action_Get, %Class_Action_Set, %Class_Action_None
  '             depending on how the property is used in the script
  '             %Class_Action_Get   when property is used inside an expression
  '             %Class_Action_Set   when property is used at the beginning of a line
  '             %Class_Action_None  SHOULD NEVER HAPPEN
  '--------------------------------------------------------------------------------------
  ' Syntax:
  '   
  '--------------------------------------------------------------------------------------
    Local lEqualType    As Long
    Local vDummyVariant As variant
    
    ReDim oObject(1& To 1&) As iExcel_Application At pObject
    ReDim lExcel_Application(1& To 1&) As Excel_Application At oObject(1&).cExcel_Application_GetClassObject

    Select Case Long lAction
      '---GET action happen usually when property is used inside an expression
      Case %Class_Action_Get                                                  
        '---Return value
        Object Get lExcel_Application(1&).AlertBeforeOverwriting To vDummyVariant
        Function = Variant#(vDummyVariant)
        
      '---SET action happen usually when property is used directly as first token in a source code line
      Case %Class_Action_Set 
        '---Here programmer has to parse equal (=) sign and than a numeric or string expression
        '---But remember: equal sign can also be += &= -= *= /= \=
        lEqualType = thinBasic_CheckEqualType_Mandatory
        If thinBasic_ErrorFree Then
          Select Case Long lEqualType

            '---Assignment
            Case  %Equal_EQ
              vDummyVariant = thinBasic_Parse1Number
              Object Let lExcel_Application(1&).AlertBeforeOverwriting = vDummyVariant

'            '---Assignment with concatenation      
'            Case  %Equal_PLUSEQUAL, _
'                  %Equal_CONCEQUAL, _
'                  %Equal_POINTEQUAL
'
            Case Else
              '---Not supperted for strings 
              MsgBox FuncName$ & Str$(lEqualType) 
              
          End Select
        End If
                
      Case %Class_Action_None
        '---Should not happen but ... I will leave in any case
    End Select

  End Function

  '--------------------------------------------------------------------------------------
  Function cExcel_Application_Property_DisplayAlerts(ByVal pObject As Long, ByVal lAction As Long) As Long
  '--------------------------------------------------------------------------------------
  '  Properties functions has 2 params:
  '   pObject   is a pointer to module class or any other memory area handled by a pseudo class
  '   lAction   can be %Class_Action_Get, %Class_Action_Set, %Class_Action_None
  '             depending on how the property is used in the script
  '             %Class_Action_Get   when property is used inside an expression
  '             %Class_Action_Set   when property is used at the beginning of a line
  '             %Class_Action_None  SHOULD NEVER HAPPEN
  '--------------------------------------------------------------------------------------
  ' Syntax:
  '   
  '--------------------------------------------------------------------------------------
    Local lEqualType    As Long
    Local vDummyVariant As variant
    
    ReDim oObject(1& To 1&) As iExcel_Application At pObject
    ReDim lExcel_Application(1& To 1&) As Excel_Application At oObject(1&).cExcel_Application_GetClassObject

    Select Case Long lAction
      '---GET action happen usually when property is used inside an expression
      Case %Class_Action_Get                                                  
        '---Return value
        Object Get lExcel_Application(1&).DisplayAlerts To vDummyVariant
        Function = Variant#(vDummyVariant)
        
      '---SET action happen usually when property is used directly as first token in a source code line
      Case %Class_Action_Set 
        '---Here programmer has to parse equal (=) sign and than a numeric or string expression
        '---But remember: equal sign can also be += &= -= *= /= \=
        lEqualType = thinBasic_CheckEqualType_Mandatory
        If thinBasic_ErrorFree Then
          Select Case Long lEqualType

            '---Assignment
            Case  %Equal_EQ
              vDummyVariant = thinBasic_Parse1Number
              Object Let lExcel_Application(1&).DisplayAlerts = vDummyVariant

'            '---Assignment with concatenation      
'            Case  %Equal_PLUSEQUAL, _
'                  %Equal_CONCEQUAL, _
'                  %Equal_POINTEQUAL
'
            Case Else
              '---Not supperted for strings 
              MsgBox FuncName$ & Str$(lEqualType) 
              
          End Select
        End If
                
      Case %Class_Action_None
        '---Should not happen but ... I will leave in any case
    End Select

  End Function

  '--------------------------------------------------------------------------------------
  Function cExcel_Application_Property_Version(ByVal pObject As Long, ByVal lAction As Long) As String
  '--------------------------------------------------------------------------------------
  '  Properties functions has 2 params:
  '   pObject   is a pointer to module class or any other memory area handled by a pseudo class
  '   lAction   can be %Class_Action_Get, %Class_Action_Set, %Class_Action_None
  '             depending on how the property is used in the script
  '             %Class_Action_Get   when property is used inside an expression
  '             %Class_Action_Set   when property is used at the beginning of a line
  '             %Class_Action_None  SHOULD NEVER HAPPEN
  '--------------------------------------------------------------------------------------
  ' Syntax:
  '   cString.Value
  '--------------------------------------------------------------------------------------
    Local lEqualType    As Long
    Local lDummyVariant As variant
    
    ReDim oObject(1& To 1&) As iExcel_Application At pObject
    ReDim lExcel_Application(1& To 1&) As Excel_Application At oObject(1&).cExcel_Application_GetClassObject

    Select Case Long lAction
      '---GET action happen usually when property is used inside an expression
      Case %Class_Action_Get                                                  
        '---Return value
        'Function = oObject(1&).cExcel_Application_GET_Version
        Object Get lExcel_Application(1&).Version To lDummyVariant
        Function = Variant$(lDummyVariant)
        
      '---SET action happen usually when property is used directly as first token in a source code line
      Case %Class_Action_Set 
'        '---Here programmer has to parse equal (=) sign and than a numeric or string expression
'        '---But remember: equal sign can also be += &= -= *= /= \=
'        lEqualType = thinBasic_CheckEqualType_Mandatory
'        If thinBasic_ErrorFree Then
'          Select Case Long lEqualType
'
'            '---Assignment
'            Case  %Equal_EQ
'            
'              Function = oObject(1&).cADODB_Connection_SET_ConnectionString(thinBasic_Parse1String)
'
''            '---Assignment with concatenation      
''            Case  %Equal_PLUSEQUAL, _
''                  %Equal_CONCEQUAL, _
''                  %Equal_POINTEQUAL
''
''                  oObject(1&).sSet(oObject(1&).sGet() & thinBasic_Parse1String)
'            Case Else
'              '---Not supperted for strings 
'              MsgBox FuncName$ & Str$(lEqualType)
'          End Select
'        End If
                
      Case %Class_Action_None
        '---Should not happen but ... I will leave in any case
    End Select

  End Function


  '--------------------------------------------------------------------------------------
  Function cExcel_Application_Property_ActiveWindow(ByVal pObject As Long, ByVal lAction As Long) As String
  '--------------------------------------------------------------------------------------
  '  Properties functions has 2 params:
  '   pObject   is a pointer to module class or any other memory area handled by a pseudo class
  '   lAction   can be %Class_Action_Get, %Class_Action_Set, %Class_Action_None
  '             depending on how the property is used in the script
  '             %Class_Action_Get   when property is used inside an expression
  '             %Class_Action_Set   when property is used at the beginning of a line
  '             %Class_Action_None  SHOULD NEVER HAPPEN
  '--------------------------------------------------------------------------------------
  ' Syntax:
  '   
  '--------------------------------------------------------------------------------------
    Local lEqualType    As Long
    Local vDummyVariant As variant
    Local sSubItem    As String
    
    ReDim oObject(1& To 1&) As iExcel_Application At pObject
    ReDim lExcel_Application(1& To 1&) As Excel_Application At oObject(1&).cExcel_Application_GetClassObject

    If thinBasic_CheckPoint_Optional Then
      sSubItem = thinBasic_GetUnknownToken   
    End If
    
    If thinBasic_ErrorFree Then
      Select Case UCase$(sSubItem)
        Case "CLOSE"
          Select Case Long lAction

            '---GET action happen usually when property is used inside an expression
            Case %Class_Action_Get                                                  
              '---Return value
              Object Get lExcel_Application(1&).ActiveWindow.Close To vDummyVariant
              Function = Variant$(vDummyVariant)
              
            '---SET action happen usually when property is used directly as first token in a source code line
            Case %Class_Action_Set 
              '---Here programmer has to parse equal (=) sign and than a numeric or string expression
              '---But remember: equal sign can also be += &= -= *= /= \=
              lEqualType = thinBasic_CheckEqualType_Mandatory
              If thinBasic_ErrorFree Then
                Select Case Long lEqualType
      
                  '---Assignment
                  Case  %Equal_EQ
                    vDummyVariant = thinBasic_Parse1Number
                    Object Let lExcel_Application(1&).ActiveWindow.Close = vDummyVariant
      
      '            '---Assignment with concatenation      
      '            Case  %Equal_PLUSEQUAL, _
      '                  %Equal_CONCEQUAL, _
      '                  %Equal_POINTEQUAL
      '
                  Case Else
                    '---Not supperted for strings 
                    MsgBox FuncName$ & Str$(lEqualType) 
                    
                End Select
              End If
                      
            Case %Class_Action_None
              '---Should not happen but ... I will leave in any case
          End Select
      
      '  Case ???
      
      End Select  'UCase$(sSubItem)
      
    End If
  End Function



