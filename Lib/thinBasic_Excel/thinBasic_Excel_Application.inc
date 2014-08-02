'
' Excel_Application class wrapper
'

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


    End Interface

  End Class
  '----------------------------------------------------------------------------
  '----------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  Sub cExcel_Application_Create( ByVal pObject As Long )
  '----------------------------------------------------------------------------

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
  '--------------------------------------------------------------------------------------
  ' Help: True if Microsoft Excel displays a message before overwriting nonblank cells during a drag-and-drop editing operation  
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



