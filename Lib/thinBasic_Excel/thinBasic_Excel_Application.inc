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

      '----------------------------------------------------------------------------
      ' IsObject
      '----------------------------------------------------------------------------
      Method cExcel_Application_IsNothing() As Dword 
        Method = IsNothing(pExcel_Application)
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
  Function cExcel_Application_GetClassObject(ByVal pObject As Long) As Dword
  '----------------------------------------------------------------------------
    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Application At pObject
      Function = oObject(1&).cExcel_Application_GetClassObject()
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Application_Method_IsNothing(ByVal pObject As Long) As Long
  '----------------------------------------------------------------------------
    Local lIsObject As Long
    If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Application At pObject
     
      lIsObject = oObject(1&).cExcel_Application_IsNothing
      If IsTrue(lIsObject) Then
        lIsObject = %TRUE
      Else
        lIsObject = %FALSE
      End If
      
      Function = lIsObject
      
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Application_Method_Quit(ByVal pObject As Long) As Long
  '----------------------------------------------------------------------------
    If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Application At pObject
      ReDim lExcel_Application(1& To 1&) As Excel_Application At oObject(1&).cExcel_Application_GetClassObject
     
      Object Call lExcel_Application(1&).ActiveWindow.Close
      Object Call lExcel_Application(1&).Quit

    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Application_Method_SendKeys(ByVal pObject As Long) As Long
  '----------------------------------------------------------------------------
    Local pp    As Byte
    Local sKeys As String
    Local vKeys As variant
    Local eWait As Ext
    Local vWait As variant
    
    pp = thinBasic_CheckOpenParens_Optional
    thinBasic_ParseString sKeys
    If thinBasic_CheckComma_Optional Then
      thinBasic_ParseNumber eWait
    End If
    If pp Then thinBasic_CheckCloseParens_Mandatory              


    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Application At pObject
      ReDim lExcel_Application(1& To 1&) As Excel_Application At oObject(1&).cExcel_Application_GetClassObject
                                                               
      vKeys = sKeys
      vWait = eWait As Long
      Object Call lExcel_Application(1&).SendKeys(vKeys, vWait)

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
    Local sSubItem      As String                      
    Local eSaveChanges  As Ext
    Local vSaveChanges  As Variant
    Local sFileName     As String
    Local vFileName     As Variant
    
    ReDim oObject(1& To 1&) As iExcel_Application At pObject
    ReDim lExcel_Application(1& To 1&) As Excel_Application At oObject(1&).cExcel_Application_GetClassObject

    If thinBasic_CheckPoint_Optional Then
      sSubItem = thinBasic_GetUnknownToken   
    End If
    
    If thinBasic_ErrorFree Then
      Select Case UCase$(sSubItem)
        Case "CLOSE"                                            
          If thinBasic_CheckOpenParens_Optional Then
            thinBasic_ParseNumber eSaveChanges            
            Let vSaveChanges = Int(eSaveChanges)
            If thinBasic_CheckComma_Optional Then
              thinBasic_ParseString sFileName   
              Let vFileName = sFileName
            End If
            thinBasic_CheckCloseParens_Mandatory
          End If
                    
          Object Call lExcel_Application(1&).ActiveWindow.Close(vSaveChanges, vFileName)
           
'          Select Case Long lAction
'
'            '---GET action happen usually when property is used inside an expression
'            Case %Class_Action_Get                                                  
'              '---Return value
'              Object Get lExcel_Application(1&).ActiveWindow.Close To vDummyVariant
'              Function = Variant$(vDummyVariant)
'              
'            '---SET action happen usually when property is used directly as first token in a source code line
'            Case %Class_Action_Set 
'              '---Here programmer has to parse equal (=) sign and than a numeric or string expression
'              '---But remember: equal sign can also be += &= -= *= /= \=
''              lEqualType = thinBasic_CheckEqualType_Mandatory
''              If thinBasic_ErrorFree Then
''                Select Case Long lEqualType
''      
''                  '---Assignment
''                  Case  %Equal_EQ
''                    vDummyVariant = thinBasic_Parse1Number
'                    Object Let lExcel_Application(1&).ActiveWindow.Close = vDummyVariant
''      
''      '            '---Assignment with concatenation      
''      '            Case  %Equal_PLUSEQUAL, _
''      '                  %Equal_CONCEQUAL, _
''      '                  %Equal_POINTEQUAL
''      '
'                  Case Else
'                    '---Not supperted for strings 
'                    MsgBox FuncName$ & Str$(lEqualType) 
'                    
'                End Select
'              End If
'                      
'            Case %Class_Action_None
'              '---Should not happen but ... I will leave in any case
'          End Select
      
      '  Case ???
      
      End Select  'UCase$(sSubItem)
      
    End If
  End Function

  '--------------------------------------------------------------------------------------
  Function cExcel_Application_Property_Workbooks(ByVal pObject As Long, ByVal lAction As Long) As String
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
    Local sSubItem      As String
    Local sFileName     As String
    Local vFileName     As variant

    Local eUpdateLinks  As Ext
    Local vUpdateLinks  As variant
    Local eReadOnly     As Ext
    Local vReadOnly     As variant
    Local eFileFormat   As Ext
    Local vFileFormat   As variant
    Local sPassword     As String
    Local vPassword     As variant

    Local pp            As Byte
    
    Static lDummyExcelWorkbook As Excel_Workbook
    
    ReDim oObject(1& To 1&) As iExcel_Application At pObject
    ReDim lExcel_Application(1& To 1&) As Excel_Application At oObject(1&).cExcel_Application_GetClassObject

    If thinBasic_CheckPoint_Optional Then
      sSubItem = thinBasic_GetUnknownToken   
    End If
    
    If thinBasic_ErrorFree Then
      Select Case UCase$(sSubItem)

        Case "COUNT"
          Object Get lExcel_Application(1&).Workbooks.Count To vDummyVariant
          Function = Str$(variant#(vDummyVariant))
          
        Case "ADD"
          Object Call lExcel_Application(1&).Workbooks.Add To lDummyExcelWorkbook
          Function = Str$(ObjPtr(lDummyExcelWorkbook))

        Case "OPEN"   

'  MEMBER CALL Open <682> (
                            'In Filename As WSTRING<0>, 
                            'Opt In UpdateLinks As VARIANT<1>, 
                            'Opt In ReadOnly As VARIANT<2>, Opt _
                            'In Format As VARIANT<3>, 
                            'Opt In Password As VARIANT<4>, 
                            'Opt In WriteResPassword As VARIANT<5>, 
                            'Opt In IgnoreReadOnlyRecommended As VARIANT<6>, 
                            'Opt In Origin As VARIANT<7>, 
                            'Opt In Delimiter As VARIANT<8>, 
                            'Opt In Editable As VARIANT<9>, 
                            'Opt In PB_Notify As VARIANT<10>, 
                            'Opt In Converter As VARIANT<11>, 
                            'Opt In AddToMru As VARIANT<12>
                          ') As IDispatch

'    Local eUpdateLinks  As Ext
'    Local vUpdateLinks  As variant
'    Local eReadOnly     As Ext
'    Local vReadOnly     As variant
'    Local eFileFormat   As Ext
'    Local vFileFormat   As variant
'    Local sPassword     As String
'    Local vPassword     As variant
          
          eUpdateLinks  = %FALSE
          eReadOnly     = %FALSE
          eFileFormat   = %XlFileFormat.xlWorkbookNormal
          
          pp = thinBasic_CheckOpenParens_Optional
          thinBasic_ParseString sFileName
          If thinBasic_CheckComma_Optional Then
            thinBasic_ParseNumber eUpdateLinks
            If thinBasic_CheckComma_Optional Then
              thinBasic_ParseNumber eReadOnly
              If thinBasic_CheckComma_Optional Then
                thinBasic_ParseNumber eFileFormat
                If thinBasic_CheckComma_Optional Then
                  thinBasic_ParseString sPassword
                End If
              End If
            End If
          End If
          
          If pp Then thinBasic_CheckCloseParens_Mandatory              
          
          If thinBasic_ErrorFree Then
            vFileName     = sFileName
            vUpdateLinks  = eUpdateLinks
            vReadOnly     = eReadOnly   
            vFileFormat   = eFileFormat 
            vPassword     = sPassword 
            
'            Object Call lExcel_Application(1&).Workbooks.Open( _
'                                                                vFileName     , _
'                                                                vUpdateLinks  , _
'                                                                vReadOnly     , _
'                                                                vFileFormat   , _
'                                                                vPassword       _
'            Object Call lExcel_Application(1&).Workbooks.Open( _
'                                                                vFileName     , _
'                                                                vUpdateLinks  , _
'                                                                vReadOnly     , _
'                                                                vFileFormat     _
'                                                              ) To lDummyExcelWorkbook
            Object Call lExcel_Application(1&).Workbooks.Open( _
                                                                vFileName     , _
                                                                vUpdateLinks  , _
                                                                vReadOnly     , _
                                                                vFileFormat   , _
                                                                vPassword       _
                                                              ) To lDummyExcelWorkbook
            Function = Str$(ObjPtr(lDummyExcelWorkbook))
          End If

        Case "CLOSE"
          Object Call lExcel_Application(1&).Workbooks.Close
        
      End Select  'UCase$(sSubItem)
      
    End If
  End Function


  '--------------------------------------------------------------------------------------
  Function cExcel_Application_Property_ActiveWorkbook(ByVal pObject As Long, ByVal lAction As Long) As String
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
    Local sSubItem      As String
    Local sFileName     As String
    Local vFileName     As variant
    Local pp            As Byte
    
    Static lDummyExcelWorkbook  As Excel_Workbook
    Static lDummyExcelSheet     As Excel_Worksheet
    
    ReDim oObject(1& To 1&) As iExcel_Application At pObject
    ReDim lExcel_Application(1& To 1&) As Excel_Application At oObject(1&).cExcel_Application_GetClassObject

    If thinBasic_CheckPoint_Optional Then
      sSubItem = thinBasic_GetUnknownToken   
    End If
    
    If thinBasic_ErrorFree Then
      Select Case UCase$(sSubItem)
        Case ""
          Object Get lExcel_Application(1&).ActiveWorkbook To lDummyExcelWorkbook
          Function = Str$(ObjPtr(lDummyExcelWorkbook))

'        Case "COUNT"
'          Object Get lExcel_Application(1&).Workbooks.Count To vDummyVariant
'          Function = Str$(Variant#(vDummyVariant))
'          
'        Case "ADD"
'          Object Call lExcel_Application(1&).Workbooks.Add To lDummyExcelWorkbook
'          Function = Str$(ObjPtr(lDummyExcelWorkbook))
'
'        Case "OPEN"            
'          pp = thinBasic_CheckOpenParens_Optional
'          thinBasic_ParseString sFileName
'          If pp Then thinBasic_CheckCloseParens_Mandatory              
'          
'          If thinBasic_ErrorFree Then
'            vFileName = sFileName
'            Object Call lExcel_Application(1&).Workbooks.Open(vFileName) To lDummyExcelWorkbook
'            Function = Str$(ObjPtr(lDummyExcelWorkbook))
'          End If
              
      End Select  'UCase$(sSubItem)
      
    End If
  End Function

  '--------------------------------------------------------------------------------------
  Function cExcel_Application_Property_ActiveSheet(ByVal pObject As Long, ByVal lAction As Long) As String
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
    Local sSubItem      As String
    Local sFileName     As String
    Local vFileName     As variant
    Local pp            As Byte
    
    Static lDummyExcelWorkbook  As Excel_Workbook
    Static lDummyExcelWorkSheet As Excel_Worksheet
    
    ReDim oObject(1& To 1&) As iExcel_Application At pObject
    ReDim lExcel_Application(1& To 1&) As Excel_Application At oObject(1&).cExcel_Application_GetClassObject

    If thinBasic_CheckPoint_Optional Then
      sSubItem = thinBasic_GetUnknownToken   
    End If
    
    If thinBasic_ErrorFree Then
      Select Case UCase$(sSubItem)
        Case ""
          Object Get lExcel_Application(1&).ActiveSheet To lDummyExcelWorksheet
          Function = Str$(ObjPtr(lDummyExcelWorksheet))

'        Case ""
'          Object Call lExcel_Application(1&).ActiveWorkbook To lDummyExcelWorkbook
'          Function = Str$(ObjPtr(lDummyExcelWorkbook))
'
'        Case "COUNT"
'          Object Get lExcel_Application(1&).Workbooks.Count To vDummyVariant
'          Function = Str$(Variant#(vDummyVariant))
'          
'        Case "ADD"
'          Object Call lExcel_Application(1&).Workbooks.Add To lDummyExcelWorkbook
'          Function = Str$(ObjPtr(lDummyExcelWorkbook))
'
'        Case "OPEN"            
'          pp = thinBasic_CheckOpenParens_Optional
'          thinBasic_ParseString sFileName
'          If pp Then thinBasic_CheckCloseParens_Mandatory              
'          
'          If thinBasic_ErrorFree Then
'            vFileName = sFileName
'            Object Call lExcel_Application(1&).Workbooks.Open(vFileName) To lDummyExcelWorkbook
'            Function = Str$(ObjPtr(lDummyExcelWorkbook))
'          End If
              
      End Select  'UCase$(sSubItem)
      
    End If
  End Function

  '--------------------------------------------------------------------------------------
  Function cExcel_Application_Property_Selection(ByVal pObject As Long, ByVal lAction As Long) As String
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
    Local sSubItem      As String
    Local sFileName     As String
    Local vFileName     As variant
    Local pp            As Byte
    
    Static lDummyExcelWorkbook  As Excel_Workbook
    Static lDummyExcelRange As Range
    
    ReDim oObject(1& To 1&) As iExcel_Application At pObject
    ReDim lExcel_Application(1& To 1&) As Excel_Application At oObject(1&).cExcel_Application_GetClassObject

    'If thinBasic_CheckPoint_Optional Then
    '  sSubItem = thinBasic_GetUnknownToken   
    'End If
    
    If thinBasic_ErrorFree Then
      Select Case UCase$(sSubItem)
        Case ""
          Object Get lExcel_Application(1&).Selection To lDummyExcelRange
          Function = Str$(ObjPtr(lDummyExcelRange))

'        Case ""
'          Object Call lExcel_Application(1&).ActiveWorkbook To lDummyExcelWorkbook
'          Function = Str$(ObjPtr(lDummyExcelWorkbook))
'
'        Case "COUNT"
'          Object Get lExcel_Application(1&).Workbooks.Count To vDummyVariant
'          Function = Str$(Variant#(vDummyVariant))
'          
'        Case "ADD"
'          Object Call lExcel_Application(1&).Workbooks.Add To lDummyExcelWorkbook
'          Function = Str$(ObjPtr(lDummyExcelWorkbook))
'
'        Case "OPEN"            
'          pp = thinBasic_CheckOpenParens_Optional
'          thinBasic_ParseString sFileName
'          If pp Then thinBasic_CheckCloseParens_Mandatory              
'          
'          If thinBasic_ErrorFree Then
'            vFileName = sFileName
'            Object Call lExcel_Application(1&).Workbooks.Open(vFileName) To lDummyExcelWorkbook
'            Function = Str$(ObjPtr(lDummyExcelWorkbook))
'          End If
              
      End Select  'UCase$(sSubItem)
      
    End If
  End Function

  '--------------------------------------------------------------------------------------
  Function cExcel_Application_Property_ActiveCell(ByVal pObject As Long, ByVal lAction As Long) As String
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
    Local sSubItem      As String
    Local sFileName     As String
    Local vFileName     As variant
    Local pp            As Byte
    
    Static lDummyExcelWorkbook  As Excel_Workbook
    Static lDummyExcelRange     As Range 'Excel_Range
    
    ReDim oObject(1& To 1&) As iExcel_Application At pObject
    ReDim lExcel_Application(1& To 1&) As Excel_Application At oObject(1&).cExcel_Application_GetClassObject

    If thinBasic_CheckPoint_Optional Then
      sSubItem = thinBasic_GetUnknownToken   
    End If
    
    If thinBasic_ErrorFree Then
      Select Case UCase$(sSubItem)
        Case ""
          Object Get lExcel_Application(1&).ActiveCell To lDummyExcelRange
          Function = Str$(ObjPtr(lDummyExcelRange))

'        Case ""
'          Object Call lExcel_Application(1&).ActiveWorkbook To lDummyExcelWorkbook
'          Function = Str$(ObjPtr(lDummyExcelWorkbook))
'
'        Case "COUNT"
'          Object Get lExcel_Application(1&).Workbooks.Count To vDummyVariant
'          Function = Str$(Variant#(vDummyVariant))
'          
'        Case "ADD"
'          Object Call lExcel_Application(1&).Workbooks.Add To lDummyExcelWorkbook
'          Function = Str$(ObjPtr(lDummyExcelWorkbook))
'
'        Case "OPEN"            
'          pp = thinBasic_CheckOpenParens_Optional
'          thinBasic_ParseString sFileName
'          If pp Then thinBasic_CheckCloseParens_Mandatory              
'          
'          If thinBasic_ErrorFree Then
'            vFileName = sFileName
'            Object Call lExcel_Application(1&).Workbooks.Open(vFileName) To lDummyExcelWorkbook
'            Function = Str$(ObjPtr(lDummyExcelWorkbook))
'          End If
              
      End Select  'UCase$(sSubItem)
      
    End If
  End Function
