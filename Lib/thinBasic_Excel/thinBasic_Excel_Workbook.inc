  '----------------------------------------------------------------------------
  ' Class cExcel_Workbook implements ...
  '----------------------------------------------------------------------------
  Class cExcel_Workbook

    '----------------------------------------------------------------------------
    '---Class Private variables
    '----------------------------------------------------------------------------
    Instance pExcel_Workbook As Excel_Workbook
    
    '----------------------------------------------------------------------------
    ' We need to define an interface that will contain public methods
    '----------------------------------------------------------------------------
    Interface iExcel_Workbook

      Inherit IUnknown

      '----------------------------------------------------------------------------
      ' Creator
      '----------------------------------------------------------------------------
      Method cExcel_Workbook_Create(ByVal ptrExcel_Application As Dword) As Dword
        ReDim oExcelApp(1& To 1&) As Excel_Application At ptrExcel_Application 
        Object Call oExcelApp(1&).WorkBooks.Add To pExcel_Workbook

        If IsNothing(pExcel_Workbook) Then Exit Method
      End Method

      '----------------------------------------------------------------------------
      ' Creator Direct
      '----------------------------------------------------------------------------
      Method cExcel_Workbook_Create_Direct(ByVal ptrToDirectWorkbook As Dword) As Dword
        '---Here is the magic: we store at the memory address of pExcel_Workbook (that is a Excel Workbook COM pointer)
        '---The address passed by the script. The address is usually returned by another Class method or property
        '---Instructed to return an OBJPTR to an Excel Workbook
        Poke Dword, VarPtr(pExcel_Workbook), ptrToDirectWorkbook
      End Method

      '----------------------------------------------------------------------------
      ' Destructor
      '----------------------------------------------------------------------------
      Method cExcel_Workbook_Destroy() As Dword 
        Set pExcel_Workbook = Nothing
      End Method

      '----------------------------------------------------------------------------
      ' GetClassObject
      '----------------------------------------------------------------------------
      Method cExcel_Workbook_GetClassObject() As Dword 
        Method = VarPtr(pExcel_Workbook)
      End Method


    End Interface

  End Class
  '----------------------------------------------------------------------------
  '----------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  Sub cExcel_Workbook_Create( ByVal pObject As Long )
  '----------------------------------------------------------------------------
    Local pp          As Byte
    Local pExcelApp   As Ext

    pp = thinBasic_CheckOpenParens_Optional
      thinBasic_ParseNumber pExcelApp
    If pp Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then

      'Here is the magic !!!!!!!!!!!!
      'Use the magic REDIM ... AT and the trick is done
      'We are defining a new class interface variable at the pointer location allocated by thinBasic and passed to all methods
      ReDim oObject(1& To 1&) As iExcel_Workbook At pObject

      '---We instantiate a new class
      oObject(1&) = Class "cExcel_Workbook"
      '---We call sSet in order to assign initial value
      oObject(1&).cExcel_Workbook_Create(pExcelApp)


    'else
      '---emit some error
    End If
  End Sub

  '----------------------------------------------------------------------------
  Sub cExcel_Workbook_Create_Direct(ByVal pObject As Dword, ByVal pDirectClass As Dword)
  '----------------------------------------------------------------------------
    Local pp          As Byte
    Local pExcelApp   As Ext

    '---If all went fine ...
    If thinBasic_ErrorFree Then

      'Here is the magic !!!!!!!!!!!!
      'Use the magic REDIM ... AT and the trick is done
      'We are defining a new class interface variable at the pointer location allocated by thinBasic and passed to all methods
      ReDim oObject(1& To 1&) As iExcel_Workbook At pObject

      '---We instantiate a new class
      oObject(1&) = Class "cExcel_Workbook"
      '---We call sSet in order to assign initial value
      'oObject(1&).cExcel_Workbook_Create(pExcelApp)
      oObject(1&).cExcel_Workbook_Create_Direct(pDirectClass)

    'else
      '---emit some error
    End If
  End Sub

  '----------------------------------------------------------------------------
  Sub cExcel_Workbook_Destroy(ByVal pObject As Long)
  '----------------------------------------------------------------------------
    'If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Workbook At pObject
      oObject(1&).cExcel_Workbook_Destroy()
    End If
  End Sub

  '----------------------------------------------------------------------------
  Function cExcel_Workbook_GetClassObject(ByVal pObject As Long) As Dword
  '----------------------------------------------------------------------------
    'If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Workbook At pObject
      Function = oObject(1&).cExcel_Workbook_GetClassObject()
    End If
  End Function


  '----------------------------------------------------------------------------
  Function cExcel_Workbook_Method_SaveAs(ByVal pObject As Long) As Long
  '----------------------------------------------------------------------------
    Dim pp          As Byte
    Dim sFileName   As String
    Dim vFileName   As variant
    Dim eFileFormat As Ext
    Dim vFileFormat As variant
    Dim sPassword   As String
    Dim vPassword   As variant
    
    
    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Workbook At pObject
      ReDim lExcel_Workbook(1& To 1&) As Excel_Workbook At oObject(1&).cExcel_Workbook_GetClassObject

      pp = thinBasic_CheckOpenParens_Optional
      thinBasic_ParseString sFileName    
      eFileFormat = %XlFileFormat.xlWorkbookNormal
'MsgBox sFileName
      If thinBasic_CheckComma_Optional Then
        thinBasic_ParseNumber eFileFormat
'MsgBox Str$(eFileFormat)
        If thinBasic_CheckComma_Optional Then
          thinBasic_ParseString sPassword
'MsgBox sPassword
        End If
      End If
      If pp Then thinBasic_CheckCloseParens_Mandatory
      
      If thinBasic_ErrorFree Then
        vFileName   = sFileName
        vFileFormat = eFileFormat
        vPassword   = sPassword
        Object Call lExcel_Workbook(1&).SaveAs(vFileName, vFileFormat, vPassword)

      End If
      
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Workbook_Method_Activate(ByVal pObject As Long) As Long
  '----------------------------------------------------------------------------
    Dim pp        As Byte
    Dim sFileName As String
    Dim vFileName As variant
    
    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Workbook At pObject
      ReDim lExcel_Workbook(1& To 1&) As Excel_Workbook At oObject(1&).cExcel_Workbook_GetClassObject

      Object Call lExcel_Workbook(1&).Activate
      
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Workbook_Method_Save(ByVal pObject As Long) As Long
  '----------------------------------------------------------------------------
    Dim pp        As Byte
    Dim sFileName As String
    Dim vFileName As variant
    
    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Workbook At pObject
      ReDim lExcel_Workbook(1& To 1&) As Excel_Workbook At oObject(1&).cExcel_Workbook_GetClassObject

      Object Call lExcel_Workbook(1&).Save
      
    End If
  End Function

  '--------------------------------------------------------------------------------------
  Function cExcel_Workbook_Property_Worksheets(ByVal pObject As Long, ByVal lAction As Long) As String
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
    Local sIndex        As String
    Local eIndex        As Ext
    Local vIndex        As variant
    Local bIndexPresent As Byte
    Local sSubItem      As String

    Static lDummyExcelWorksheet As Excel_Worksheet
    
    ReDim oObject(1& To 1&) As iExcel_Workbook At pObject
    ReDim lExcel_Workbook(1& To 1&) As Excel_Workbook At oObject(1&).cExcel_Workbook_GetClassObject
    
    '---Determine possible syntax like: Worksheets(index)
    '---Here we parse the index if we have open parens
    If thinBasic_CheckOpenParens_Optional Then
      Select Case thinBasic_DetermineType
        Case %thinBasic_ReturnNumber
          thinBasic_ParseNumber eIndex
          Let vIndex = eIndex As Long
        Case %thinBasic_ReturnString
          thinBasic_ParseString sIndex
          Let vIndex = sIndex
      End Select
      bIndexPresent = %TRUE                         
      thinBasic_CheckCloseParens_Mandatory
    End If
    
    '---Now check if there is a sub item method/property
    If thinBasic_CheckPoint_Optional Then
      sSubItem = thinBasic_GetUnknownToken   
    End If
    
 
    If thinBasic_ErrorFree Then
      Select Case UCase$(sSubItem)
        Case ""
          Object Call lExcel_Workbook(1&).Worksheets(vIndex) To lDummyExcelWorksheet
'MsgBox FuncName$ & Str$(Variant#(vIndex))
          Function = Str$(ObjPtr(lDummyExcelWorksheet))
        
        Case "SELECT"
          Object Call lExcel_Workbook(1&).Worksheets(vIndex).Select 'To lDummyExcelWorksheet
'MsgBox FuncName$ & Str$(Variant#(vIndex))
          'Function = Str$(ObjPtr(lDummyExcelWorksheet))
        
        Case "COUNT"
          Object Get lExcel_Workbook(1&).Worksheets.Count To vDummyVariant
          Function = Str$(Variant#(vDummyVariant))
          
        Case "ADD"                                                 
          'Select Case Long lAction
          '
          '  '---GET action happen usually when property is used inside an expression
          '  Case %Class_Action_Get                                                  
              '---Return value
'MsgBox FuncName$ & " start"

'MsgBox "1"
'              Dim lDummyExcelWorkbookPtr As Dword Ptr
'MsgBox "2"
'              ReDim lDummyExcelWorkbook(1& To 1&) As Excel_Workbook At lDummyExcelWorkbookPtr
'MsgBox "3"
'              Object Call lExcel_Application(1&).Workbooks.Add To lDummyExcelWorkbook(1&)
'MsgBox "4"
'              Function = Str$(@lDummyExcelWorkbookPtr)
'MsgBox "5"

              Object Call lExcel_Workbook(1&).Worksheets.Add To lDummyExcelWorksheet
              Function = Str$(ObjPtr(lDummyExcelWorksheet))

'MsgBox FuncName$ & " end"
              
          '  '---SET action happen usually when property is used directly as first token in a source code line
          '  Case %Class_Action_Set 
          '    '---Here programmer has to parse equal (=) sign and than a numeric or string expression
          '    '---But remember: equal sign can also be += &= -= *= /= \=
          '    lEqualType = thinBasic_CheckEqualType_Mandatory
          '    If thinBasic_ErrorFree Then
          '      Select Case Long lEqualType
          '
          '        '---Assignment
          '        Case  %Equal_EQ
          '          vDummyVariant = thinBasic_Parse1Number
          '          Object Let lExcel_Application(1&).ActiveWindow.Close = vDummyVariant
          '
      '   '         '---Assignment with concatenation      
      '   '         Case  %Equal_PLUSEQUAL, _
      '   '               %Equal_CONCEQUAL, _
      '   '               %Equal_POINTEQUAL
      '   '
          '        Case Else
          '          '---Not supperted for strings 
          '          MsgBox FuncName$ & Str$(lEqualType) 
          '          
          '      End Select
          '    End If
          '            
          '  Case %Class_Action_None
          '    '---Should not happen but ... I will leave in any case
          'End Select
      
      '  Case ???
      
      End Select  'UCase$(sSubItem)
      
    End If
  End Function

  '--------------------------------------------------------------------------------------
  Function cExcel_Workbook_Property_Activesheet(ByVal pObject As Long, ByVal lAction As Long) As String
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
    
    Static lDummyExcelWorksheet As Excel_Worksheet
    
    ReDim oObject(1& To 1&) As iExcel_Workbook At pObject
    ReDim lExcel_Workbook(1& To 1&) As Excel_Workbook At oObject(1&).cExcel_Workbook_GetClassObject

    If thinBasic_CheckPoint_Optional Then
      sSubItem = thinBasic_GetUnknownToken   
    End If
    
    If thinBasic_ErrorFree Then
      Select Case UCase$(sSubItem)
        Case ""                                                 
          'Select Case Long lAction
          '
          '  '---GET action happen usually when property is used inside an expression
          '  Case %Class_Action_Get                                                  
              '---Return value
'MsgBox FuncName$ & " start"

'MsgBox "1"
'              Dim lDummyExcelWorkbookPtr As Dword Ptr
'MsgBox "2"
'              ReDim lDummyExcelWorkbook(1& To 1&) As Excel_Workbook At lDummyExcelWorkbookPtr
'MsgBox "3"
'              Object Call lExcel_Application(1&).Workbooks.Add To lDummyExcelWorkbook(1&)
'MsgBox "4"
'              Function = Str$(@lDummyExcelWorkbookPtr)
'MsgBox "5"

              Object Get lExcel_Workbook(1&).Activesheet To lDummyExcelWorksheet'vDummyVariant
              'Set lDummyExcelWorksheet = vDummyVariant
'Object Get lDummyExcelWorksheet.Name To vDummyVariant
'MsgBox Str$(ObjPtr(lDummyExcelWorksheet))'lDummyExcelWorksheet.Name

              Function = Str$(ObjPtr(lDummyExcelWorksheet))

'MsgBox FuncName$ & " end"
              
          '  '---SET action happen usually when property is used directly as first token in a source code line
          '  Case %Class_Action_Set 
          '    '---Here programmer has to parse equal (=) sign and than a numeric or string expression
          '    '---But remember: equal sign can also be += &= -= *= /= \=
          '    lEqualType = thinBasic_CheckEqualType_Mandatory
          '    If thinBasic_ErrorFree Then
          '      Select Case Long lEqualType
          '
          '        '---Assignment
          '        Case  %Equal_EQ
          '          vDummyVariant = thinBasic_Parse1Number
          '          Object Let lExcel_Application(1&).ActiveWindow.Close = vDummyVariant
          '
      '   '         '---Assignment with concatenation      
      '   '         Case  %Equal_PLUSEQUAL, _
      '   '               %Equal_CONCEQUAL, _
      '   '               %Equal_POINTEQUAL
      '   '
          '        Case Else
          '          '---Not supperted for strings 
          '          MsgBox FuncName$ & Str$(lEqualType) 
          '          
          '      End Select
          '    End If
          '            
          '  Case %Class_Action_None
          '    '---Should not happen but ... I will leave in any case
          'End Select
      
      '  Case ???
      
      End Select  'UCase$(sSubItem)
      
    End If
  End Function

  '--------------------------------------------------------------------------------------
  Function cExcel_Workbook_Property_Name(ByVal pObject As Long, ByVal lAction As Long) As String
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
    
    ReDim oObject(1& To 1&) As iExcel_Workbook At pObject
    ReDim lExcel_Workbook(1& To 1&) As Excel_Workbook At oObject(1&).cExcel_Workbook_GetClassObject

    If thinBasic_CheckPoint_Optional Then
      sSubItem = thinBasic_GetUnknownToken   
    End If
    
    If thinBasic_ErrorFree Then
      Select Case UCase$(sSubItem)
        Case ""                                                 
          Object Get lExcel_Workbook(1&).Name To vDummyVariant'vDummyVariant
          Function = Variant$(vDummyVariant)
      
      End Select  'UCase$(sSubItem)
      
    End If
  End Function

  '--------------------------------------------------------------------------------------
  Function cExcel_Workbook_Property_FullName(ByVal pObject As Long, ByVal lAction As Long) As String
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
    
    ReDim oObject(1& To 1&) As iExcel_Workbook At pObject
    ReDim lExcel_Workbook(1& To 1&) As Excel_Workbook At oObject(1&).cExcel_Workbook_GetClassObject

    If thinBasic_CheckPoint_Optional Then
      sSubItem = thinBasic_GetUnknownToken   
    End If
    
    If thinBasic_ErrorFree Then
      Select Case UCase$(sSubItem)
        Case ""                                                 
          Object Get lExcel_Workbook(1&).FullName To vDummyVariant'vDummyVariant
          Function = Variant$(vDummyVariant)
      
      End Select  'UCase$(sSubItem)
      
    End If
  End Function

  '--------------------------------------------------------------------------------------
  Function cExcel_Workbook_Property_Saved(ByVal pObject As Long, ByVal lAction As Long) As Long
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
    Local lSaved        As Long
    Local vDummyVariant As variant
    
    ReDim oObject(1& To 1&) As iExcel_Workbook At pObject
    ReDim lExcel_Workbook(1& To 1&) As Excel_Workbook At oObject(1&).cExcel_Workbook_GetClassObject

    Select Case Long lAction
    
      '---GET action happen usually when property is used inside an expression
      Case %Class_Action_Get                                                  
        '---Return value
        Object Get lExcel_Workbook(1&).Saved To vDummyVariant'vDummyVariant
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
              lSaved = thinBasic_Parse1Number
              Object Let lExcel_Workbook(1&).Saved = lSaved
    
            '---Assignment with concatenation      
            Case  %Equal_PLUSEQUAL, _
                  %Equal_CONCEQUAL, _
                  %Equal_POINTEQUAL
            
            Case Else
              '---Not supperted for strings 
              MsgBox FuncName$ & Str$(lEqualType) 
              
          End Select
        End If
                
      Case %Class_Action_None
        '---Should not happen but ... I will leave in any case
    End Select

  End Function
