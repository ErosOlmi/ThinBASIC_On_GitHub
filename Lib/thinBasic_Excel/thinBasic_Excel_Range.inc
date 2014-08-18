  '----------------------------------------------------------------------------
  ' Class cExcel_Range implements ...
  '----------------------------------------------------------------------------
  Class cExcel_Range

    '----------------------------------------------------------------------------
    '---Class Private variables
    '----------------------------------------------------------------------------
    Instance pExcel_Range As Range
    
    '----------------------------------------------------------------------------
    ' We need to define an interface that will contain public methods
    '----------------------------------------------------------------------------
    Interface iExcel_Range

      Inherit IUnknown

      '----------------------------------------------------------------------------
      ' Creator
      '----------------------------------------------------------------------------
      Method cExcel_Range_Create(ByVal ptrSheet As Dword) As Dword
'MsgBox FuncName$
'        ReDim oRange(1& To 1&) As Range At ptrWorkbook 
'        Object Call oWorkbook(1&).Worksheets.Add To pExcel_Worksheet
'
'        If IsNothing(pExcel_Worksheet) Then Exit Method
      End Method

      '----------------------------------------------------------------------------
      ' Creator Direct
      '----------------------------------------------------------------------------
      Method cExcel_Range_Create_Direct(ByVal ptrToDirectRange As Dword) As Dword
        '---Here is the magic: we store at the memory address of pExcel_Worksheet (that is a Excel Worksheet COM pointer)
        '---The address passed by the script. The address is usually returned by another Class method or property
        '---Instructed to return an OBJPTR to an Excel Worksheet
        Poke Dword, VarPtr(pExcel_Range), ptrToDirectRange
      End Method

      '----------------------------------------------------------------------------
      ' Destructor
      '----------------------------------------------------------------------------
      Method cExcel_Range_Destroy() As Dword 
        Set pExcel_Range = Nothing
      End Method

      '----------------------------------------------------------------------------
      ' GetClassObject
      '----------------------------------------------------------------------------
      Method cExcel_Range_GetClassObject() As Dword 
        Method = VarPtr(pExcel_Range)
      End Method

    End Interface

  End Class
  '----------------------------------------------------------------------------
  '----------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  Sub cExcel_Range_Create( ByVal pObject As Long )
  '----------------------------------------------------------------------------
'MsgBox FuncName$
'    Local pp          As Byte
'    Local pWorkbook   As Ext
'
'    pp = thinBasic_CheckOpenParens_Optional
'      thinBasic_ParseNumber pWorkbook
'    If pp Then thinBasic_CheckCloseParens_Mandatory
'
'    '---If all went fine ...
'    If thinBasic_ErrorFree Then
'
'      'Here is the magic !!!!!!!!!!!!
'      'Use the magic REDIM ... AT and the trick is done
'      'We are defining a new class interface variable at the pointer location allocated by thinBasic and passed to all methods
'      ReDim oObject(1& To 1&) As iExcel_Range At pObject
'
'      '---We instantiate a new class
'      oObject(1&) = Class "cExcel_Range"
'      '---We call sSet in order to assign initial value
'      oObject(1&).cExcel_Range_Create(pWorkbook)
'
'    'else
'      '---emit some error
'    End If
  End Sub

  '----------------------------------------------------------------------------
  Sub cExcel_Range_Create_Direct( ByVal pObject As Long, ByVal pDirectClass As Dword)
  '----------------------------------------------------------------------------
'MsgBox FuncName$
    Local pp          As Byte
    Local pWorkbook   As Ext

    '---If all went fine ...
    If thinBasic_ErrorFree Then

      'Here is the magic !!!!!!!!!!!!
      'Use the magic REDIM ... AT and the trick is done
      'We are defining a new class interface variable at the pointer location allocated by thinBasic and passed to all methods
      ReDim oObject(1& To 1&) As iExcel_Range At pObject

    '---We instantiate a new class
      If IsObject(oObject(1&)) Then Set oObject(1&) = Nothing
      oObject(1&) = Class "cExcel_Range"
      '---We call sSet in order to assign initial value
      oObject(1&).cExcel_Range_Create_Direct(pDirectClass)

    'else
      '---emit some error
    End If
  End Sub

  '----------------------------------------------------------------------------
  Sub cExcel_Range_Destroy(ByVal pObject As Long)
  '----------------------------------------------------------------------------
    'If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

'MsgBox "Start " & FuncName$
    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Range At pObject
      oObject(1&).cExcel_Range_Destroy()
    End If
'MsgBox "End " & FuncName$
  End Sub

  '----------------------------------------------------------------------------
  Function cExcel_Range_GetClassObject(ByVal pObject As Long) As Ext
  '----------------------------------------------------------------------------
    'If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Range At pObject
      Function = oObject(1&).cExcel_Range_GetClassObject()
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Range_Method_Select(ByVal pObject As Long) As Ext
  '----------------------------------------------------------------------------
    'If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Range At pObject
      ReDim lExcel_Range(1& To 1&) As Range At oObject(1&).cExcel_Range_GetClassObject

      Object Call lExcel_Range(1&).Select
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Range_Method_Clear(ByVal pObject As Long) As Ext
  '----------------------------------------------------------------------------
    'If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Range At pObject
      ReDim lExcel_Range(1& To 1&) As Range At oObject(1&).cExcel_Range_GetClassObject

      Object Call lExcel_Range(1&).Clear
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Range_Method_ClearContents(ByVal pObject As Long) As Ext
  '----------------------------------------------------------------------------
    'If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Range At pObject
      ReDim lExcel_Range(1& To 1&) As Range At oObject(1&).cExcel_Range_GetClassObject

      Object Call lExcel_Range(1&).ClearContents
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Range_Method_ClearFormats(ByVal pObject As Long) As Ext
  '----------------------------------------------------------------------------
    'If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Range At pObject
      ReDim lExcel_Range(1& To 1&) As Range At oObject(1&).cExcel_Range_GetClassObject

      Object Call lExcel_Range(1&).ClearFormats
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Range_Method_ClearComments(ByVal pObject As Long) As Ext
  '----------------------------------------------------------------------------
    'If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Range At pObject
      ReDim lExcel_Range(1& To 1&) As Range At oObject(1&).cExcel_Range_GetClassObject

      Object Call lExcel_Range(1&).ClearComments
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Range_Method_ClearNotes(ByVal pObject As Long) As Ext
  '----------------------------------------------------------------------------
    'If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Range At pObject
      ReDim lExcel_Range(1& To 1&) As Range At oObject(1&).cExcel_Range_GetClassObject

      Object Call lExcel_Range(1&).ClearNotes
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Range_Property_Value(ByVal pObject As Long, ByVal lAction As Long) As String
  '----------------------------------------------------------------------------
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
    Local lEqualType  As Long
    Local eValue      As Ext
    Local sValue      As String
    Local vValue      As variant
    
    ReDim oObject(1& To 1&) As iExcel_Range At pObject
    ReDim lExcel_Range(1& To 1&) As Range At oObject(1&).cExcel_Range_GetClassObject

    If thinBasic_ErrorFree Then
      Select Case Long lAction
      
        '---GET action happen usually when property is used inside an expression
        Case %Class_Action_Get                                                  
          '---Return value
          Object Get lExcel_Range(1&).Value To vValue
          Function = Variant$(vValue)
          
        '---SET action happen usually when property is used directly as first token in a source code line
        Case %Class_Action_Set 
          '---Here programmer has to parse equal (=) sign and than a numeric or string expression
          '---But remember: equal sign can also be += &= -= *= /= \=
          lEqualType = thinBasic_CheckEqualType_Mandatory
          If thinBasic_ErrorFree Then
            Select Case Long lEqualType
  
              '---Assignment
              Case  %Equal_EQ
                Select Case thinBasic_DetermineType
                  Case %thinBasic_ReturnNumber
                    thinBasic_ParseNumber eValue
                    Let vValue = eValue As Double
                  Case %thinBasic_ReturnString
                    thinBasic_ParseString sValue
                    vValue = sValue
                End Select
  
                Object Let lExcel_Range(1&).Value = vValue
  
              Case Else
                '---Not supperted for strings 
                MsgBox FuncName$ & Str$(lEqualType)
                
            End Select
          End If
                  
        Case %Class_Action_None
          '---Should not happen but ... I will leave in any case
      End Select
          
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Range_Property_Address(ByVal pObject As Long, ByVal lAction As Long) As String
  '----------------------------------------------------------------------------
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
    Local lEqualType  As Long
    Local eValue      As Ext
    Local sValue      As String
    Local vValue      As variant
    
    ReDim oObject(1& To 1&) As iExcel_Range At pObject
    ReDim lExcel_Range(1& To 1&) As Range At oObject(1&).cExcel_Range_GetClassObject

'    thinBasic_CheckOpenParens_Mandatory
'    thinBasic_ParseNumber eY
'    If thinBasic_CheckComma_Mandatory Then
'      thinBasic_ParseNumber eX
'    End If
'    thinBasic_CheckCloseParens_Mandatory
'    
'    If thinBasic_CheckPoint_Optional Then
'      sSubItem = thinBasic_GetUnknownToken   
'    End If
'    
'    vX = eX  
'    vY = eY
    
    If thinBasic_ErrorFree Then
'      Select Case UCase$(sSubItem)
'        Case "VALUE", ""
          Select Case Long lAction
          
            '---GET action happen usually when property is used inside an expression
            Case %Class_Action_Get                                                  
              '---Return value
              Object Get lExcel_Range(1&).Address To vValue
              Function = Variant$(vValue)
              
            '---SET action happen usually when property is used directly as first token in a source code line
            Case %Class_Action_Set 
              '---Here programmer has to parse equal (=) sign and than a numeric or string expression
              '---But remember: equal sign can also be += &= -= *= /= \=
              lEqualType = thinBasic_CheckEqualType_Mandatory
              If thinBasic_ErrorFree Then
                Select Case Long lEqualType
      
                  '---Assignment
                  Case  %Equal_EQ
                    Select Case thinBasic_DetermineType
                      Case %thinBasic_ReturnNumber
                        thinBasic_ParseNumber eValue
                        Let vValue = eValue As Double
                      Case %thinBasic_ReturnString
                        thinBasic_ParseString sValue
                        vValue = sValue
                    End Select
      
                    Object Let lExcel_Range(1&).Address = vValue
      
                  Case Else
                    '---Not supperted for strings 
                    MsgBox FuncName$ & Str$(lEqualType)
                    
                End Select
              End If
                      
            Case %Class_Action_None
              '---Should not happen but ... I will leave in any case
          End Select
          
'      End Select
'
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Range_Property_Formula(ByVal pObject As Long, ByVal lAction As Long) As String
  '----------------------------------------------------------------------------
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
    Local lEqualType  As Long
    Local eValue      As Ext
    Local sValue      As String
    Local vValue      As variant
    
    ReDim oObject(1& To 1&) As iExcel_Range At pObject
    ReDim lExcel_Range(1& To 1&) As Range At oObject(1&).cExcel_Range_GetClassObject

'    thinBasic_CheckOpenParens_Mandatory
'    thinBasic_ParseNumber eY
'    If thinBasic_CheckComma_Mandatory Then
'      thinBasic_ParseNumber eX
'    End If
'    thinBasic_CheckCloseParens_Mandatory
'    
'    If thinBasic_CheckPoint_Optional Then
'      sSubItem = thinBasic_GetUnknownToken   
'    End If
'    
'    vX = eX  
'    vY = eY
    
    If thinBasic_ErrorFree Then
'      Select Case UCase$(sSubItem)
'        Case "VALUE", ""
          Select Case Long lAction
          
            '---GET action happen usually when property is used inside an expression
            Case %Class_Action_Get                                                  
              '---Return value
              Object Get lExcel_Range(1&).Formula To vValue
              Function = Variant$(vValue)
              
            '---SET action happen usually when property is used directly as first token in a source code line
            Case %Class_Action_Set 
              '---Here programmer has to parse equal (=) sign and than a numeric or string expression
              '---But remember: equal sign can also be += &= -= *= /= \=
              lEqualType = thinBasic_CheckEqualType_Mandatory
              If thinBasic_ErrorFree Then
                Select Case Long lEqualType
      
                  '---Assignment
                  Case  %Equal_EQ
                    Select Case thinBasic_DetermineType
                      Case %thinBasic_ReturnNumber
                        thinBasic_ParseNumber eValue
                        Let vValue = eValue As Double
                      Case %thinBasic_ReturnString
                        thinBasic_ParseString sValue
                        vValue = sValue
                    End Select
      
                    Object Let lExcel_Range(1&).Formula = vValue
      
                  Case Else
                    '---Not supperted for strings 
                    MsgBox FuncName$ & Str$(lEqualType)
                    
                End Select
              End If
                      
            Case %Class_Action_None
              '---Should not happen but ... I will leave in any case
          End Select
          
'      End Select
'
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Range_Property_HorizontalAlignment(ByVal pObject As Long, ByVal lAction As Long) As Long
  '----------------------------------------------------------------------------
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
    Local lEqualType  As Long
    Local eValue      As Ext
    Local sValue      As String
    Local vValue      As variant
    
    ReDim oObject(1& To 1&) As iExcel_Range At pObject
    ReDim lExcel_Range(1& To 1&) As Range At oObject(1&).cExcel_Range_GetClassObject

'    thinBasic_CheckOpenParens_Mandatory
'    thinBasic_ParseNumber eY
'    If thinBasic_CheckComma_Mandatory Then
'      thinBasic_ParseNumber eX
'    End If
'    thinBasic_CheckCloseParens_Mandatory
'    
'    If thinBasic_CheckPoint_Optional Then
'      sSubItem = thinBasic_GetUnknownToken   
'    End If
'    
'    vX = eX  
'    vY = eY
    
    If thinBasic_ErrorFree Then
'      Select Case UCase$(sSubItem)
'        Case "VALUE", ""
          Select Case Long lAction
          
            '---GET action happen usually when property is used inside an expression
            Case %Class_Action_Get                                                  
              '---Return value
              Object Get lExcel_Range(1&).HorizontalAlignment To vValue
              Function = Variant#(vValue)
              
            '---SET action happen usually when property is used directly as first token in a source code line
            Case %Class_Action_Set 
              '---Here programmer has to parse equal (=) sign and than a numeric or string expression
              '---But remember: equal sign can also be += &= -= *= /= \=
              lEqualType = thinBasic_CheckEqualType_Mandatory
              If thinBasic_ErrorFree Then
                Select Case Long lEqualType
      
                  '---Assignment
                  Case  %Equal_EQ
                    'Select Case thinBasic_DetermineType
                    '  Case %thinBasic_ReturnNumber
                        thinBasic_ParseNumber eValue
                        Let vValue = eValue As Double
                    '  Case %thinBasic_ReturnString
                    '    thinBasic_ParseString sValue
                    '    vValue = sValue
                    'End Select
      
                    Object Let lExcel_Range(1&).HorizontalAlignment = vValue
      
                  Case Else
                    '---Not supperted for strings 
                    MsgBox FuncName$ & Str$(lEqualType)
                    
                End Select
              End If
                      
            Case %Class_Action_None
              '---Should not happen but ... I will leave in any case
          End Select
          
'      End Select
'
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Range_Property_ColumnWidth(ByVal pObject As Long, ByVal lAction As Long) As Long
  '----------------------------------------------------------------------------
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
    Local lEqualType  As Long
    Local eValue      As Ext
    Local sValue      As String
    Local vValue      As variant
    
    ReDim oObject(1& To 1&) As iExcel_Range At pObject
    ReDim lExcel_Range(1& To 1&) As Range At oObject(1&).cExcel_Range_GetClassObject

'    thinBasic_CheckOpenParens_Mandatory
'    thinBasic_ParseNumber eY
'    If thinBasic_CheckComma_Mandatory Then
'      thinBasic_ParseNumber eX
'    End If
'    thinBasic_CheckCloseParens_Mandatory
'    
'    If thinBasic_CheckPoint_Optional Then
'      sSubItem = thinBasic_GetUnknownToken   
'    End If
'    
'    vX = eX  
'    vY = eY
    
    If thinBasic_ErrorFree Then
'      Select Case UCase$(sSubItem)
'        Case "VALUE", ""
          Select Case Long lAction
          
            '---GET action happen usually when property is used inside an expression
            Case %Class_Action_Get                                                  
              '---Return value
              Object Get lExcel_Range(1&).ColumnWidth To vValue
              Function = Variant#(vValue)
              
            '---SET action happen usually when property is used directly as first token in a source code line
            Case %Class_Action_Set 
              '---Here programmer has to parse equal (=) sign and than a numeric or string expression
              '---But remember: equal sign can also be += &= -= *= /= \=
              lEqualType = thinBasic_CheckEqualType_Mandatory
              If thinBasic_ErrorFree Then
                Select Case Long lEqualType
      
                  '---Assignment
                  Case  %Equal_EQ
                    'Select Case thinBasic_DetermineType
                    '  Case %thinBasic_ReturnNumber
                        thinBasic_ParseNumber eValue
                        Let vValue = eValue As Double
                    '  Case %thinBasic_ReturnString
                    '    thinBasic_ParseString sValue
                    '    vValue = sValue
                    'End Select
      
                    Object Let lExcel_Range(1&).ColumnWidth = vValue
      
                  Case Else
                    '---Not supperted for strings 
                    MsgBox FuncName$ & Str$(lEqualType)
                    
                End Select
              End If
                      
            Case %Class_Action_None
              '---Should not happen but ... I will leave in any case
          End Select
          
'      End Select
'
    End If
  End Function
  
  
  '----------------------------------------------------------------------------
  Function cExcel_Range_Property_Interior(ByVal pObject As Long, ByVal lAction As Long) As String
  '----------------------------------------------------------------------------
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
    Local sText       As String
    Local eText       As Ext
    Local vText       As variant
                                
    Local vResult     As variant
    
    Local lEqualType  As Long
    Local sSubItem    As String 
    
    Local sRange      As String
    Local vRange      As variant

    Local ArrayPresent    As Long
    Local ArrayName       As String
    Local ArrayElements   As Long
    Local ArrayPtr        As Long

    Local pVar            As Long
    Local pDirect         As Long
    Local lMainType       As Long
    Local lSubType        As Long
    Local lIsArray        As Long 
    Local Counter         As Long
    Local lSize           As Long

    Static lDummyExcelRange  As Range
    
    ReDim oObject(1& To 1&) As iExcel_Range At pObject
    ReDim lExcel_Range(1& To 1&) As Range At oObject(1&).cExcel_Range_GetClassObject

    'thinBasic_CheckOpenParens_Mandatory
    'thinBasic_ParseString sRange
    'Let vRange = sRange
    'thinBasic_ParseNumber eY
    'If thinBasic_CheckComma_Mandatory Then
    '  thinBasic_ParseNumber eX
    'End If
    'thinBasic_CheckCloseParens_Mandatory
    
    If thinBasic_CheckPoint_Optional Then
      sSubItem = thinBasic_GetUnknownToken   
    End If
    
    If thinBasic_ErrorFree Then
      Select Case UCase$(sSubItem)
'        Case ""
'          'Object Get lExcel_Application(1&).ActiveWorkbook To lDummyExcelWorkbook
'          'Function = Str$(ObjPtr(lDummyExcelWorkbook))
'          Object Get lExcel_Worksheet(1&).Range(vRange) To lDummyExcelRange
'          Function = Str$(ObjPtr(lDummyExcelRange))
'        
'        Case "COLORINDEX"
'          Object Call lExcel_Worksheet(1&).Range(vRange).Select
'          'Function = Str$(Variant#(vResult))
'
'        Case "CLEAR"
'          Object Call lExcel_Worksheet(1&).Range(vRange).Clear

        Case "COLORINDEX"
          Select Case Long lAction
          
            '---GET action happen usually when property is used inside an expression
            Case %Class_Action_Get                                                  
              '---Return value
              Object Get lExcel_Range(1&).Interior.ColorIndex To vText
              Function = Variant$(vText)
              
            '---SET action happen usually when property is used directly as first token in a source code line
            Case %Class_Action_Set 
              '---Here programmer has to parse equal (=) sign and than a numeric or string expression
              '---But remember: equal sign can also be += &= -= *= /= \=
              lEqualType = thinBasic_CheckEqualType_Mandatory
              If thinBasic_ErrorFree Then
                Select Case Long lEqualType
      
                  '---Assignment
                  Case %Equal_EQ

'                    Select Case thinBasic_DetermineType
'                      Case %thinBasic_ReturnNumber
                        thinBasic_ParseNumber eText
                        Let vText = eText As Double
'                      Case %thinBasic_ReturnString
'                        thinBasic_ParseString sText
'                        vText = sText
'                    End Select
                    
                    Object Let lExcel_Range(1&).Interior.ColorIndex = vText
             
                  Case Else
                    '---Not supperted for strings 
                    MsgBox FuncName$ & Str$(lEqualType)
                    
                End Select
              End If
                      
            Case %Class_Action_None
              '---Should not happen but ... I will leave in any case
          End Select

        Case "COLOR"
          Select Case Long lAction
          
            '---GET action happen usually when property is used inside an expression
            Case %Class_Action_Get                                                  
              '---Return value
              Object Get lExcel_Range(1&).Interior.Color To vText
              Function = Variant$(vText)
              
            '---SET action happen usually when property is used directly as first token in a source code line
            Case %Class_Action_Set 
              '---Here programmer has to parse equal (=) sign and than a numeric or string expression
              '---But remember: equal sign can also be += &= -= *= /= \=
              lEqualType = thinBasic_CheckEqualType_Mandatory
              If thinBasic_ErrorFree Then
                Select Case Long lEqualType
      
                  '---Assignment
                  Case %Equal_EQ

'                    Select Case thinBasic_DetermineType
'                      Case %thinBasic_ReturnNumber
                        thinBasic_ParseNumber eText
                        Let vText = eText As Double
'                      Case %thinBasic_ReturnString
'                        thinBasic_ParseString sText
'                        vText = sText
'                    End Select
                    
                    Object Let lExcel_Range(1&).Interior.Color = vText
             
                  Case Else
                    '---Not supperted for strings 
                    MsgBox FuncName$ & Str$(lEqualType)
                    
                End Select
              End If
                      
            Case %Class_Action_None
              '---Should not happen but ... I will leave in any case
          End Select

'        Case "FORMULA"
'          Select Case Long lAction
'          
'            '---GET action happen usually when property is used inside an expression
'            Case %Class_Action_Get                                                  
'              '---Return value
'              Object Get lExcel_Worksheet(1&).Range(vRange).Formula To vText
'              Function = Variant$(vText)
'              
'            '---SET action happen usually when property is used directly as first token in a source code line
'            Case %Class_Action_Set 
'              '---Here programmer has to parse equal (=) sign and than a numeric or string expression
'              '---But remember: equal sign can also be += &= -= *= /= \=
'              lEqualType = thinBasic_CheckEqualType_Mandatory
'              If thinBasic_ErrorFree Then
'                Select Case Long lEqualType
'      
'                  '---Assignment
'                  Case %Equal_EQ
'
'                    Select Case thinBasic_DetermineType
'                      Case %thinBasic_ReturnNumber
'                        thinBasic_ParseNumber eText
'                        Let vText = eText As Double
'                      Case %thinBasic_ReturnString
'                        thinBasic_ParseString sText
'                        vText = sText
'                    End Select
'                    
'                    Object Let lExcel_Worksheet(1&).Range(vRange).Formula = vText
'             
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
          
      End Select

    End If
  End Function
