  '----------------------------------------------------------------------------
  ' Class cExcel_Worksheet implements ...
  '----------------------------------------------------------------------------
  Class cExcel_Worksheet

    '----------------------------------------------------------------------------
    '---Class Private variables
    '----------------------------------------------------------------------------
    Instance pExcel_Worksheet As Excel_Worksheet
    
    '----------------------------------------------------------------------------
    ' We need to define an interface that will contain public methods
    '----------------------------------------------------------------------------
    Interface iExcel_Worksheet

      Inherit IUnknown

      '----------------------------------------------------------------------------
      ' Creator
      '----------------------------------------------------------------------------
      Method cExcel_Worksheet_Create(ByVal ptrWorkbook As Dword) As Dword
        ReDim oWorkbook(1& To 1&) As Excel_Workbook At ptrWorkbook 
        Object Call oWorkbook(1&).Worksheets.Add To pExcel_Worksheet

        If IsNothing(pExcel_Worksheet) Then Exit Method
      End Method

      '----------------------------------------------------------------------------
      ' Creator Direct
      '----------------------------------------------------------------------------
      Method cExcel_Worksheet_Create_Direct(ByVal ptrToDirectSheet As Dword) As Dword
        '---Here is the magic: we store at the memory address of pExcel_Worksheet (that is a Excel Worksheet COM pointer)
        '---The address passed by the script. The address is usually returned by another Class method or property
        '---Instructed to return an OBJPTR to an Excel Worksheet
        Poke Dword, VarPtr(pExcel_Worksheet), ptrToDirectSheet
      End Method

      '----------------------------------------------------------------------------
      ' Destructor
      '----------------------------------------------------------------------------
      Method cExcel_Worksheet_Destroy() As Dword 
        Set pExcel_Worksheet = Nothing
      End Method

      '----------------------------------------------------------------------------
      ' GetClassObject
      '----------------------------------------------------------------------------
      Method cExcel_Worksheet_GetClassObject() As Dword 
        Method = VarPtr(pExcel_Worksheet)
      End Method


    End Interface

  End Class
  '----------------------------------------------------------------------------
  '----------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  Sub cExcel_Worksheet_Create( ByVal pObject As Long )
  '----------------------------------------------------------------------------
'    Local sConnect  As String
    Local pp          As Byte
    Local pWorkbook   As Ext

    pp = thinBasic_CheckOpenParens_Optional
      thinBasic_ParseNumber pWorkbook
    If pp Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then

      'Here is the magic !!!!!!!!!!!!
      'Use the magic REDIM ... AT and the trick is done
      'We are defining a new class interface variable at the pointer location allocated by thinBasic and passed to all methods
      ReDim oObject(1& To 1&) As iExcel_Worksheet At pObject

      '---We instantiate a new class
      oObject(1&) = Class "cExcel_Worksheet"
      '---We call sSet in order to assign initial value
      oObject(1&).cExcel_Worksheet_Create(pWorkbook)

    'else
      '---emit some error
    End If
  End Sub

  '----------------------------------------------------------------------------
  Sub cExcel_Worksheet_Create_Direct( ByVal pObject As Long, ByVal pDirectClass As Dword)
  '----------------------------------------------------------------------------
    Local pp          As Byte
    Local pWorkbook   As Ext

    '---If all went fine ...
    If thinBasic_ErrorFree Then

      'Here is the magic !!!!!!!!!!!!
      'Use the magic REDIM ... AT and the trick is done
      'We are defining a new class interface variable at the pointer location allocated by thinBasic and passed to all methods
      ReDim oObject(1& To 1&) As iExcel_Worksheet At pObject

      '---We instantiate a new class
      oObject(1&) = Class "cExcel_Worksheet"
      '---We call sSet in order to assign initial value
      oObject(1&).cExcel_Worksheet_Create_Direct(pDirectClass)

    'else
      '---emit some error
    End If
  End Sub

  '----------------------------------------------------------------------------
  Sub cExcel_Worksheet_Destroy(ByVal pObject As Long)
  '----------------------------------------------------------------------------
    'If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Worksheet At pObject
      oObject(1&).cExcel_Worksheet_Destroy()
    End If
  End Sub

  '----------------------------------------------------------------------------
  Function cExcel_Worksheet_GetClassObject(ByVal pObject As Long) As Dword
  '----------------------------------------------------------------------------
    'If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Worksheet At pObject
      Function = oObject(1&).cExcel_Worksheet_GetClassObject()
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Worksheet_Method_PrintPreview(ByVal pObject As Long) As Ext
  '----------------------------------------------------------------------------
    If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Worksheet At pObject
      ReDim lExcel_Worksheet(1& To 1&) As Excel_Worksheet At oObject(1&).cExcel_Worksheet_GetClassObject

      Object Call lExcel_Worksheet(1&).PrintPreview
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Worksheet_Method_Activate(ByVal pObject As Long) As Ext
  '----------------------------------------------------------------------------
    If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Worksheet At pObject
      ReDim lExcel_Worksheet(1& To 1&) As Excel_Worksheet At oObject(1&).cExcel_Worksheet_GetClassObject

      Object Call lExcel_Worksheet(1&).Activate
    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Worksheet_Property_Cells(ByVal pObject As Long, ByVal lAction As Long) As String
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
    Local eX, eY      As Ext
    Local vX, vY      As variant
    Local sText       As String
    Local eText       As Ext
    Local vText       As variant
    Local lEqualType  As Long
    Local sSubItem    As String
    
    ReDim oObject(1& To 1&) As iExcel_Worksheet At pObject
    ReDim lExcel_Worksheet(1& To 1&) As Excel_Worksheet At oObject(1&).cExcel_Worksheet_GetClassObject

    thinBasic_CheckOpenParens_Mandatory
    thinBasic_ParseNumber eY
    If thinBasic_CheckComma_Mandatory Then
      thinBasic_ParseNumber eX
    End If
    thinBasic_CheckCloseParens_Mandatory
    
    If thinBasic_CheckPoint_Optional Then
      sSubItem = thinBasic_GetUnknownToken   
    End If
    
    vX = eX  
    vY = eY
    
    If thinBasic_ErrorFree Then
      Select Case UCase$(sSubItem)
        Case "VALUE", ""
          Select Case Long lAction
          
            '---GET action happen usually when property is used inside an expression
            Case %Class_Action_Get                                                  
              '---Return value
              Object Get lExcel_Worksheet(1&).Cells(vY, vX).Value To vText
              Select Case VariantVT(vText)
                Case %VT_BStr, %VT_LPStr, %VT_LPWStr
                  Function = Variant$(vText)
                Case %VT_Date
                  Function = Variant$(vText)
                Case Else
                  Function = Format$(Variant#(vText))
              End Select
                            
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
                        thinBasic_ParseNumber eText
                        Let vText = eText As Double
                      Case %thinBasic_ReturnString
                        thinBasic_ParseString sText
                        vText = sText
                    End Select
      
                    Object Let lExcel_Worksheet(1&).Cells(vY, vX).Value = vText
      
                  Case Else
                    '---Not supperted for strings 
                    MsgBox FuncName$ & Str$(lEqualType)
                    
                End Select
              End If
                      
            Case %Class_Action_None
              '---Should not happen but ... I will leave in any case
          End Select
          
      End Select

    End If
  End Function

  '----------------------------------------------------------------------------
  Function cExcel_Worksheet_Property_Name(ByVal pObject As Long, ByVal lAction As Long) As String
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
    Local sName       As String
    Local vName       As variant
    Local lEqualType  As Long
    
    ReDim oObject(1& To 1&) As iExcel_Worksheet At pObject
    ReDim lExcel_Worksheet(1& To 1&) As Excel_Worksheet At oObject(1&).cExcel_Worksheet_GetClassObject

    If thinBasic_ErrorFree Then

          Select Case Long lAction
          
            '---GET action happen usually when property is used inside an expression
            Case %Class_Action_Get                                                  
              '---Return value
              Object Get lExcel_Worksheet(1&).Name To vName
              Function = Variant$(vName)
              
            '---SET action happen usually when property is used directly as first token in a source code line
            Case %Class_Action_Set 
              '---Here programmer has to parse equal (=) sign and than a numeric or string expression
              '---But remember: equal sign can also be += &= -= *= /= \=
              lEqualType = thinBasic_CheckEqualType_Mandatory
              If thinBasic_ErrorFree Then
                Select Case Long lEqualType
      
                  '---Assignment
                  Case  %Equal_EQ 
                  
                    thinBasic_ParseString sName
                    vName = sName
                    Object Let lExcel_Worksheet(1&).Name = vName
      
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
  Function cExcel_Worksheet_Property_Range(ByVal pObject As Long, ByVal lAction As Long) As String
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
    
    ReDim oObject(1& To 1&) As iExcel_Worksheet At pObject
    ReDim lExcel_Worksheet(1& To 1&) As Excel_Worksheet At oObject(1&).cExcel_Worksheet_GetClassObject

    thinBasic_CheckOpenParens_Mandatory
    thinBasic_ParseString sRange
    Let vRange = sRange
    'thinBasic_ParseNumber eY
    'If thinBasic_CheckComma_Mandatory Then
    '  thinBasic_ParseNumber eX
    'End If
    thinBasic_CheckCloseParens_Mandatory
    
    If thinBasic_CheckPoint_Optional Then
      sSubItem = thinBasic_GetUnknownToken   
    End If
    
    If thinBasic_ErrorFree Then
      Select Case UCase$(sSubItem)
        Case ""
          'Object Get lExcel_Application(1&).ActiveWorkbook To lDummyExcelWorkbook
          'Function = Str$(ObjPtr(lDummyExcelWorkbook))
          Object Get lExcel_Worksheet(1&).Range(vRange) To lDummyExcelRange
          Function = Str$(ObjPtr(lDummyExcelRange))
'MsgBox Str$(ObjPtr(lDummyExcelRange)) & variant$(vRange)
        
        Case "SELECT"
          Object Call lExcel_Worksheet(1&).Range(vRange).Select
          'Function = Str$(Variant#(vResult))

        Case "CLEAR"
          Object Call lExcel_Worksheet(1&).Range(vRange).Clear

        Case "VALUE"
          Select Case Long lAction
          
            '---GET action happen usually when property is used inside an expression
            Case %Class_Action_Get                                                  
              '---Return value
              Object Get lExcel_Worksheet(1&).Range(vRange).Value To vText
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

                    Select Case thinBasic_DetermineType
                      Case %thinBasic_ReturnNumber
                        thinBasic_ParseNumber eText
                        Let vText = eText As Double
                      Case %thinBasic_ReturnString
                        thinBasic_ParseString sText
                        vText = sText
                    End Select
                    
                    Object Let lExcel_Worksheet(1&).Range(vRange).Value = vText
             
                  Case Else
                    '---Not supperted for strings 
                    MsgBox FuncName$ & Str$(lEqualType)
                    
                End Select
              End If
                      
            Case %Class_Action_None
              '---Should not happen but ... I will leave in any case
          End Select

        Case "FORMULA"
          Select Case Long lAction
          
            '---GET action happen usually when property is used inside an expression
            Case %Class_Action_Get                                                  
              '---Return value
              Object Get lExcel_Worksheet(1&).Range(vRange).Formula To vText
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

                    Select Case thinBasic_DetermineType
                      Case %thinBasic_ReturnNumber
                        thinBasic_ParseNumber eText
                        Let vText = eText As Double
                      Case %thinBasic_ReturnString
                        thinBasic_ParseString sText
                        vText = sText
                    End Select
                    
                    Object Let lExcel_Worksheet(1&).Range(vRange).Formula = vText
             
                  Case Else
                    '---Not supperted for strings 
                    MsgBox FuncName$ & Str$(lEqualType)
                    
                End Select
              End If
                      
            Case %Class_Action_None
              '---Should not happen but ... I will leave in any case
          End Select
          
      End Select

    End If
  End Function
