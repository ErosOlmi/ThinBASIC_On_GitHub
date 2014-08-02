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
  Function cExcel_Worksheet_GetClassObject(ByVal pObject As Long) As Ext
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
    'If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Worksheet At pObject
      ReDim lExcel_Worksheet(1& To 1&) As Excel_Worksheet At oObject(1&).cExcel_Worksheet_GetClassObject

      Object Call lExcel_Worksheet(1&).PrintPreview
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
              Function = Variant$(vText)
              
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
