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
'    Local sConnect  As String
    Local pp          As Byte
    Local pExcelApp   As Ext

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
  Function cExcel_Workbook_GetClassObject(ByVal pObject As Long) As Ext
  '----------------------------------------------------------------------------
    'If thinBasic_CheckOpenParens_Optional Then thinBasic_CheckCloseParens_Mandatory

    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Workbook At pObject
      Function = oObject(1&).cExcel_Workbook_GetClassObject()
    End If
  End Function


  '----------------------------------------------------------------------------
  Function cExcel_Workbook_Method_SaveAs(ByVal pObject As Long) As Ext
  '----------------------------------------------------------------------------
    Dim pp        As Byte
    Dim sFileName As String
    Dim vFileName As variant
    
    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iExcel_Workbook At pObject
      ReDim lExcel_Workbook(1& To 1&) As Excel_Workbook At oObject(1&).cExcel_Workbook_GetClassObject

      pp = thinBasic_CheckOpenParens_Optional
      thinBasic_ParseString sFileName
      'If thinBasic_CheckComma_Mandatory Then
      '  thinBasic_ParseNumber eX
      'End If
      If pp Then thinBasic_CheckCloseParens_Mandatory
      
      If thinBasic_ErrorFree Then
        vFileName = sFileName
        Object Call lExcel_Workbook(1&).SaveAs(vFileName)

      End If
      
    End If
  End Function
