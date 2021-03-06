'
' Excel support for ThinBASIC
'
' Petr Schreiber/Eros Olmi
'

#COMPILE DLL
#REGISTER NONE
#DIM ALL

#RESOURCE VERSIONINFO
#RESOURCE FILEVERSION 1, 0, 0, 1
#RESOURCE PRODUCTVERSION 1, 0, 0, 1

#RESOURCE STRINGINFO "0409", "04B0"

#RESOURCE VERSION$ "CompanyName",      "ThinBASIC"
#RESOURCE VERSION$ "FileDescription",  "thinBasic module for Excel support"
#Resource VERSION$ "FileVersion",      "1.0.0.0"
#RESOURCE VERSION$ "InternalName",     "Excel"
#RESOURCE VERSION$ "OriginalFilename", "ThinBASIC_Excel.dll"
#RESOURCE VERSION$ "LegalCopyright",   "Copyright � Petr Schreiber/Eros Olmi 2014"
#RESOURCE VERSION$ "ProductName",      "Module"
#Resource VERSION$ "ProductVersion",   "1.0.0.0"
#RESOURCE VERSION$ "Comments",         "Support site: http://www.thinbasic.com/"

GLOBAL gPath AS STRING

#Include Once "WIN32API.INC"

'---Every used defined thinBasic module must include this file
#Include Once "\ThinBASIC\Lib\thinCore.inc"

#Include Once ".\Excel.inc"
#Include Once ".\thinBasic_Excel_Application.inc"
#Include Once ".\thinBasic_Excel_Workbook.inc"
#Include Once ".\thinBasic_Excel_Worksheet.inc"
#Include Once ".\thinBasic_Excel_Range.inc" 
#Include Once ".\thinBasic_Excel_Utilities.inc" 



  '----------------------------------------------------------------------------
  '---References:
  '     Excel Object Model: http://msdn.microsoft.com/en-us/library/bb149081(v=office.12).aspx
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  Function LoadLocalSymbols Alias "LoadLocalSymbols" (Optional ByVal sPath As String) Export As Long
  ' This function is automatically called by thinCore whenever this DLL is loaded.
  ' This function MUST be present in every external DLL you want to use
  ' with thinBasic
  ' Use this function to initialize every variable you need and for loading the
  ' new symbol (read Keyword) you have created.
  '----------------------------------------------------------------------------
    Local RetCode                   As Long
    LOCAL pClass_cExcel_Application AS LONG
    LOCAL pClass_cExcel_Workbook    AS LONG
    Local pClass_cExcel_Worksheet   As Long
    Local pClass_cExcel_Range       As Long

    ' -- Save DLL loading path to global var
    gPath = sPath

'msgbox str$(%LOCALE_SLIST)
    '---
    ' Excel Application Class
    '---
      pClass_cExcel_Application = thinBasic_Class_Add("Excel_Application", 0)

      ' -- If class was created, define all methods and properties, each connected to a CODEPTR module function/sub
      IF pClass_cExcel_Application THEN

        ' -- Constructor wrapper function needs to be linked in as _Create
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Application, "_Create"           , %thinBasic_ReturnNone       , CodePtr(cExcel_Application_Create           ))
        ' -- Destructor wrapper function needs to be linked in as _Destroy
        ' -- WARNING: You MUST supply destructor and set the object to NOTHING, otherwise you risk memory leak
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Application, "_Destroy"          , %thinBasic_ReturnNone       , CodePtr(cExcel_Application_Destroy          ))
        ' -- ClassObject
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Application, "_GetClassObject"   , %thinBasic_ReturnCodedWord  , CodePtr(cExcel_Application_GetClassObject   ))

        ' -- Common methods can take any name
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Application, "IsNothing"         , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Application_Method_IsNothing ))
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Application, "Quit"              , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Application_Method_Quit      ))
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Application, "SendKeys"          , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Application_Method_SendKeys  ))

        ' -- Common properties can take any name
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "Version"                 , %thinBasic_ReturnString     , CODEPTR(cExcel_Application_Property_Version ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "Visible"                 , %thinBasic_ReturnCodeLong   , CODEPTR(cExcel_Application_Property_Visible ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "AlertBeforeOverwriting"  , %thinBasic_ReturnCodeLong   , CODEPTR(cExcel_Application_Property_AlertBeforeOverwriting  ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "DisplayAlerts"           , %thinBasic_ReturnCodeLong   , CODEPTR(cExcel_Application_Property_DisplayAlerts           ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "ActiveWindow"            , %thinBasic_ReturnString     , CodePtr(cExcel_Application_Property_ActiveWindow            ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "Workbooks"               , %thinBasic_ReturnString     , CodePtr(cExcel_Application_Property_Workbooks               ))

        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "ActiveWorkbook"          , %thinBasic_ReturnString     , CodePtr(cExcel_Application_Property_ActiveWorkbook          ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "ActiveSheet"             , %thinBasic_ReturnString     , CodePtr(cExcel_Application_Property_ActiveSheet             ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "ActiveCell"              , %thinBasic_ReturnString     , CodePtr(cExcel_Application_Property_ActiveCell              ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "Selection"               , %thinBasic_ReturnString     , CodePtr(cExcel_Application_Property_Selection               ))

      END IF

    '---
    ' Excel Workbook Class
    '---
      pClass_cExcel_Workbook = thinBasic_Class_Add("Excel_Workbook", 0)

      ' -- If class was created, define all methods and properties, each connected to a CODEPTR module function/sub
      IF pClass_cExcel_Workbook THEN

        ' -- Constructor wrapper function needs to be linked in as _Create
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Workbook, "_Create"                    , %thinBasic_ReturnNone       , CodePtr(cExcel_Workbook_Create            ))
        ' -- Constructor wrapper function used for direct creation (without the use of NEW keyword) _CreateDirect
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Workbook, "_CreateDirect"              , %thinBasic_ReturnNone       , CodePtr(cExcel_Workbook_Create_Direct     ))

        ' -- Destructor wrapper function needs to be linked in as _Destroy
        ' -- WARNING: You MUST supply destructor and set the object to NOTHING, otherwise you risk memory leak
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Workbook, "_Destroy"                   , %thinBasic_ReturnNone       , CodePtr(cExcel_Workbook_Destroy           ))
        ' -- ClassObject
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Workbook, "_GetClassObject"            , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Workbook_GetClassObject    ))

        ' -- Common methods can take any name
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Workbook, "Save"                       , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Workbook_Method_Save       ))
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Workbook, "SaveAs"                     , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Workbook_Method_SaveAs     ))
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Workbook, "Activate"                   , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Workbook_Method_Activate   ))

        ' -- Common properties can take any name
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Workbook, "Worksheets"                 , %thinBasic_ReturnString     , CodePtr(cExcel_Workbook_Property_Worksheets ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Workbook, "ActiveSheet"                , %thinBasic_ReturnString     , CodePtr(cExcel_Workbook_Property_Activesheet))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Workbook, "Name"                       , %thinBasic_ReturnString     , CodePtr(cExcel_Workbook_Property_Name       ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Workbook, "FullName"                   , %thinBasic_ReturnString     , CodePtr(cExcel_Workbook_Property_FullName   ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Workbook, "Saved"                      , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Workbook_Property_Saved      ))

      End If 

    '---
    ' Excel Worksheet Class
    '---
      pClass_cExcel_Worksheet = thinBasic_Class_Add("Excel_Worksheet", 0)

      ' -- If class was created, define all methods and properties, each connected to a CODEPTR module function/sub
      IF pClass_cExcel_Worksheet THEN

        ' -- Constructor wrapper function needs to be linked in as _Create
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Worksheet, "_Create"           , %thinBasic_ReturnNone       , CodePtr(cExcel_Worksheet_Create           ))
        ' -- Constructor wrapper function needs to be linked in as _Create
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Worksheet, "_CreateDirect"     , %thinBasic_ReturnNone       , CodePtr(cExcel_Worksheet_Create_Direct    ))
        ' -- Destructor wrapper function needs to be linked in as _Destroy
        ' -- WARNING: You MUST supply destructor and set the object to NOTHING, otherwise you risk memory leak
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Worksheet, "_Destroy"          , %thinBasic_ReturnNone       , CODEPTR(cExcel_Worksheet_Destroy          ))
        ' -- ClassObject
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Worksheet, "_GetClassObject"   , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Worksheet_GetClassObject   ))

        ' -- Common methods can take any name
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Worksheet, "PrintPreview"      , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Worksheet_Method_PrintPreview     ))
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Worksheet, "Activate"          , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Worksheet_Method_Activate         ))

        ' -- Common properties can take any name
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Worksheet, "Cells"             , %thinBasic_ReturnString     , CODEPTR(cExcel_Worksheet_Property_Cells   ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Worksheet, "Name"              , %thinBasic_ReturnString     , CodePtr(cExcel_Worksheet_Property_Name    ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Worksheet, "Range"             , %thinBasic_ReturnString     , CodePtr(cExcel_Worksheet_Property_Range   ))

      END IF


    '---
    ' Excel Range Class
    '---
      pClass_cExcel_Range = thinBasic_Class_Add("Excel_Range", 0)

      ' -- If class was created, define all methods and properties, each connected to a CODEPTR module function/sub
      If pClass_cExcel_Range Then

        ' -- Constructor wrapper function needs to be linked in as _Create
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "_Create"           , %thinBasic_ReturnNone       , CodePtr(cExcel_Range_Create               ))
        ' -- Constructor wrapper function needs to be linked in as _Create
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "_CreateDirect"     , %thinBasic_ReturnNone       , CodePtr(cExcel_Range_Create_Direct        ))
        ' -- Destructor wrapper function needs to be linked in as _Destroy
        ' -- WARNING: You MUST supply destructor and set the object to NOTHING, otherwise you risk memory leak
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "_Destroy"          , %thinBasic_ReturnNone       , CodePtr(cExcel_Range_Destroy              ))
        ' -- ClassObject
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "_GetClassObject"   , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Range_GetClassObject       ))

        ' -- Common methods can take any name
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "Select"            , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Range_Method_Select          ))
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "Clear"             , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Range_Method_Clear           ))
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "ClearContents"     , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Range_Method_ClearContents   ))
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "ClearFormats"      , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Range_Method_ClearFormats    ))
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "ClearComments"     , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Range_Method_ClearComments   ))
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "ClearNotes"        , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Range_Method_ClearNotes      ))
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "AutoFit"           , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Range_Method_AutoFit         ))
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "Insert"            , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Range_Method_Insert          ))
'        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "Activate"          , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Worksheet_Method_Activate         ))
'
'        ' -- Common properties can take any name
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Range, "Value"               , %thinBasic_ReturnString     , CodePtr(cExcel_Range_Property_Value       ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Range, "Address"             , %thinBasic_ReturnString     , CodePtr(cExcel_Range_Property_Address     ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Range, "Formula"             , %thinBasic_ReturnString     , CodePtr(cExcel_Range_Property_Formula     ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Range, "FormulaR1C1"         , %thinBasic_ReturnString     , CodePtr(cExcel_Range_Property_FormulaR1C1 ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Range, "Font"                , %thinBasic_ReturnString     , CodePtr(cExcel_Range_Property_Font        ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Range, "HorizontalAlignment" , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Range_Property_HorizontalAlignment ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Range, "ColumnWidth"         , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Range_Property_ColumnWidth ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Range, "Interior"            , %thinBasic_ReturnString     , CodePtr(cExcel_Range_Property_Interior    ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Range, "NumberFormat"        , %thinBasic_ReturnString     , CodePtr(cExcel_Range_Property_NumberFormat))

        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Range, "End"                 , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Range_Property_End         ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Range, "Column"              , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Range_Property_Column      ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Range, "Row"                 , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Range_Property_Row         ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Range, "Copy"                , %thinBasic_ReturnCodeString , CodePtr(cExcel_Range_Property_Copy        ))
        
      End If

      thinBasic_AddEquate  "%XlHAlign_xlHAlignCenter"                 , "", %XlHAlign.xlHAlignCenter                    
      thinBasic_AddEquate  "%XlHAlign_xlHAlignCenterAcrossSelection"  , "", %XlHAlign.xlHAlignCenterAcrossSelection     
      thinBasic_AddEquate  "%XlHAlign_xlHAlignDistributed"            , "", %XlHAlign.xlHAlignDistributed               
      thinBasic_AddEquate  "%XlHAlign_xlHAlignFill"                   , "", %XlHAlign.xlHAlignFill                      
      thinBasic_AddEquate  "%XlHAlign_xlHAlignGeneral"                , "", %XlHAlign.xlHAlignGeneral                   
      thinBasic_AddEquate  "%XlHAlign_xlHAlignJustify"                , "", %XlHAlign.xlHAlignJustify                   
      thinBasic_AddEquate  "%XlHAlign_xlHAlignLeft"                   , "", %XlHAlign.xlHAlignLeft                      
      thinBasic_AddEquate  "%XlHAlign_xlHAlignRight"                  , "", %XlHAlign.xlHAlignRight
      
      thinBasic_AddEquate  "%XlColorIndex_xlColorIndexAutomatic"      , "", %XlColorIndex.xlColorIndexAutomatic                      
      thinBasic_AddEquate  "%XlColorIndex_xlColorIndexNone"           , "", %XlColorIndex.xlColorIndexNone

      thinBasic_AddEquate  "%XlUnderlineStyle_xlUnderlineStyleDouble"           , "", %XlUnderlineStyle.xlUnderlineStyleDouble  
      thinBasic_AddEquate  "%XlUnderlineStyle_xlUnderlineStyleDoubleAccounting" , "", %XlUnderlineStyle.xlUnderlineStyleDoubleAccounting
      thinBasic_AddEquate  "%XlUnderlineStyle_xlUnderlineStyleNone"             , "", %XlUnderlineStyle.xlUnderlineStyleNone    
      thinBasic_AddEquate  "%XlUnderlineStyle_xlUnderlineStyleSingle"           , "", %XlUnderlineStyle.xlUnderlineStyleSingle     
      thinBasic_AddEquate  "%XlUnderlineStyle_xlUnderlineStyleSingleAccounting" , "", %XlUnderlineStyle.xlUnderlineStyleSingleAccounting

      thinBasic_AddEquate  "%XlFileFormat_xlAddIn                      " , "", %XlFileFormat.xlAddIn                      
      thinBasic_AddEquate  "%XlFileFormat_xlCSV                        " , "", %XlFileFormat.xlCSV                        
      thinBasic_AddEquate  "%XlFileFormat_xlCSVMac                     " , "", %XlFileFormat.xlCSVMac                     
      thinBasic_AddEquate  "%XlFileFormat_xlCSVMSDOS                   " , "", %XlFileFormat.xlCSVMSDOS                   
      thinBasic_AddEquate  "%XlFileFormat_xlCSVWindows                 " , "", %XlFileFormat.xlCSVWindows                 
      thinBasic_AddEquate  "%XlFileFormat_xlDBF2                       " , "", %XlFileFormat.xlDBF2                       
      thinBasic_AddEquate  "%XlFileFormat_xlDBF3                       " , "", %XlFileFormat.xlDBF3                       
      thinBasic_AddEquate  "%XlFileFormat_xlDBF4                       " , "", %XlFileFormat.xlDBF4                       
      thinBasic_AddEquate  "%XlFileFormat_xlDIF                        " , "", %XlFileFormat.xlDIF                        
      thinBasic_AddEquate  "%XlFileFormat_xlExcel2                     " , "", %XlFileFormat.xlExcel2                     
      thinBasic_AddEquate  "%XlFileFormat_xlExcel2FarEast              " , "", %XlFileFormat.xlExcel2FarEast              
      thinBasic_AddEquate  "%XlFileFormat_xlExcel3                     " , "", %XlFileFormat.xlExcel3                     
      thinBasic_AddEquate  "%XlFileFormat_xlExcel4                     " , "", %XlFileFormat.xlExcel4                     
      thinBasic_AddEquate  "%XlFileFormat_xlExcel5                     " , "", %XlFileFormat.xlExcel5                     
      thinBasic_AddEquate  "%XlFileFormat_xlExcel7                     " , "", %XlFileFormat.xlExcel7                     
      thinBasic_AddEquate  "%XlFileFormat_xlExcel9795                  " , "", %XlFileFormat.xlExcel9795                  
      thinBasic_AddEquate  "%XlFileFormat_xlExcel4Workbook             " , "", %XlFileFormat.xlExcel4Workbook             
      thinBasic_AddEquate  "%XlFileFormat_xlIntlAddIn                  " , "", %XlFileFormat.xlIntlAddIn                  
      thinBasic_AddEquate  "%XlFileFormat_xlIntlMacro                  " , "", %XlFileFormat.xlIntlMacro                  
      thinBasic_AddEquate  "%XlFileFormat_xlWorkbookNormal             " , "", %XlFileFormat.xlWorkbookNormal             
      thinBasic_AddEquate  "%XlFileFormat_xlSYLK                       " , "", %XlFileFormat.xlSYLK                       
      thinBasic_AddEquate  "%XlFileFormat_xlTemplate                   " , "", %XlFileFormat.xlTemplate                   
      thinBasic_AddEquate  "%XlFileFormat_xlCurrentPlatformText        " , "", %XlFileFormat.xlCurrentPlatformText        
      thinBasic_AddEquate  "%XlFileFormat_xlTextMac                    " , "", %XlFileFormat.xlTextMac                    
      thinBasic_AddEquate  "%XlFileFormat_xlTextMSDOS                  " , "", %XlFileFormat.xlTextMSDOS                  
      thinBasic_AddEquate  "%XlFileFormat_xlTextPrinter                " , "", %XlFileFormat.xlTextPrinter                
      thinBasic_AddEquate  "%XlFileFormat_xlTextWindows                " , "", %XlFileFormat.xlTextWindows                
      thinBasic_AddEquate  "%XlFileFormat_xlWJ2WD1                     " , "", %XlFileFormat.xlWJ2WD1                     
      thinBasic_AddEquate  "%XlFileFormat_xlWK1                        " , "", %XlFileFormat.xlWK1                        
      thinBasic_AddEquate  "%XlFileFormat_xlWK1ALL                     " , "", %XlFileFormat.xlWK1ALL                     
      thinBasic_AddEquate  "%XlFileFormat_xlWK1FMT                     " , "", %XlFileFormat.xlWK1FMT                     
      thinBasic_AddEquate  "%XlFileFormat_xlWK3                        " , "", %XlFileFormat.xlWK3                        
      thinBasic_AddEquate  "%XlFileFormat_xlWK4                        " , "", %XlFileFormat.xlWK4                        
      thinBasic_AddEquate  "%XlFileFormat_xlWK3FM3                     " , "", %XlFileFormat.xlWK3FM3                     
      thinBasic_AddEquate  "%XlFileFormat_xlWKS                        " , "", %XlFileFormat.xlWKS                        
      thinBasic_AddEquate  "%XlFileFormat_xlWorks2FarEast              " , "", %XlFileFormat.xlWorks2FarEast              
      thinBasic_AddEquate  "%XlFileFormat_xlWQ1                        " , "", %XlFileFormat.xlWQ1                        
      thinBasic_AddEquate  "%XlFileFormat_xlWJ3                        " , "", %XlFileFormat.xlWJ3                        
      thinBasic_AddEquate  "%XlFileFormat_xlWJ3FJ3                     " , "", %XlFileFormat.xlWJ3FJ3                     
      thinBasic_AddEquate  "%XlFileFormat_xlUnicodeText                " , "", %XlFileFormat.xlUnicodeText                
      thinBasic_AddEquate  "%XlFileFormat_xlHtml                       " , "", %XlFileFormat.xlHtml                       
      thinBasic_AddEquate  "%XlFileFormat_xlWebArchive                 " , "", %XlFileFormat.xlWebArchive                 
      thinBasic_AddEquate  "%XlFileFormat_xlXMLSpreadsheet             " , "", %XlFileFormat.xlXMLSpreadsheet             
      thinBasic_AddEquate  "%XlFileFormat_xlExcel12                    " , "", %XlFileFormat.xlExcel12                    
      thinBasic_AddEquate  "%XlFileFormat_xlOpenXMLWorkbook            " , "", %XlFileFormat.xlOpenXMLWorkbook            
      thinBasic_AddEquate  "%XlFileFormat_xlOpenXMLWorkbookMacroEnabled" , "", %XlFileFormat.xlOpenXMLWorkbookMacroEnabled
      thinBasic_AddEquate  "%XlFileFormat_xlOpenXMLTemplateMacroEnabled" , "", %XlFileFormat.xlOpenXMLTemplateMacroEnabled
      thinBasic_AddEquate  "%XlFileFormat_xlTemplate8                  " , "", %XlFileFormat.xlTemplate8                  
      thinBasic_AddEquate  "%XlFileFormat_xlOpenXMLTemplate            " , "", %XlFileFormat.xlOpenXMLTemplate            
      thinBasic_AddEquate  "%XlFileFormat_xlAddIn8                     " , "", %XlFileFormat.xlAddIn8                     
      thinBasic_AddEquate  "%XlFileFormat_xlOpenXMLAddIn               " , "", %XlFileFormat.xlOpenXMLAddIn               
      thinBasic_AddEquate  "%XlFileFormat_xlExcel8                     " , "", %XlFileFormat.xlExcel8                     
      thinBasic_AddEquate  "%XlFileFormat_xlOpenDocumentSpreadsheet    " , "", %XlFileFormat.xlOpenDocumentSpreadsheet    
      thinBasic_AddEquate  "%XlFileFormat_xlOpenXMLStrictWorkbook      " , "", %XlFileFormat.xlOpenXMLStrictWorkbook      
      thinBasic_AddEquate  "%XlFileFormat_xlWorkbookDefault            " , "", %XlFileFormat.xlWorkbookDefault            
       
      thinBasic_AddEquate  "%XlDirection_xlDown                        " , "", %XlDirection.xlDown              
      thinBasic_AddEquate  "%XlDirection_xlToLeft                      " , "", %XlDirection.xlToLeft            
      thinBasic_AddEquate  "%XlDirection_xlToRight                     " , "", %XlDirection.xlToRight           
      thinBasic_AddEquate  "%XlDirection_xlUp                          " , "", %XlDirection.xlUp                


 


  End Function

  '----------------------------------------------------------------------------
  Function UnLoadLocalSymbols Alias "UnLoadLocalSymbols" () Export As Long
  ' This function is automatically called by thinCore whenever this DLL is unloaded.
  ' This function is not mandatory. CAN be present but it is not necessary.
  ' Use this function to perform uninitialize process, if needed.
  '----------------------------------------------------------------------------

  FUNCTION = 0&

  End Function


  %DLL_PROCESS_ATTACH   = 1
  %DLL_THREAD_ATTACH    = 2
  %DLL_THREAD_DETACH    = 3
  %DLL_PROCESS_DETACH   = 0
  Function LibMain Alias "LibMain" (ByVal hInstance   As Long, _
                                    ByVal fwdReason   As Long, _
                                    ByVal lpvReserved As Long) Export As Long
    Select Case fwdReason
      Case %DLL_PROCESS_ATTACH
  
        Function = 1
        Exit Function
      Case %DLL_PROCESS_DETACH
  
        Function = 1
        Exit Function
      Case %DLL_THREAD_ATTACH
  
        Function = 1
        Exit Function
      Case %DLL_THREAD_DETACH
  
        Function = 1
        Exit Function
    End Select
  
  End Function
