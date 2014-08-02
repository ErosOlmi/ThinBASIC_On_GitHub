'
' Excel support for ThinBASIC
'
' Petr Schreiber
'

#COMPILE DLL
#REGISTER NONE
#DIM ALL

#RESOURCE VERSIONINFO
#RESOURCE FILEVERSION 1, 0, 0, 1
#RESOURCE PRODUCTVERSION 1, 0, 0, 1

#RESOURCE STRINGINFO "0409", "04B0"

#RESOURCE VERSION$ "CompanyName",      "Petr Schreiber"
#RESOURCE VERSION$ "FileDescription",  "thinBasic module for Excel support"
#RESOURCE VERSION$ "FileVersion",      "1.0.0.1"
#RESOURCE VERSION$ "InternalName",     "Excel"
#RESOURCE VERSION$ "OriginalFilename", "ThinBASIC_Excel.dll"
#RESOURCE VERSION$ "LegalCopyright",   "Copyright � Petr Schreiber 2014"
#RESOURCE VERSION$ "ProductName",      "Module"
#RESOURCE VERSION$ "ProductVersion",   "1.0.0.1"
#RESOURCE VERSION$ "Comments",         "Support site: http://www.thinbasic.com/"

Global gPath As String

'---Every used defined thinBasic module must include this file
#Include Once "..\thinCore.inc"  

#Include Once "Excel.inc"
#Include Once "thinBasic_Excel_Application.inc"
#Include Once "thinBasic_Excel_Workbook.inc"
#Include Once "thinBasic_Excel_Worksheet.inc"
 
 
'----------------------------------------------------------------------------
'---References:
'     Excel Object Model: http://msdn.microsoft.com/en-us/library/bb149081(v=office.12).aspx
'----------------------------------------------------------------------------

'----------------------------------------------------------------------------
FUNCTION LoadLocalSymbols ALIAS "LoadLocalSymbols" (OPTIONAL BYVAL sPath AS STRING) EXPORT AS LONG
    Local RetCode                   As Long
    Local pClass_cExcel_Application As Long
    Local pClass_cExcel_Workbook    As Long
    Local pClass_cExcel_Worksheet   As Long

    '---Save DLL loading path to global var
    gPath = sPath

    '---
    ' Excel Application Class
    '---
      pClass_cExcel_Application = thinBasic_Class_Add("Excel_Application", 0)
  
      '---If class was created, define all methods and properties, each connected to a CODEPTR module function/sub
      If pClass_cExcel_Application Then
  
        ' -- Constructor wrapper function needs to be linked in as _Create
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Application, "_Create"           , %thinBasic_ReturnNone       , CodePtr(cExcel_Application_Create            ))
        ' -- Destructor wrapper function needs to be linked in as _Create
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Application, "_Destroy"          , %thinBasic_ReturnNone       , CodePtr(cExcel_Application_Destroy           ))
        ' -- ClassObject
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Application, "_GetClassObject"   , %thinBasic_ReturnNone       , CodePtr(cExcel_Application_GetClassObject    ))
  
        ' -- Common methods can take any name
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Application, "Quit"              , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Application_Method_Quit     ))

        ' -- Common properties can take any name
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "Version"                 , %thinBasic_ReturnString     , CodePtr(cExcel_Application_Property_Version ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "Visible"                 , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Application_Property_Visible ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "AlertBeforeOverwriting"  , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Application_Property_AlertBeforeOverwriting  ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "DisplayAlerts"           , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Application_Property_DisplayAlerts           ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "ActiveWindow"            , %thinBasic_ReturnString     , CodePtr(cExcel_Application_Property_ActiveWindow            ))
                                                                                                                                           
      End If

    '---
    ' Excel Workbook Class
    '---
      pClass_cExcel_Workbook = thinBasic_Class_Add("Excel_Workbook", 0)
  
      '---If class was created, define all methods and properties, each connected to a CODEPTR module function/sub
      If pClass_cExcel_Workbook Then
  
        ' -- Constructor wrapper function needs to be linked in as _Create
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Workbook, "_Create"           , %thinBasic_ReturnNone       , CodePtr(cExcel_Workbook_Create            ))
        ' -- Destructor wrapper function needs to be linked in as _Create
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Workbook, "_Destroy"          , %thinBasic_ReturnNone       , CodePtr(cExcel_Workbook_Destroy           ))
        ' -- ClassObject
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Workbook, "_GetClassObject"   , %thinBasic_ReturnNone       , CodePtr(cExcel_Workbook_GetClassObject    ))
  
        ' -- Common methods can take any name
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Workbook, "SaveAs"             , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Workbook_Method_SaveAs     ))

        ' -- Common properties can take any name
        
      End If

    '---
    ' Excel Worksheet Class
    '---
      pClass_cExcel_Worksheet = thinBasic_Class_Add("Excel_Worksheet", 0)
  
      '---If class was created, define all methods and properties, each connected to a CODEPTR module function/sub
      If pClass_cExcel_Worksheet Then
  
        ' -- Constructor wrapper function needs to be linked in as _Create
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Worksheet, "_Create"           , %thinBasic_ReturnNone       , CodePtr(cExcel_Worksheet_Create           ))
        ' -- Destructor wrapper function needs to be linked in as _Create
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Worksheet, "_Destroy"          , %thinBasic_ReturnNone       , CodePtr(cExcel_Worksheet_Destroy          ))
        ' -- ClassObject
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Worksheet, "_GetClassObject"   , %thinBasic_ReturnNone       , CodePtr(cExcel_Worksheet_GetClassObject   ))

        ' -- Common methods can take any name
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Worksheet, "PrintPreview"      , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Worksheet_Method_PrintPreview     ))
  
        ' -- Common properties can take any name
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Worksheet, "Cells"             , %thinBasic_ReturnString     , CodePtr(cExcel_Worksheet_Property_Cells   ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Worksheet, "Name"              , %thinBasic_ReturnString     , CodePtr(cExcel_Worksheet_Property_Name    ))
        
      End If

END FUNCTION

'----------------------------------------------------------------------------
FUNCTION UnLoadLocalSymbols ALIAS "UnLoadLocalSymbols" () EXPORT AS LONG

  FUNCTION = 0&

END FUNCTION

%DLL_PROCESS_ATTACH   = 1
%DLL_THREAD_ATTACH    = 2
%DLL_THREAD_DETACH    = 3
%DLL_PROCESS_DETACH   = 0
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