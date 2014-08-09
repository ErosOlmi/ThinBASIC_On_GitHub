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
#RESOURCE VERSION$ "FileVersion",      "1.0.0.1"
#RESOURCE VERSION$ "InternalName",     "Excel"
#RESOURCE VERSION$ "OriginalFilename", "ThinBASIC_Excel.dll"
#RESOURCE VERSION$ "LegalCopyright",   "Copyright © Petr Schreiber/Eros Olmi 2014"
#RESOURCE VERSION$ "ProductName",      "Module"
#RESOURCE VERSION$ "ProductVersion",   "1.0.0.1"
#RESOURCE VERSION$ "Comments",         "Support site: http://www.thinbasic.com/"

GLOBAL gPath AS STRING

'---Every used defined thinBasic module must include this file
#Include Once "\ThinBASIC\Lib\thinCore.inc"

#INCLUDE ONCE "Excel.inc"
#Include Once ".\thinBasic_Excel_Application.inc"
#Include Once ".\thinBasic_Excel_Workbook.inc"
#Include Once ".\thinBasic_Excel_Worksheet.inc"
#Include Once ".\thinBasic_Excel_Range.inc"


'----------------------------------------------------------------------------
'---References:
'     Excel Object Model: http://msdn.microsoft.com/en-us/library/bb149081(v=office.12).aspx
'----------------------------------------------------------------------------

'----------------------------------------------------------------------------
FUNCTION LoadLocalSymbols ALIAS "LoadLocalSymbols" (OPTIONAL BYVAL sPath AS STRING) EXPORT AS LONG
    LOCAL RetCode                   AS LONG
    LOCAL pClass_cExcel_Application AS LONG
    LOCAL pClass_cExcel_Workbook    AS LONG
    Local pClass_cExcel_Worksheet   As Long
    Local pClass_cExcel_Range       As Long

    ' -- Save DLL loading path to global var
    gPath = sPath

    '---
    ' Excel Application Class
    '---
      pClass_cExcel_Application = thinBasic_Class_Add("Excel_Application", 0)

      ' -- If class was created, define all methods and properties, each connected to a CODEPTR module function/sub
      IF pClass_cExcel_Application THEN

        ' -- Constructor wrapper function needs to be linked in as _Create
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Application, "_Create"           , %thinBasic_ReturnNone       , CODEPTR(cExcel_Application_Create            ))
        ' -- Destructor wrapper function needs to be linked in as _Destroy
        ' -- WARNING: You MUST supply destructor and set the object to NOTHING, otherwise you risk memory leak
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Application, "_Destroy"          , %thinBasic_ReturnNone       , CODEPTR(cExcel_Application_Destroy           ))
        ' -- ClassObject
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Application, "_GetClassObject"   , %thinBasic_ReturnNone       , CODEPTR(cExcel_Application_GetClassObject    ))

        ' -- Common methods can take any name
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Application, "Quit"              , %thinBasic_ReturnCodeLong   , CODEPTR(cExcel_Application_Method_Quit     ))

        ' -- Common properties can take any name
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "Version"                 , %thinBasic_ReturnString     , CODEPTR(cExcel_Application_Property_Version ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "Visible"                 , %thinBasic_ReturnCodeLong   , CODEPTR(cExcel_Application_Property_Visible ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "AlertBeforeOverwriting"  , %thinBasic_ReturnCodeLong   , CODEPTR(cExcel_Application_Property_AlertBeforeOverwriting  ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "DisplayAlerts"           , %thinBasic_ReturnCodeLong   , CODEPTR(cExcel_Application_Property_DisplayAlerts           ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "ActiveWindow"            , %thinBasic_ReturnString     , CodePtr(cExcel_Application_Property_ActiveWindow            ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "Workbooks"               , %thinBasic_ReturnString     , CodePtr(cExcel_Application_Property_Workbooks               ))

        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "ActiveWorkbook"          , %thinBasic_ReturnString     , CodePtr(cExcel_Application_Property_ActiveWorkbook          ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Application, "ActiveSheet"             , %thinBasic_ReturnString     , CodePtr(cExcel_Application_Property_ActiveSheet             ))

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
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Workbook, "_GetClassObject"            , %thinBasic_ReturnNone       , CodePtr(cExcel_Workbook_GetClassObject    ))

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

      END IF

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
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Worksheet, "_GetClassObject"   , %thinBasic_ReturnNone       , CODEPTR(cExcel_Worksheet_GetClassObject   ))

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
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "_Create"           , %thinBasic_ReturnNone       , CodePtr(cExcel_Range_Create           ))
        ' -- Constructor wrapper function needs to be linked in as _Create
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "_CreateDirect"     , %thinBasic_ReturnNone       , CodePtr(cExcel_Range_Create_Direct    ))
        ' -- Destructor wrapper function needs to be linked in as _Destroy
        ' -- WARNING: You MUST supply destructor and set the object to NOTHING, otherwise you risk memory leak
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "_Destroy"          , %thinBasic_ReturnNone       , CodePtr(cExcel_Range_Destroy          ))
        ' -- ClassObject
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "_GetClassObject"   , %thinBasic_ReturnNone       , CodePtr(cExcel_Range_GetClassObject   ))

        ' -- Common methods can take any name
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "Select"            , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Range_Method_Select     ))
        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "Clear"             , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Range_Method_Clear      ))
'        RetCode = thinBasic_Class_AddMethod   (pClass_cExcel_Range, "Activate"          , %thinBasic_ReturnCodeLong   , CodePtr(cExcel_Worksheet_Method_Activate         ))
'
'        ' -- Common properties can take any name
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Range, "Value"             , %thinBasic_ReturnString     , CodePtr(cExcel_Range_Property_Value   ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Range, "Address"           , %thinBasic_ReturnString     , CodePtr(cExcel_Range_Property_Address ))
        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Range, "Formula"           , %thinBasic_ReturnString     , CodePtr(cExcel_Range_Property_Formula ))
'        RetCode = thinBasic_Class_AddProperty (pClass_cExcel_Range, "Name"              , %thinBasic_ReturnString     , CodePtr(cExcel_Worksheet_Property_Name    ))

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
