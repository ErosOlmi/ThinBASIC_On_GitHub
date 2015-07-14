  #If 0
  '=============================================================================
  'File Name   : thinCore.inc
  'Author      : Eros Olmi
  'Version     : 1.9.0.0
  'Description : thinBasic external library Interface definitions
  'Web site    : http://www.thinbasic.com
  '=============================================================================
  'COPYRIGHT And PERMISSION NOTICE
  '=============================================================================
  'Copyright (c) 2003 - 2010, Eros Olmi, <eros.olmi@thinbasic.com>
  ' 
  'All rights reserved.
  ' 
  'Permission To use this software Is granted Only For the purpose To develop
  'thinBasic language modules both For commercial Or non commercial purpose.
  'In Any Case the above copyright notice And this permission notice must appear
  'In All Copies.
  ' 
  'THE SOFTWARE Is PROVIDED "AS IS", WITHOUT WARRANTY Of Any KIND, EXPRESS Or
  'IMPLIED, INCLUDING BUT Not LIMITED To THE WARRANTIES Of MERCHANTABILITY,
  'FITNESS For A PARTICULAR PURPOSE And NONINFRINGEMENT Of THIRD PARTY RIGHTS.
  'In NO EVENT SHALL THE AUTHORS Or COPYRIGHT HOLDERS BE LIABLE For Any CLAIM,
  'DAMAGES Or OTHER LIABILITY, WHETHER In AN ACTION Of CONTRACT, TORT Or
  'OTHERWISE, ARISING From, Out Of Or In CONNECTION With THE SOFTWARE Or THE
  'USE Or OTHER DEALINGS In THE SOFTWARE.
  '
  '=============================================================================
  'Edit History [+]Add, [-]Fixed, [*]Changed, [N]Note
  '_____________________________________________________________________________
  'Date_______Type_Note_________________________________________________________
  '
  '2010/08/21 [+]  thinBasic_Class_Add()         
  '=============================================================================
  #EndIf

  #IF NOT %DEF(%TRUE)
    %TRUE  = 1
  #ENDIF
  #IF NOT %DEF(%FALSE)
    %FALSE = 0
  #ENDIF


  '------------------------------------------------------------------------------
  'Global declares
  '------------------------------------------------------------------------------
  %thinBasic_ForceOverWrite         =  1   'Used in thinBasic_LoadSymbol to force symbol over writing

  %VarMainType_IsNumber             = 20
  %VarMainType_IsString             = 30
  %VarMainType_IsAsciiZ             = 25
  %VarMainType_IsVariant            = 50
                                    
  %VarSubType_Byte                  =  1
  %VarSubType_Integer               =  2
  %VarSubType_Word                  =  3
  %VarSubType_DWord                 =  4
  %VarSubType_Long                  =  5
  %VarSubType_Quad                  =  6              
  %VarSubType_Single                =  7 
  %VarSubType_Double                =  8
  %VarSubType_Currency              =  9
  %VarSubType_Ext                   = 10
  %VarSubType_Variant               = 50


  '------------------------------------------------------------------------------
  'Return Types from functions/Subs
  '------------------------------------------------------------------------------
  %thinBasic_ReturnNone             =  0   'Used in thinBasic_LoadSymbol to define a sub
  %thinBasic_ReturnNumber           = 20   'Used in thinBasic_LoadSymbol to define a function returning a EXT number
  %thinBasic_ReturnString           = 30   'Used in thinBasic_LoadSymbol to define a function returning a string

  %thinBasic_ReturnCodeByte         =  1
  %thinBasic_ReturnCodeInteger      =  2
  %thinBasic_ReturnCodeWord         =  3
  %thinBasic_ReturnCodeDWord        =  4
  %thinBasic_ReturnCodeLong         =  5
  %thinBasic_ReturnCodeQuad         =  6
  %thinBasic_ReturnCodeSingle       =  7 
  %thinBasic_ReturnCodeDouble       =  8
  %thinBasic_ReturnCodeCurrency     =  9
  %thinBasic_ReturnCodeExt          = 10
  '------------------------------------------------------------------------------

  %thinBasic_BufferType_IsFile      = 0
  %thinBasic_BufferType_IsScript    = 1

  Declare Function  thinBasic_Init                                _
                          Lib "thinCore.dll"                      _
                          Alias "thinBasic_Init"                  _
                          (                                       _
                            ByVal hWnd            As Long       , _
                            ByVal cInstance       As Long       , _
                            ByVal sKey            As String       _
                          ) As Long
  Declare Function  thinBasic_Release                             _
                          Lib "thinCore.dll"                      _
                          Alias "thinBasic_Release"               _
                          (                                       _
                            ByVal hScript         As Long         _
                          ) As Long
  Declare Function  thinBasic_Run                                 _
                          Lib "thinCore.dll"                      _
                          Alias "thinBasic_Run"                   _
                          (                                       _
                            ByVal hScript         As Long       , _
                            byval sBuffer         As string     , _
                            ByVal BufferType      As Long       , _
                  Optional  ByVal Options         As Long       , _
                            ByVal DebugMode       As Long       , _
                            ByVal LogMode         As Long       , _
                            ByVal ObfuscateMode   As Long       , _
                            ByVal CallingProgram  As Long       , _
                            ByVal DependancyMode  As Long         _
                          ) As Long

  
  '----------------------------------------------------------------------------
  'thinBasic_LoadSymbol
  '----------------------------------------------------------------------------
  ' Description
  '   Creates e new keyword inside interpreter.
  '   You define your own Function or Sub inside your program. Then call
  '   thinBasic_LoadSymbol to add a new keyword connected to your new function/sub
  '
  ' Parameters
  '   SymbolName  : Name you want to give to your sub or function inside the
  '                 interpreter. It is the new keyword name.
  '
  '   ReturnCode  : specify what kind of var your function returns
  '                 Use predefined constants:
  '                 %thinBasic_ReturnNumber
  '                 %thinBasic_ReturnString
  '                 %thinBasic_ReturnNone
  '                 If SymbolName will end with "$" char, ReturnCode will
  '                 be forced to %thinBasic_ReturnString
  '                 Note: if you pass a pointer to a SUB routine and ReturnCode
  '                       is different from %thinBasic_ReturnNone, a GPF will be
  '                       fired by the Operating System.
  '                 Note: if ReturnCode <= 0 then %thinBasic_ReturnNone will be
  '                       assumed 
  '
  '   FunctionOrSubPointer:
  '                 DWORD returned by CODEPTR function pointing to the function
  '                 that will be called when SymbolName keyword will be encountered
  '                 during script execution
  '
  '   ForceOverWrite:
  '                 If different from 0 it means that even if SymbolName
  '                 already exists as a keyword, you want to overwrite it with
  '                 a new one.
  '                 For example, if you want to overwrite behave of MID$
  '                 creating your own MID$ function, set this parameter to a value
  '                 > 0 like in this example:
  '                 thinBasic_LoadSymbol("MID$", %thinBasic_ReturnString, CODEPTR(MyMid), 1)
  '
  ' Possible Return Code
  '   >0 = no error, return the associated SymbolName ID
  '   -1 = empty SymbolName
  '   -2 = duplicated SymbolName and no ForceOverWrite flag
  '   -3 = invalid sub/function pointer
  '   -4 = SymbolName contains invalid char(s)
  '
  ' Example
  '   thinBasic_LoadSymbol "MyNewFunction", %thinBasic_ReturnNumber, CODEPTR(MyFunc)
  '----------------------------------------------------------------------------
  Declare Function  thinBasic_LoadSymbol                            _
                          Lib "thinCore.DLL"                        _
                          Alias "thinBasic_LoadSymbol"              _
                          (                                         _
                            ByVal SymbolName            As String , _
                            ByVal ReturnCode            As Long   , _
                            ByVal FunctionOrSubPointer  As Dword  , _
                            Optional                                _
                            ByVal ForceOverWrite        As Long     _
                          ) As Long

  '----------------------------------------------------------------------------
  'thinBasic_LoadSymbolEX
  '----------------------------------------------------------------------------
  ' Description
  '   Creates e new keyword inside interpreter.
  '   You define your own Function or Sub inside your program. Then call
  '   thinBasic_LoadSymbol to add a new keyword connected to your new function/sub
  '
  ' Parameters
  '   SymbolName  : Name you want to give to your sub or function inside the
  '                 interpreter. It is the new keyword name.
  '
  '   ReturnCode  : specify what kind of var your function returns
  '                 Use predefined constants:
  '                 %thinBasic_ReturnNumber
  '                 %thinBasic_ReturnString
  '                 %thinBasic_ReturnNone
  '                 If SymbolName will end with "$" char, ReturnCode will
  '                 be forced to %thinBasic_ReturnString
  '                 Note: if you pass a pointer to a SUB routine and ReturnCode
  '                       is different from %thinBasic_ReturnNone, a GPF will be
  '                       fired by the Operating System.
  '                 Note: if ReturnCode <= 0 then %thinBasic_ReturnNone will be
  '                       assumed 
  '
  '   FunctionOrSubPointer:
  '                 DWORD returned by CODEPTR function pointing to the function
  '                 that will be called when SymbolName keyword will be encountered
  '                 during script execution
  '
  '   ForceOverWrite:
  '                 If different from 0 it means that even if SymbolName
  '                 already exists as a keyword, you want to overwrite it with
  '                 a new one.
  '                 For example, if you want to overwrite behave of MID$
  '                 creating your own MID$ function, set this parameter to a value
  '                 > 0 like in this example:
  '                 thinBasic_LoadSymbol("MID$", %thinBasic_ReturnString, CODEPTR(MyMid), 1)
  '   sSyntax:
  '                 Syntax of the symbol name
  '   sHelp:
  '                 help about the symbol name
  '
  ' Possible Return Code
  '   >0 = no error, return the associated SymbolName ID
  '   -1 = empty SymbolName
  '   -2 = duplicated SymbolName and no ForceOverWrite flag
  '   -3 = invalid sub/function pointer
  '   -4 = SymbolName contains invalid char(s)
  '
  ' Example
  '   thinBasic_LoadSymbol "MyNewFunction", %thinBasic_ReturnNumber, CODEPTR(MyFunc)
  '----------------------------------------------------------------------------
  Declare Function  thinBasic_LoadSymbolEX                          _
                          Lib "thinCore.DLL"                        _
                          Alias "thinBasic_LoadSymbolEX"            _
                          (                                         _
                            ByVal SymbolName            As String , _
                            ByVal ReturnCode            As Long   , _
                            ByVal FunctionOrSubPointer  As Dword  , _
                            Optional                                _
                            ByVal ForceOverWrite        As Long   , _
                            ByVal sSyntax               As String , _
                            ByVal sHelp                 As String   _
                          ) As Long

  '----------------------------------------------------------------------------
  'thinBasic_VariableGetList
  '----------------------------------------------------------------------------
  ' Returns a string with the list of all variables of a specific stack level
  ' If StackLevel is missing, current stack level will be returned
  '----------------------------------------------------------------------------
  Declare Function thinBasic_VariableGetList                      _
                          LIB "thinCore.DLL"                      _
                          Alias "thinBasic_VariableGetList"       _
                          (                                       _
                            Optional                              _
                            ByVal lStackLevel   As Long         , _
                            ByVal sSep          As String         _
                          ) As String
  '----------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  'thinBasic_VariableExists
  '----------------------------------------------------------------------------
  ' Returns true/false depending if the given name is a variable or not
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_VariableExists                      _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_VariableExists"        _
                          (                                       _
                            BYVAL vName       AS STRING           _
                          ) AS LONG


  '----------------------------------------------------------------------------
  'thinBasic_VariableGetValueNum
  '----------------------------------------------------------------------------
  ' Returns numeric value of a numeric variable given variable name and optionally array index
  '----------------------------------------------------------------------------
  Declare Function  thinBasic_VariableGetValueNum                 _
                          LIB "thinCore.DLL"                      _
                          Alias "thinBasic_VariableGetValueNum"   _
                          (                                       _
                                      ByVal SearchKey As String , _
                            Optional  ByVal lIndex    As Long     _
                          ) As Ext
  '----------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  'thinBasic_VariableGetInfo
  '----------------------------------------------------------------------------
  ' Returns additional variable info giving variable name.
  ' ATTENTION: pass variable as reference. Function will populate
  '            with relevant information
  ' Function returns a pointer to internal thinBAsic variable structure data
  ' that can be used in other situations.
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_VariableGetInfo                     _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_VariableGetInfo"       _
                          (                                       _
                            BYVAL SearchKey   AS STRING         , _
                            BYREF MainType    AS LONG           , _       '---ATTENTION: parameter passed BYREF will return info
                            BYREF SubType     AS LONG           , _       '---ATTENTION: parameter passed BYREF will return info
                            BYREF IsArray     AS LONG             _       '---ATTENTION: parameter passed BYREF will return info
                          ) AS LONG
                          
  Declare Function  thinBasic_VariableGetInfoPtr                  _
                          Lib "thinCore.DLL"                      _
                          Alias "thinBasic_VariableGetInfoPtr"    _
                          (                                       _
                            ByVal pVariable   As Long           , _       '---A pointer to Variable
                            ByRef MainType    As Long           , _       '---ATTENTION: parameter passed BYREF will return info
                            ByRef SubType     As Long           , _       '---ATTENTION: parameter passed BYREF will return info
                            ByRef IsArray     As Long             _       '---ATTENTION: parameter passed BYREF will return info
                          ) As Long

  '----------------------------------------------------------------------------
  'thinBasic_VariableGetInfoEX
  '----------------------------------------------------------------------------
  ' Returns additional variable info giving variable name.
  ' ATTENTION: pass variable as reference. Function will populate
  '            with relevant information
  ' Function returns a pointer to internal thinBAsic variable structure data
  ' that can be used in other situations.
  '----------------------------------------------------------------------------
  declare function  thinBasic_VariableGetInfoEX                   _
                          lib "thinCore.DLL"                      _
                          alias "thinBasic_VariableGetInfoEX"     _
                          (                                       _
                            byval SearchKey   as string         , _
                            byref MainType    as long           , _   '---ATTENTION: parameter passed BYREF will return info
                            byref SubType     as long           , _   '---ATTENTION: parameter passed BYREF will return info
                            byref IsArray     as long           , _   '---ATTENTION: parameter passed BYREF will return info
                            byref DataPtr     as long           , _   '---ATTENTION: parameter passed BYREF will return info
                            ByRef nElements   As Long           , _   '---ATTENTION: parameter passed BYREF will return info
                            Optional                              _
                            ByVal WhichLevel  As Long             _   
                          ) As Long

'  '---Equates for variable Sub Type
'    %SubType_Byte             =  1?
'    %SubType_Integer          =  2?
'    %SubType_Boolean          =  2?
'    %SubType_Word             =  3?
'    %SubType_DWord            =  4?
'    %SubType_Long             =  5?
'    %SubType_Quad             =  6?
'    %SubType_Single           =  7?
'    %SubType_Double           =  8?
'    %SubType_Currency         =  9?
'    %SubType_Ext              = 10?
'
'    %MainType_IsNumber        = 15?
'
'    %SubType_AsciiZ           = 16? 
'    %SubType_String           = 17? 
'
'    %MainType_GUID            = 20?               : %MainType_IsGUID    = %MainType_GUID
'    %MainType_String          = %SubType_String   : %MainType_IsString  = %MainType_String
'
'    %MainType_Variant         = 23?               : %MainType_IsVariant = %MainType_Variant
'    %MainType_UDT             = 24?               : %MainType_IsUDT     = %MainType_UDT
'    %MainType_PTR             = 25?               : %MainType_IsPTR     = %MainType_PTR
'    %MainType_Object          = 26?               : %MainType_IsObject  = %MainType_Object
'    %MainType_Class           = 27?               : %MainType_IsClass   = %MainType_Class
  '---Equates for variable Sub Type
    %SubType_Byte             =  1?
    %SubType_Integer          =  2?
    %SubType_Boolean          =  2?
    %SubType_Word             =  3?
    %SubType_DWord            =  4?
    %SubType_Long             =  5?
    %SubType_Quad             =  6?
    %SubType_Single           =  7?
    %SubType_Double           =  8?
    %SubType_Currency         =  9?
    %SubType_Ext              = 10?

    %MainType_IsNumber        = 20?

    %SubType_AsciiZ           = 25? 
    %SubType_String           = 30? 

    %MainType_GUID            = 40?               : %MainType_IsGUID    = %MainType_GUID
    %MainType_String          = %SubType_String   : %MainType_IsString  = %MainType_String

    %MainType_Variant         = 50?               : %MainType_IsVariant = %MainType_Variant
    %MainType_UDT             = 60?               : %MainType_IsUDT     = %MainType_UDT
    %MainType_PTR             = 70?               : %MainType_IsPTR     = %MainType_PTR
    %MainType_Object          = 80?               : %MainType_IsObject  = %MainType_Object
    %MainType_Class           = 90?               : %MainType_IsClass   = %MainType_Class
  



  '----------------------------------------------------------------------------
  'thinBasic_VariableIsArray
  '----------------------------------------------------------------------------
  ' Returns true/false depending if given name is an array
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_VariableIsArray                     _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_VariableIsArray"       _
                          (                                       _
                            BYVAL vName       AS STRING           _
                          ) AS LONG

  '----------------------------------------------------------------------------
  'thinBasic_VariableRedim
  '----------------------------------------------------------------------------
  ' Redim a variable giving variable ptr
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_VariableRedim                       _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_VariableRedim"         _
                          (                                       _
                            BYVAL pVar AS LONG,                   _
                            BYVAL lPreserve AS LONG,              _
                            BYVAL VarDim1 AS LONG,                _
                            ByVal VarDim2 As Long,                _
                            BYVAL VarDim3 AS LONG                 _
                          ) AS LONG

  '----------------------------------------------------------------------------
  'thinBasic_ArrayGetElements
  '----------------------------------------------------------------------------
  ' Returns total number of elements inside an array giving the array name
  ' ATTENTION: this function works on single or multi dimensions arrays
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_ArrayGetElements                    _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_ArrayGetElements"      _
                          (                                       _
                            BYVAL vName       AS STRING           _
                          ) AS LONG

  '----------------------------------------------------------------------------
  'thinBasic_ArrayGetPtr
  '----------------------------------------------------------------------------
  ' Returns the data structure pointer of an array giving its name
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_ArrayGetPtr                         _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_ArrayGetPtr"           _
                          (                                       _
                            BYVAL vName       AS STRING           _
                          ) AS LONG

  '----------------------------------------------------------------------------
  'thinBasic_ArrayGetInfo
  '----------------------------------------------------------------------------
  ' Returns information on an array passing array pointer returned by thinBasic_ArrayGetPtr
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_ArrayGetInfo                        _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_ArrayGetInfo"          _
                          (                                       _
                            BYVAL ArrayPtr    AS LONG,            _
                            BYVAL InfoType    AS LONG             _
                          ) AS LONG

  '---Equates for InfoType
    %Array_ElementsCount        =  1     'Total number of elements in the array
    %Array_ElementSize          =  2     'The size of the single element (for example an array of LONGs will return 4)
    %Array_Dimensions           =  3     'Number of dimensions (from 1 to 3)
    %Array_Size                 =  4     'Memory size allocated for the array. This will not compute the memory needed for strings
    %Array_ElementsType         = 10     'Type of elements strored into the array. See equates for thinBasic_VariableGetInfo
    %Array_ElementsAreFixed     = 15     '%TRUE if elements are fixed size, like fixed strings or UDT

    %Array_UBoundDim_1          = 91     'Returns the Upper Bound of dimension 1
    %Array_UBoundDim_2          = 92     'Returns the Upper Bound of dimension 2
    %Array_UBoundDim_3          = 93     'Returns the Upper Bound of dimension 3
  '----------------------------------------------------------------------------
  'thinBasic_AddEquate
  '----------------------------------------------------------------------------
  ' Add a new equate (constant)
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_AddEquate                           _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_AddEquate"             _
                          (                                       _
                            BYVAL vName       AS STRING,          _
                            BYVAL lValString  AS STRING,          _
                            BYVAL lValNumber  AS EXT   ,          _
                            OPTIONAL                              _
                            BYVAL ForceType   AS LONG             _      
                          ) As Long                                
                          
  '---Used for thinBasic_AddEquate to force an equate type in case of missing first char $ or %
  %EquateTypeNumber    = 3
  %EquateTypeString    = 5

  '----------------------------------------------------------------------------
  'thinBasic_AddVariable
  '----------------------------------------------------------------------------
  ' Add a new variable to current stack level
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_AddVariable                         _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_AddVariable"           _
                          (                                       _
                            BYVAL vName       AS STRING,          _
                            BYVAL lValString  AS STRING,          _
                            BYVAL lValNumber  AS EXT   ,          _
                            BYVAL ForceType   AS LONG  ,          _      
                   Optional ByVal VarMemPtr   As Long  ,          _
                            ByVal Reserved1   As Long  ,          _
                            ByVal Reserved2   As Long  ,          _
                            ByVal Reserved3   As Long  ,          _
                            ByVal Reserved4   As Long             _
                          ) As Long

  '----------------------------------------------------------------------------
  'thinBasic_AddUdt
  '----------------------------------------------------------------------------
  ' Add a new UDT structure
  '----------------------------------------------------------------------------
  Declare Function  thinBasic_AddUdt                              _
                          LIB "thinCore.DLL"                      _
                          Alias "thinBasic_AddUdt"                _
                          (                                       _
                            ByVal sUDT_Code   As String           _
                          ) As Long

  '----------------------------------------------------------------------------
  'thinBasic_DeclareFunction
  '----------------------------------------------------------------------------
  ' Add a new function using DECLARE statement and function pointer
  '----------------------------------------------------------------------------
  Declare Function thinBasic_DeclareFunction                      _
                          Lib "thinCore.DLL"                      _
                          Alias "thinBasic_DeclareFunction"       _
                          (                                       _
                            ByVal sDeclare As String    ,         _
                            ByVal pFun As Dword                   _
                          ) As Long

  '----------------------------------------------------------------------------
  'thinBasic_ChangeVariableNumber
  '----------------------------------------------------------------------------
  ' Change the value of a numeric variable giving its name a the new value
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_ChangeVariableNumber                _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_ChangeVariableNumber"  _
                          (                                       _
                            BYVAL vName       AS STRING,          _
                            BYVAL lValNumber  AS EXT              _
                          ) AS LONG

  '----------------------------------------------------------------------------
  'thinBasic_ChangeVariableString
  '----------------------------------------------------------------------------
  ' Change the value of a string variable giving its name a the new value
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_ChangeVariableString                _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_ChangeVariableString"  _
                          (                                       _
                            BYVAL vName       AS STRING,          _
                            BYVAL lValNumber  AS STRING           _
                          ) AS LONG


  '----------------------------------------------------------------------------
  'thinBasic_GetToken
  '----------------------------------------------------------------------------
  ' Returns the next token as token ID.
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_GetToken                            _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_GetToken"              _
                          (                                       _
                            OPTIONAL BYVAL AutoPutBack AS LONG    _
                          ) AS LONG

  '----------------------------------------------------------------------------
  'thinBasic_GetToken
  '----------------------------------------------------------------------------
  ' Returns the next token as string.
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_GetTokenName                        _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_GetTokenName"          _
                          (                                       _
                            OPTIONAL BYVAL AutoPutBack AS LONG    _
                          ) AS STRING

  '----------------------------------------------------------------------------
  'thinBasic_GetTokenID
  '----------------------------------------------------------------------------
  ' Return internal engine unique ID giving the token name
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_GetTokenID                          _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_GetTokenID"            _
                          (                                       _
                            BYVAL TokenName AS String             _
                          ) AS LONG


  '----------------------------------------------------------------------------
  'thinBasic_TokenGetCurrentID
  '----------------------------------------------------------------------------
  ' Return current token ID
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_TokenGetCurrentID                   _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_TokenGetCurrentID"     _
                          () AS LONG

  '----------------------------------------------------------------------------
  'thinBasic_PutBack
  '----------------------------------------------------------------------------
  ' Put back into parsing stack the last parsed token
  '----------------------------------------------------------------------------
  DECLARE FUNCTION thinBasic_Putback                              _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_PutBack"               _
                          (                                       _
                            OPTIONAL BYVAL lPtr AS LONG           _
                          ) AS LONG

  '----------------------------------------------------------------------------
  'thinBasic_FunctionExists
  '----------------------------------------------------------------------------
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_FunctionExists                      _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_FunctionExists"        _
                          (                                       _
                            BYVAL fName       AS STRING           _
                          ) AS LONG

  '----------------------------------------------------------------------------
  'thinBasic_FunctionGetPtr
  '----------------------------------------------------------------------------
  '----------------------------------------------------------------------------
  Declare Function  thinBasic_FunctionGetPtr                      _
                          Lib "thinCore.DLL"                      _
                          Alias "thinBasic_FunctionGetPtr"        _
                          (                                       _
                            ByVal fName       As String           _
                          ) As Long

  '----------------------------------------------------------------------------
  'thinBasic_FunctionGetPtr
  '----------------------------------------------------------------------------
  '----------------------------------------------------------------------------
  Declare Function  thinBasic_FunctionGetName                     _
                          Lib "thinCore.DLL"                      _
                          Alias "thinBasic_FunctionGetName"       _
                          (                                       _
                            ByVal FunctionPtr As Long             _  '---PTR to function returned by (for example) thinBasic_FunctionGetPtr
                          ) As String


  '----------------------------------------------------------------------------
  'thinBasic_FunctionIsCallBack
  '----------------------------------------------------------------------------
  ' Determine if fucntion passed as PTR has been defined as CALLBACK function in script
  '----------------------------------------------------------------------------
  Declare Function  thinBasic_FunctionIsCallBack                  _
                          Lib "thinCore.DLL"                      _
                          Alias "thinBasic_FunctionIsCallBack"    _
                          (                                       _
                            ByVal FunctionPtr As Long             _  '---PTR to function returned by thinBasic_FunctionGetPtr
                          ) As Long


  '----------------------------------------------------------------------------
  'thinBasic_FunctionGetCBParam
  '----------------------------------------------------------------------------
  ' Returns the value of a designated CallBack pseudo variable.
  ' ATTENTION: interface function valid only whileinside a callback function       
  '----------------------------------------------------------------------------
  ' Parameters:
  '   CBParamType: one of the following equates
        '---Equates for CBParamType
        %CBParam_CBHNDL       = 100
        %CBParam_CBMSG        = 110
        %CBParam_CBCTL        = 120
        %CBParam_CBCTLMSG     = 130
        %CBParam_CBLPARAM     = 140
        %CBParam_CBWPARAM     = 150
        %CBParam_CBNMCODE     = 160
        %CBParam_CBNMHDR      = 170
        %CBParam_CBNMHWND     = 180
        %CBParam_CBNMID       = 190             
  '   lError: a pointer to a long variable that will receive possible error variables
  '           Possible errors:
  '                            -10: not inside at any function
  '                            -20: inside a function but not a callback function
  '                            -30: unsupported/invalid CBParamType
  '----------------------------------------------------------------------------
  Declare Function thinBasic_FunctionGetCBParam                     _
                          Lib "thinCore.DLL"                        _
                          Alias "thinBasic_FunctionGetCBParam"      _
                          (                                         _
                                      ByVal CBParamType As Long   , _
                            Optional  ByVal lError As Long Ptr      _
                          ) As Long

  '----------------------------------------------------------------------------
  'thinBasic_FunctionGetNumberOfParams
  '----------------------------------------------------------------------------
  ' Returns the number of parameters defined in a script function
  '----------------------------------------------------------------------------
  Declare Function  thinBasic_FunctionGetNumberOfParams           _
                          Lib "thinCore.DLL"                      _
                          Alias "thinBasic_FunctionGetNumberOfParams"     _
                          (                                       _
                            ByVal FunctionPtr As Long             _  '---PTR to function returned by thinBasic_FunctionGetPtr
                          ) As Long

  '----------------------------------------------------------------------------
  'thinBasic_FunctionGetReturnMainType
  '----------------------------------------------------------------------------
  ' Gets the main data type returned by a script function. See %VarMainType* equates above
  '----------------------------------------------------------------------------
  Declare Function  thinBasic_FunctionGetReturnMainType           _
                          Lib "thinCore.DLL"                      _
                          Alias "thinBasic_FunctionGetReturnMainType"     _
                          (                                       _
                            ByVal FunctionPtr As Long             _  '---PTR to function returned by thinBasic_FunctionGetPtr
                          ) As Long

  '----------------------------------------------------------------------------
  'thinBasic_FunctionParseAndGetPtr
  '----------------------------------------------------------------------------
  ' Parse the next token (or string expression). Check if it is a function name
  ' If yes, returns a PTR to internal data function
  '----------------------------------------------------------------------------
  Declare Function thinBasic_FunctionParseAndGetPtr                 _
                          Lib "thinCore.DLL"                        _
                          Alias "thinBasic_FunctionParseAndGetPtr"  _
                          (                                         _
                            Optional ByVal CheckIfCallBack As Long  _
                          ) As Long

  '----------------------------------------------------------------------------
  'thinBasic_FunctionSimpleCall
  '----------------------------------------------------------------------------
  ' Call a script function and optionally returns it value (numeric or string)
  '----------------------------------------------------------------------------
  Declare Function  thinBasic_FunctionSimpleCall                  _
                          Lib "thinCore.DLL"                      _
                          Alias "thinBasic_FunctionSimpleCall"    _
                          (                                       _
                            ByVal FunctionName As String,         _  '---Name of the function
                            Optional                              _
                            ByVal ptrEXT As Ext Ptr,              _  '---Used to get back from the called function a numeric value
                            ByVal ptrSTR As String Ptr            _  '---Used to get back from the called function a string value
                          ) As Long

  '----------------------------------------------------------------------------
  'thinBasic_FunctionSimpleCall_ByPtr
  '----------------------------------------------------------------------------
  ' Call a script function and optionally returns it value (numeric or string)
  ' This function permits to call script callback functions passing automatic
  ' pseudo callbacks variables 
  ' USE only when in need of a callback function from module
  '----------------------------------------------------------------------------
  Declare Function  thinBasic_FunctionSimpleCall_ByPtr                _
                          Lib "thinCore.DLL"                          _
                          Alias "thinBasic_FunctionSimpleCall_ByPtr"  _
                          (                                           _
                            ByVal FunctionPtr As Long               , _  '---PTR to function returned by thinBasic_FunctionGetPtr or other way
                            Optional                                  _
                            ByVal ptrEXT      As Ext Ptr            , _  '---Used to get back from the called function a numeric value
                            ByVal ptrSTR      As String Ptr         , _  '---Used to get back from the called function a string value
                            ByVal IsCallBack  As Long               , _
                            ByVal lCBHNDL     As Long               , _
                            ByVal lCBMSG      As Long               , _
                            ByVal lCBCTL      As Long               , _
                            ByVal lCBCTLMSG   As Long               , _
                            ByVal lCBLPARAM   As Long               , _
                            ByVal lCBWPARAM   As Long               , _                                         
                            ByVal lCBNMCODE   As Long               , _                                         
                            ByVal lCBNMHDR    As Long               , _                                         
                            ByVal lCBNMHWND   As Long               , _                                         
                            ByVal lCBNMID     As Long                 _                                         
                          ) As Long

  '----------------------------------------------------------------------------
  'thinBasic_FunctionCall_ByPtr
  '----------------------------------------------------------------------------
  ' Call a script function and optionally returns it value (numeric or string)
  ' This function permits to call script functions passing parameters
  ' Parameters are passed using a shared memory area made by a sequence of DOUBLEs
  ' (8 bytes) one for each parameter to be passed.
  ' The paramaters memory area will be interpreted by thinBasic core engine
  ' depending by parameters declaration made in script.
  '
  ' So far the following type of parameters are supported:
  '   - numeric parameters passed BYVAL up to DOUBLE range
  '
  ' Example: if programmer wants to pass 2 parameters, set lModParams = 2 and setup
  ' a 16 bytes memory area (2 params * 8 bytes each) where to store the 2 parameter
  ' values.
  '----------------------------------------------------------------------------
  Declare Function  thinBasic_FunctionCall_ByPtr                      _
                          Lib "thinCore.DLL"                          _
                          Alias "thinBasic_FunctionCall_ByPtr"        _
                          (                                           _
                            ByVal FunctionPtr   As Long             , _  '---PTR to script function
                            ByVal lModParams    As Long             , _  '---MANDATORY: number of params module is passing
                            ByVal lModParamsPtr As Long             , _  '---MANDATORY: pointer to a memory area where to find parameters (8 bytes for each parameter)
                            Optional                                  _  '---If no need to get back script function result value, do not indicate the next 2 parameters
                            ByVal ptrDOUBLE     As Double Ptr       , _  '---Pointer to DOUBLE that will get back script function return value in case of numeric
                            ByVal ptrSTR        As String Ptr         _  '---Pointer to a OLE32 dynamic string handle that will get back script function return value in case of string
                          ) As Long

  '----------------------------------------------------------------------------
  'thinBasic_GetKeywordSpeficic
  '----------------------------------------------------------------------------
  ' Instruct parser to get next token, test if it a keyword and if yes,
  ' check if it is one of the possible keywords passed. If not, a RT-Error
  ' will be generated
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_GetKeywordSpeficic                  _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_GetKeywordSpeficic"    _
                          (                                       _
                                      BYVAL HideError   AS LONG , _
                                      BYVAL AutoPutBack AS LONG , _
                                      BYVAL lKey1   AS LONG     , _
                             OPTIONAL BYVAL lKey2   AS LONG     , _
                                      BYVAL lKey3   AS LONG     , _
                                      BYVAL lKey4   AS LONG     , _
                                      BYVAL lKey5   AS LONG     , _
                                      BYVAL lKey6   AS LONG     , _
                                      BYVAL lKey7   AS LONG     , _
                                      BYVAL lKey8   AS LONG     , _
                                      BYVAL lKey9   AS LONG     , _
                                      BYVAL lKey10  AS LONG     , _
                                      BYVAL lKey11  AS LONG     , _
                                      BYVAL lKey12  AS LONG     , _
                                      BYVAL lKey13  AS LONG     , _
                                      BYVAL lKey14  AS LONG     , _
                                      BYVAL lKey15  AS LONG     , _
                                      BYVAL lKey16  AS LONG     , _
                                      BYVAL lKey17  AS LONG     , _
                                      BYVAL lKey18  AS LONG     , _
                                      BYVAL lKey19  AS LONG     , _
                                      ByVal lKey20  As Long     , _
                                      ByVal lKey21  As Long     , _
                                      ByVal lKey22  As Long     , _
                                      ByVal lKey23  As Long     , _
                                      ByVal lKey24  As Long     , _
                                      ByVal lKey25  As Long       _
                          ) As Long

  '----------------------------------------------------------------------------
  'thinBasic_VariableParse
  '----------------------------------------------------------------------------
  ' Instruct parser to get next token, check if variable, check if array and
  ' return info needed to subsequent change variable value
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_VariableParse                       _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_VariableParse"         _
                          (                                       _
                            BYREF VariablePtr     AS LONG       , _     '---ATTENTION: parameter passed BYREF will return info
                            BYREF VariableAbsPos  AS LONG         _     '---ATTENTION: parameter passed BYREF will return info
                          ) AS LONG

  '----------------------------------------------------------------------------
  'thinBasic_VariableParsePtr
  '----------------------------------------------------------------------------
  ' Instruct parser to get next token, check if variable
  ' return its pointer
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_VariableParsePtr                    _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_VariableParsePtr"      _
                          (                                       _
                            BYREF VariablePtr     AS LONG         _     '---ATTENTION: parameter passed BYREF will return info
                          ) AS LONG

  '----------------------------------------------------------------------------
  'thinBasic_ParsePtrToSomething
  '----------------------------------------------------------------------------
  ' Parse whatever and return a PTR to something
  '----------------------------------------------------------------------------
  Declare Function thinBasic_ParsePtrToSomething                    _
                          LIB "thinCore.DLL"                        _
                          Alias "thinBasic_ParsePtrToSomething"     _
                          (                                         _
                            ByRef VariablePtr As Long               _   '---ATTENTION: parameter passed BYREF will return info
                          ) As Long

  '----------------------------------------------------------------------------
  'thinBasic_VariablePtrToDirectPtr
  '----------------------------------------------------------------------------
  ' Convert a variable ptr to its internal ptr
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_VariablePtrToDirectPtr                _
                          LIB "thinCore.DLL"                        _
                          ALIAS "thinBasic_VariablePtrToDirectPtr"  _
                          (                                         _
                            BYVAL VariablePtr     AS LONG           _
                          ) AS LONG

  '----------------------------------------------------------------------------
  'thinBasic_DirectPtrToDataPtr
  '----------------------------------------------------------------------------
  ' Convert a variable ptr to its data ptr
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_DirectPtrToDataPtr                    _
                          LIB "thinCore.DLL"                        _
                          ALIAS "thinBasic_DirectPtrToDataPtr"      _
                          (                                         _
                            BYVAL DirectPtr       AS LONG           _
                          ) AS LONG

  '----------------------------------------------------------------------------
  'thinBasic_VariableParseAndGetInfo
  '----------------------------------------------------------------------------
  '
  '----------------------------------------------------------------------------
  DECLARE FUNCTION thinBasic_VariableParseAndGetInfo                _
                          LIB "thinCore.DLL"                        _
                          ALIAS "thinBasic_VariableParseAndGetInfo" _
                          (                                         _
                            BYREF VariableName    AS STRING       , _   '---ATTENTION: parameter passed BYREF will return info
                            BYREF VariablePtr     AS LONG         , _   '---ATTENTION: parameter passed BYREF will return info
                            BYREF MainType        AS LONG         , _   '---ATTENTION: parameter passed BYREF will return info
                            BYREF SubType         AS LONG         , _   '---ATTENTION: parameter passed BYREF will return info
                            BYREF DataPtr         AS LONG         , _   '---ATTENTION: parameter passed BYREF will return info
                            BYREF nElements       AS LONG           _   '---ATTENTION: parameter passed BYREF will return info
                          ) AS LONG


  '----------------------------------------------------------------------------
  'thinBasic_ParseVariableInfo
  '----------------------------------------------------------------------------
  ' Parameters (all passed BYREF ... AS LONG):
  '   - VariablePtr       pointer to internal thinBasic Core engine data structure
  '                       used to store variable info
  '   - MainType          variable main type
  '   - SubType           variable sub type
  '   - ElementsAreFixed  %TRUE if elements inside data buffer are fixed size.
  '                       For example will return %TRUE:
  '                         - fixed STRING size, the one declared ... AS STRING * size
  '                         - fixed ASCIIZ size, the one declared ... AS ASCIIZ * size
  '                         - UDT variables
  '   - TotElements       total number of elements in variable. Usually 1 if
  '                       variable is not an array
  '   - ElementSize       size of the single element. For example if variable is a
  '                       LONG, ElementSize = 4. If EXT ElementSize = 10
  '                       if variable is a UDT, it will depend from UDT to UDT.
  '   - DataPtr           pointer to variable data 
  '   - AbsPos            absolute position inside the array if variable is an array
  '                       if variable is not an array 1 is returned
  '----------------------------------------------------------------------------
  declare function thinBasic_ParseVariableInfo                      _
                          LIB "thinCore.DLL"                        _
                          alias "thinBasic_ParseVariableInfo"       _
                          (                                         _
                            byref VariablePtr       as long           , _   '---ATTENTION: parameter passed BYREF will return info
                            byref MainType          as long           , _   '---ATTENTION: parameter passed BYREF will return info
                            byref SubType           as long           , _   '---ATTENTION: parameter passed BYREF will return info
                            byref ElementsAreFixed  as long           , _   '---ATTENTION: parameter passed BYREF will return info
                            byref TotElements       as long           , _   '---ATTENTION: parameter passed BYREF will return info
                            byref ElementSize       as long           , _   '---ATTENTION: parameter passed BYREF will return info
                            byref DataPtr           as long           , _   '---ATTENTION: parameter passed BYREF will return info
                            byref AbsPos            as long             _   '---ATTENTION: parameter passed BYREF will return info
                          ) as long
  '------------------------------------------------------------------------------

'  '----------------------------------------------------------------------------
'  'thinBasic_ParseVariableInfo_EX
'  '----------------------------------------------------------------------------
'  ' Parameters (all passed BYREF ... AS LONG):
'  '   - VariablePtr       pointer to internal thinBasic Core engine data structure
'  '                       used to store variable info
'  '   - MainType          variable main type
'  '   - SubType           variable sub type
'  '   - ElementsAreFixed  %TRUE if elements inside data buffer are fixed size.
'  '                       For example will return %TRUE:
'  '                         - fixed STRING size, the one declared ... AS STRING * size
'  '                         - fixed ASCIIZ size, the one declared ... AS ASCIIZ * size
'  '                         - UDT variables
'  '   - TotElements       total number of elements in variable. Usually 1 if
'  '                       variable is not an array
'  '   - ElementSize       size of the single element. For example if variable is a
'  '                       LONG, ElementSize = 4. If EXT ElementSize = 10
'  '                       if variable is a UDT, it will depend from UDT to UDT.
'  '   - DataPtr           pointer to variable data 
'  '   - AbsPos            absolute position inside the array if variable is an array
'  '                       if variable is not an array 1 is returned
'  '----------------------------------------------------------------------------
'  Declare Function thinBasic_ParseVariableInfo_EX                         _
'                          Lib "thinCore.DLL"                              _
'                          Alias "thinBasic_ParseVariableInfo_EX"          _
'                          (                                               _
'                            ByRef VariablePtr                 As Long   , _   '---ATTENTION: parameter passed BYREF will return info
'                            ByRef MainType                    As Long   , _   '---ATTENTION: parameter passed BYREF will return info
'                            ByRef SubType                     As Long   , _   '---ATTENTION: parameter passed BYREF will return info
'                            ByRef ElementsAreFixed            As Long   , _   '---ATTENTION: parameter passed BYREF will return info
'                            ByRef TotElements                 As Long   , _   '---ATTENTION: parameter passed BYREF will return info
'                            ByRef ElementSize                 As Long   , _   '---ATTENTION: parameter passed BYREF will return info
'                            ByRef DataPtr                     As Long   , _   '---ATTENTION: parameter passed BYREF will return info
'                            ByRef AbsPos                      As Long   , _   '---ATTENTION: parameter passed BYREF will return info
'                            ByVal Param_AbsPosCanBeMissed     As Long   , _
'                            ByRef Param_nDims                 As Long   , _
'                            ByRef Param_Dim1UBound            As Long   , _
'                            ByRef Param_Dim2UBound            As Long   , _
'                            ByRef Param_Dim3UBound            As Long     _
'                          ) As Long
'  '------------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  'thinBasic_ChangeVariableStringDirect
  '----------------------------------------------------------------------------
  ' See also thinBasic_VariableParse
  '
  ' Change a variable using direct variable pointer's data returned by
  ' thinBasic_VariableParse. This will ensure to work also with arrays
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_ChangeVariableStringDirect          _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_ChangeVariableStringDirect" _
                          (                                       _
                            BYVAL VariablePtr     AS LONG       , _
                            BYVAL VariableAbsPos  AS LONG       , _
                            BYVAL lValString      AS STRING       _
                          ) AS LONG

  '----------------------------------------------------------------------------
  'thinBasic_ChangeVariableNumberDirect
  '----------------------------------------------------------------------------
  ' See also thinBasic_VariableParse
  '
  ' Change a variable using direct variable pointer's data returned by
  ' thinBasic_VariableParse. This will ensure to work also with arrays
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_ChangeVariableNumberDirect          _
                          LIB "thinCore.DLL"                      _
                          ALIAS "thinBasic_ChangeVariableNumberDirect" _
                          (                                       _
                            BYVAL VariablePtr     AS LONG       , _
                            BYVAL VariableAbsPos  AS LONG       , _
                            BYVAL lValNumber      AS EXT          _
                          ) AS LONG


  '----------------------------------------------------------------------------
  'thinBasic_GetVariableNumberDirect
  '----------------------------------------------------------------------------
  '
  ' Returns the value of a variable giving its ptr and it abs position
  '----------------------------------------------------------------------------
  Declare Function thinBasic_GetVariableNumberDirect              _
                          LIB "thinCore.DLL"                      _
                          Alias "thinBasic_GetVariableNumberDirect" _
                          (                                       _
                            ByVal VariablePtr     As Long       , _
                            ByVal VariableAbsPos  As Long         _
                          ) As Ext

  '----------------------------------------------------------------------------
  'thinBasic_ChangeVariableUDTDirect
  '----------------------------------------------------------------------------
  ' See also thinBasic_VariableParse
  '
  ' Change a variable using direct variable pointer's data returned by
  ' thinBasic_VariableParse. This will ensure to work also with arrays
  '----------------------------------------------------------------------------
  Declare Function  thinBasic_ChangeVariableUDTDirect          _
                          Lib "thinCore.DLL"                      _
                          Alias "thinBasic_ChangeVariableUDTDirect" _
                          (                                       _
                            ByVal VariablePtr     As Long       , _
                            ByVal VariableAbsPos  As Long       , _
                            ByVal lValString      As String       _
                          ) As Long


  '----------------------------------------------------------------------------
  ' NOTE
  '----------------------------------------------------------------------------
  ' The following functions
  '
  '    thinBasic_ParseNumber
  '    thinBasic_ParseString
  '    thinBasic_CheckOpenParens
  '    thinBasic_CheckCloseParens
  '    thinBasic_CheckComma
  '
  ' are used inside user defined functions/subs to parse function params.
  ' Imagine you want create a user def function like:
  '
  '    MyFunct(String1, String2, Number1)
  '
  ' To get params you have to parse the following tokens:
  '    (           Open parens
  '      String1     a string expression
  '      ,           a comma
  '      String2     a second string expression
  '      ,           another comma
  '      Number1     a numeric expression
  '    )           Finally close parens
  '
  ' thinBasic gives this ability using 5 simple functions to parse user expressions.
  ' Let's take again the above example and see how to write a user defined
  ' function that returns a STRING
  '
  ' FUNCTION MyLeft() AS STRING
  ' 'Syntax: Number = MyLeft(S1, S2, N1)
  '   LOCAL lString1 AS STRING                           'Define first string param
  '   LOCAL lString2 AS STRING                           'Define second string param
  '   LOCAL lNumber1 AS EXT                              'Define third number param
  '   IF thinBasic_CheckOpenParens THEN                  'Check if open parens token found
  '     thinBasic_ParseString lString1                   'Now evaluate a string expression
  '     IF thinBasic_CheckComma THEN                     'Check if there is a comma
  '       thinBasic_ParseString lString2                 'Parse another string expression
  '       IF thinBasic_CheckComma THEN                   'Another comma
  '         thinBasic_ParseNumber lNumber1               'Now get a numeric expression
  '         IF thinBasic_CheckCloseParens THEN           'Finally a close parens
  '           FUNCTION = LEFT$(lString1, lNumber1) + _   '...here we are. Here you can assign
  '                      LEFT$(lString2, lNumber1)       '...to your function whatever value
  '         END IF                                       '...using the evaluated params
  '       END IF
  '     END IF
  '   END IF
  ' END FUNCTION
  '
  '
  ' thinBasic_CheckOpenParens
  ' thinBasic_CheckCloseParens
  ' thinBasic_CheckComma
  ' have the following optional parameters:
  '
  '   HideHerror 
  '               if %TRUE, in case of parsing problems, no runtime error will
  '               be generated
  '   AutoPutBack
  '               if %TRUE, in case of parsing problems, the found token
  '               will be put back into parsing stack
  '
  '----------------------------------------------------------------------------
  DECLARE SUB       thinBasic_ParseNumber                 LIB "thinCore.DLL" ALIAS "thinBasic_ParseNumber"                (Result AS EXT)
  Declare Sub       thinBasic_ParseLong                   Lib "thinCore.DLL" Alias "thinBasic_ParseLong"                  (Result As Long)
  Declare Sub       thinBasic_ParseDWord                  Lib "thinCore.DLL" Alias "thinBasic_ParseDWord"                 (Result As Dword)
  Declare Function  thinBasic_ParseLong_Optional          Lib "thinCore.DLL" Alias "thinBasic_ParseLong"                  (Result As Long) As Long
  '---

  DECLARE FUNCTION  thinBasic_Parse1Number                LIB "thinCore.DLL" ALIAS "thinBasic_Parse1Number"               () AS EXT
  DECLARE FUNCTION  thinBasic_Parse2Numbers               LIB "thinCore.DLL" ALIAS "thinBasic_Parse2Numbers"              (Num1 AS EXT, Num2 AS EXT) AS EXT
  DECLARE FUNCTION  thinBasic_Parse3Numbers               LIB "thinCore.DLL" ALIAS "thinBasic_Parse3Numbers"              (Num1 AS EXT, Num2 AS EXT, Num3 AS EXT) AS EXT
  DECLARE FUNCTION  thinBasic_Parse4Numbers               LIB "thinCore.DLL" ALIAS "thinBasic_Parse4Numbers"              (Num1 AS EXT, Num2 AS EXT, Num3 AS EXT, Num4 AS EXT) AS EXT
  DECLARE FUNCTION  thinBasic_Parse5Numbers               LIB "thinCore.DLL" ALIAS "thinBasic_Parse5Numbers"              (Num1 AS EXT, Num2 AS EXT, Num3 AS EXT, Num4 AS EXT, Num5 AS EXT) AS EXT
  DECLARE FUNCTION  thinBasic_Parse6Numbers               LIB "thinCore.DLL" ALIAS "thinBasic_Parse6Numbers"              (Num1 AS EXT, Num2 AS EXT, Num3 AS EXT, Num4 AS EXT, Num5 AS EXT, Num6 AS EXT) AS EXT
  declare function  thinBasic_Parse7Numbers               lib "thinCore.DLL" alias "thinBasic_Parse7Numbers"              (Num1 as ext, Num2 as ext, Num3 as ext, Num4 as ext, Num5 as ext, Num6 as ext, Num7 as ext) as ext
  '---

  Declare Function  thinBasic_Parse2Numbers1String        Lib "thinCore.DLL" Alias "thinBasic_Parse3Numbers1String"       (Num1 As Ext, Num2 As Ext, s1 As String) As Ext
  Declare Function  thinBasic_Parse3Numbers1String        Lib "thinCore.DLL" Alias "thinBasic_Parse3Numbers1String"       (Num1 As Ext, Num2 As Ext, Num3 As Ext, s1 As String) As Ext
  '---

  Declare Function  thinBasic_ParseXNumbers               Lib "thinCore.DLL" _
                                                          alias "thinBasic_ParseXNumbers" _
                                                          ( _
                                                            BYVAL MinNumToParse AS LONG, BYVAL MaxNumToParse AS LONG, _
                                                                      Num1  AS EXT    , _
                                                            OPTIONAL  Num2  AS EXT    , _
                                                                      Num3  AS EXT    , _
                                                                      Num4  AS EXT    , _
                                                                      Num5  AS EXT    , _
                                                                      Num6  AS EXT    , _
                                                                      Num7  AS EXT    , _
                                                                      Num8  AS EXT    , _
                                                                      Num9  AS EXT    , _
                                                                      Num10 AS EXT      _
                                                          ) as ext
  '---
  declare function  thinBasic_Parse1StringXNumbers        lib "thinCore.DLL" _
                                                          alias "thinBasic_Parse1StringXNumbers" _
                                                          ( _
                                                            BYVAL MinNumToParse AS LONG, BYVAL MaxNumToParse AS LONG, _
                                                                      Str1  AS STRING , _
                                                                      Num1  AS EXT    , _
                                                            OPTIONAL  Num2  AS EXT    , _
                                                                      Num3  AS EXT    , _
                                                                      Num4  AS EXT    , _
                                                                      Num5  AS EXT    , _
                                                                      Num6  AS EXT    , _
                                                                      Num7  AS EXT    , _
                                                                      Num8  AS EXT    , _
                                                                      Num9  AS EXT    , _
                                                                      Num10 AS EXT      _
                                                          ) as ext

  '---
  ' Parse one string expression including possible "(" and ")" before and after the string
  ' This function is the equavalent of the numeric thinBasic_Parse1Number 
  Declare Function  thinBasic_Parse1String                Lib "thinCore.DLL" Alias "thinBasic_Parse1String"               () As String

  '---
  ' Parse one string expression. It also returns the numeric representation of the parsed string in case it represents a valid number
  Declare Function  thinBasic_ParseString                 Lib "thinCore.DLL" Alias "thinBasic_ParseString"                (ByRef sResult As String) As Ext

  '---
  ' Parse one string expression.
  Declare Sub       thinBasic_ParseStr                    Lib "thinCore.DLL" Alias "thinBasic_ParseStr"                   (ByRef sResult As String)

  '---
  ' Parse a variant variable (scalar or array) and returns its pointer
  Declare Function  thinBasic_ParseVariant                Lib "thinCore.DLL" Alias "thinBasic_ParseVariant"               () As Dword


  DECLARE FUNCTION  thinBasic_CheckOpenParens             LIB "thinCore.DLL" ALIAS "thinBasic_CheckOpenParens"            (OPTIONAL BYVAL HideError AS LONG, BYVAL AutoPutBack AS LONG) AS LONG
  DECLARE FUNCTION  thinBasic_CheckOpenParens_Mandatory   LIB "thinCore.DLL" ALIAS "thinBasic_CheckOpenParens_Mandatory"  () AS LONG
  DECLARE FUNCTION  thinBasic_CheckOpenParens_Optional    LIB "thinCore.DLL" ALIAS "thinBasic_CheckOpenParens_Optional"   () AS LONG

  DECLARE FUNCTION  thinBasic_CheckCloseParens            LIB "thinCore.DLL" ALIAS "thinBasic_CheckCloseParens"           (OPTIONAL BYVAL HideError AS LONG, BYVAL AutoPutBack AS LONG) AS LONG
  DECLARE FUNCTION  thinBasic_CheckCloseParens_Mandatory  LIB "thinCore.DLL" ALIAS "thinBasic_CheckCloseParens_Mandatory" () AS LONG
  DECLARE FUNCTION  thinBasic_CheckCloseParens_Optional   LIB "thinCore.DLL" ALIAS "thinBasic_CheckCloseParens_Optional"  () AS LONG

  DECLARE FUNCTION  thinBasic_CheckComma                  LIB "thinCore.DLL" ALIAS "thinBasic_CheckComma"                 (OPTIONAL BYVAL HideError AS LONG, BYVAL AutoPutBack AS LONG) AS LONG
  DECLARE FUNCTION  thinBasic_CheckComma_Mandatory        LIB "thinCore.DLL" ALIAS "thinBasic_CheckComma_Mandatory"       () AS LONG
  DECLARE FUNCTION  thinBasic_CheckComma_Optional         LIB "thinCore.DLL" ALIAS "thinBasic_CheckComma_Optional"        () AS LONG

  Declare Function  thinBasic_CheckSemicolon_Mandatory    Lib "thinCore.DLL" Alias "thinBasic_CheckSemicolon_Mandatory"   () As Long
  Declare Function  thinBasic_CheckSemicolon_Optional     Lib "thinCore.DLL" Alias "thinBasic_CheckSemicolon_Optional"    () As Long

  Declare Function  thinBasic_CheckEqual_Mandatory        Lib "thinCore.DLL" Alias "thinBasic_CheckEqual_Mandatory"       () As Long
  Declare Function  thinBasic_CheckEqual_Optional         Lib "thinCore.DLL" Alias "thinBasic_CheckEqual_Optional"        () As Long

  Declare Function  thinBasic_CheckPoint_Mandatory        Lib "thinCore.DLL" Alias "thinBasic_CheckPoint_Mandatory"       () As Long
  Declare Function  thinBasic_CheckPoint_Optional         Lib "thinCore.DLL" Alias "thinBasic_CheckPoint_Optional"        () As Long

  Declare Function  thinBasic_CheckPlus_Optional          Lib "thinCore.DLL" Alias "thinBasic_CheckPlus_Optional"         () As Long
  Declare Function  thinBasic_CheckMinus_Optional         Lib "thinCore.DLL" Alias "thinBasic_CheckMinus_Optional"        () As Long
  Declare Function  thinBasic_CheckMult_Optional          Lib "thinCore.DLL" Alias "thinBasic_CheckMult_Optional"         () As Long
  Declare Function  thinBasic_CheckDiv_Optional           Lib "thinCore.DLL" Alias "thinBasic_CheckDiv_Optional"          () As Long

  Declare Function  thinBasic_CheckEOL_Optional           Lib "thinCore.DLL" Alias "thinBasic_CheckEOL_Optional"          () As Long


  Declare Function  thinBasic_GetUnknownToken             Lib "thinCore.DLL" Alias "thinBasic_GetUnknownToken"            (Optional ByVal AutoPutBack As Long) As String
  '----------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  Declare Function  thinBasic_CheckEqualType_Mandatory    Lib "thinCore.DLL" Alias "thinBasic_CheckEqualType_Mandatory"   () As Long
    %Equal_EQ             = 1520   ' =
    %Equal_PLUSEQUAL      = 1521   ' +=
    %Equal_MINUSEQUAL     = 1522   ' -=
    %Equal_MULTIPLYEQUAL  = 1523   ' *=
    %Equal_DIVIDEEQUAL    = 1524   ' /=
    %Equal_IDIVIDEEQUAL   = 1525   ' \=
    %Equal_CONCEQUAL      = 1526   ' =
    %Equal_POINTEQUAL     = 1527   ' .=
  '----------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  Declare Function  thinBasic_DetermineType               Lib "thinCore.DLL" Alias "thinBasic_DetermineType"              () As Long
  '----------------------------------------------------------------------------
  ' Possible return values:
  '   %thinBasic_ReturnNumber
  '   %thinBasic_ReturnString
  '----------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  'thinBasic_GetRunTimeInfo
  '----------------------------------------------------------------------------
  ' The following information can be taken:
  '  "SCRIPT_PATH"  Path to script
  '  "APP_PATH"     Path to thinBasic
  '  "HWND"         Handle to main thinBasic window
  '----------------------------------------------------------------------------
  DECLARE FUNCTION  thinBasic_GetRunTimeInfo    LIB "thinCore.DLL" ALIAS "thinBasic_GetRunTimeInfo"   (BYVAL InfoName AS STRING) AS STRING
  '----------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  'thinBasic_ScriptIsObuscated
  '----------------------------------------------------------------------------
  '----------------------------------------------------------------------------
  Declare Function  thinBasic_ScriptIsObfuscated    Lib "thinCore.DLL" Alias "thinBasic_ScriptIsObfuscated"   () As Long
  '----------------------------------------------------------------------------


  '----------------------------------------------------------------------------
  'Trace Handling
  '----------------------------------------------------------------------------
  '----------------------------------------------------------------------------
  ' TRACE Equates
    '---What to ask for
      %TRACE_What_SourcePtr                       =   1
      %TRACE_What_StackLevel                      =  10
      %TRACE_What_IFLevel                         =  20
      %TRACE_What_SourceFullPath                  =  30
  
    '---Error codes
      %TRACE_Error_NotPossible_Obfuscated         =   1
      
      %TRACE_Error_InstallHandle_Null             = 200
      %TRACE_Error_InstallHandle_AlreadyInstalled = 201

  '----------------------------------------------------------------------------
  'thinBasic_ITrace_GetValue
  '----------------------------------------------------------------------------
  ' 
  '  
  '  
  '  
  '----------------------------------------------------------------------------
    Declare Function thinBasic_ITrace_GetValue                      _
                            Lib   "thinCore.dll"                    _
                            Alias "thinBasic_ITrace_GetValue"       _
                            (                                       _
                                        ByVal lWhat As Long       , _
                              Optional  ByVal sWhat As String     , _
                                        ByVal lRet1 As Long Ptr   , _
                                        ByVal lRet2 As Long Ptr   , _
                                        ByVal sRet1 As String Ptr   _
                            ) As Long

  '----------------------------------------------------------------------------
  'thinBasic_ITrace_InstallHandle
  '----------------------------------------------------------------------------
  ' 
  ' 
  '
  '
  '----------------------------------------------------------------------------
    Declare Function thinBasic_ITrace_InstallHandle                 _
                            Lib   "thinCore.dll"                    _
                            Alias "thinBasic_ITrace_InstallHandle"  _
                            (                                       _
                                        ByVal fHandle As Long       _
                            ) As Long
  '----------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  'thinBasic_StackGetCurrent
  '----------------------------------------------------------------------------
  ' Returns the current script stack level.
  ' 1 = Global, 2 = next function executed, 3 = ...
  '----------------------------------------------------------------------------
    Declare Function thinBasic_StackGetCurrent                      _
                            Lib   "thinCore.dll"                    _
                            Alias "thinBasic_StackGetCurrent"       _
                            (                                       _
                            ) As Long
  '------------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  'thinBasic_StackGetList
  '----------------------------------------------------------------------------
  ' Returns a string containing the names of the current nesting stack levels
  ' As a minimum this function will return one string containing "GLOBAL"
  '----------------------------------------------------------------------------
    Declare Function thinBasic_StackGetList                         _
                            Lib   "thinCore.dll"                    _
                            Alias "thinBasic_StackGetList"          _
                            (                                       _
                              Optional ByVal sSep As String         _
                            ) As String
  '----------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  'ERROR Handling
  '----------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  'thinBasic_GetLastError
  '----------------------------------------------------------------------------
  ' Returns the last occurred error code
  '----------------------------------------------------------------------------
  Declare Function  thinBasic_GetLastError      Lib "thinCore.DLL" Alias "thinBasic_GetLastError"        () As Long

  '----------------------------------------------------------------------------
  'thinBasic_ErrorFlag
  '----------------------------------------------------------------------------
  ' Returns a value different from 0 if a runtime error occurred
  '----------------------------------------------------------------------------
  Declare Function  thinBasic_ErrorFlag         Lib "thinCore.DLL" Alias "thinBasic_ErrorFlag"        () As Long

  '----------------------------------------------------------------------------
  'thinBasic_ErrorFree
  '----------------------------------------------------------------------------
  ' Returns -1 if there thinBasic is not under runtime error state
  '----------------------------------------------------------------------------
  Declare Function  thinBasic_ErrorFree         Lib "thinCore.DLL" Alias "thinBasic_ErrorFree"        () As Long
  '----------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  'thinBasic_ErrorUserConfirmed
  '----------------------------------------------------------------------------
  ' Returns -1 if, after an error, user has confirmed the dialog
  '----------------------------------------------------------------------------
  Declare Function  thinBasic_ErrorUserConfirmed         Lib "thinCore.DLL" Alias "thinBasic_ErrorUserConfirmed"        () As Long
  '----------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  'thinBasic_RunTimeError
  '----------------------------------------------------------------------------
  ' Use internal thinCore engine to generate a runtime error
  '  ErrorCode      Error code. Get one from the below list
  '  AdditionalInfo Personalized text to be shown into runtime error dialog
  '----------------------------------------------------------------------------
  Declare Function thinBasic_RunTimeError                               _
                          Lib "thinCore.dll"                            _
                          Alias "thinBasic_RunTimeError"                _
                          ( _
                                      ByVal ErrorCode       As Long   , _
                             Optional ByVal AdditionalInfo  As String   _
                          ) As Long

  '------------------------------------------------------------------------------
  ' RunTime Errors Codes
  '------------------------------------------------------------------------------
  %ERR__PARENS                                =   1
  %ERR__NOEXP                                 =   2
  %ERR__DIV_ZERO                              =   3
  %ERR__EQUAL_EXP                             =   4
  %ERR__NOT_VAR                               =   5
  %ERR__LAB_TAB_FULL                          =   6
  %ERR__DUP_LAB                               =   7
  %ERR__UNDEF_LAB                             =   8
  %ERR__THEN_EXPECTED                         =   9
  %ERR__TO_EXPECTED                           =  10
  %ERR__TOO_MANY_FOR                          =  11
  %ERR__NEXT_WITHOUT_FOR                      =  12
  %ERR__MISSING_SEMICOLON                     =  13
                                              
  %ERR__COMMAND_RET_NONE_IN_EXPRESSION        =  14
                                              
  %ERR__MISS_QUOTE                            =  15
  %ERR__BAD_FILE                              =  16
  %ERR__STR_EXPECTED                          =  17
  %ERR__UNKNOWN_KEYWORD                       =  18
  %ERR__MISSING_OPENPARENS                    =  19
  %ERR__MISSING_CLOSEPARENS                   =  20
  %ERR__MISSING_COMMA                         =  21
  %ERR__MISSING_SQUARECLOSEPARENS             =  22
                                              
  %ERR__EOL_INCORRECT                         =  23
                                              
  %ERR__DOUBLE_CONCATENATION                  =  24
  %ERR__COMMAND_UNKNOWN_RETURNED_PARAMETER    =  25
  %ERR__PRIMITIVE_STR                         =  26
  %ERR__PRIMITIVE_NUM                         =  27
  %ERR__PRINT_ERROR                           =  28
  %ERR__PRINT_ERROR_NOENDOFLINE               =  29
  %ERR__VARIABLE_NOT_DEFINED                  =  30
  %ERR__ATOM_TOKEN_TYPE_NUMERIC               =  31
  %ERR__ATOM_NUM                              =  31  '---Used as an abbreviation in macros
  %ERR__IF_WITHOUT_ENDIF                      =  32
  %ERR__IF_ENDIF_WITHOUT_IF                   =  33
  %ERR__TOO_MANY_WHILE                        =  34
  %ERR__WHILE_WITHOUT_WEND                    =  35
  %ERR__DIM_TYPE_NOT_DEFINED                  =  36
  %ERR__EXIT_NO_KEY_FOUND                     =  37
  %ERR__NO_END_FUNCTION_FOUND                 =  38
  %ERR__FUNCTION_NAME_DUPLICATE               =  39
  %ERR__FUNCTION_NAME_IS_KEY                  =  40
  %ERR__FUNCTION_NAME_IS_VAR                  =  41
  %ERR__FUNCTION_NAME_IS_UDT                  =  42
  %ERR__END_NO_KEY_FOUND                      =  43
                                              
  %ERR__DIM_MISSING_AS                        =  44
  %ERR__UNDEFINED_TOKEN                       =  45
  %ERR__UNSUPPORTED_CHAR                      =  46
                                              
  %ERR__FUNCTION_DECLARE_MISSING_AS           =  47
  %ERR__DECLARE_FUNCTION_OR_SUB_EXPECTED      =  48
                                              
  %ERR__DOLOOP_WRONG_CONDITION                =  49
                                              
  %ERR__MISSING_ALIAS                         =  50
  %ERR__MISSING_LIB                           =  51
  %ERR__UNDEFINED_VAR_TYPE                    =  52
  %ERR__PARAMETER_BYREF_INCORRECT_PTR         =  53
                                              
  %ERR__LOOP_EXPECTED_WHILE_OR_UNTIL          =  55

  %ERR__END_FUNCTION_WITHOUT_FUNCTION         =  58
  %ERR__END_FUNCTION_FOUND                    =  59
                                              
  %ERR__ITERATE_NO_KEY_FOUND                  =  60
                                              
  %ERR__WITH_ENDWITH_ERROR                    =  64
                                              
  %ERR__TOO_NESTED_DOLOOP                     =  65
  %ERR__DOLOOP_MISSING_DO_OR_LOOP             =  66
                                              
  %ERR__VARIABLE_NOT_DIMENSIONED              =  70
  %ERR__VARIABLE_ISNOT_ARRAY                  =  71
  %ERR__VARIABLE_MUSTBE_STRING_TYPE           =  72
                                              
  %ERR__REDIM_NEW_TYPE_NOT_SUPPORTED          =  73
  %ERR__REDIM_PRESERVE_NOTVALID_ABSOLUTE      =  74
                                              
  %ERR__VARIABLE_MUSTBE_UDT_TYPE              =  75
                                              
  %ERR__KEYWORD_NOT_EXPECTED                  =  76
  %ERR__TOKEN_NOT_EXPECTED                    =  77
                                              
  %ERR__INCLUDE_FILE_NOT_FOUND                =  80

  %ERR__DIM_UNEXPECTED_KEYWORD                =  85
                                              
  %ERR__FUNCTION_NOT_SUPPORTED                =  90
                                              
  %ERR__ARRAY_FUNCTION_NOT_SUPPORTED          =  91
                                              
  %ERR__UDT_ELEMENT_NOT_FOUND                 = 100
  %ERR__UDT_EXPECTED                          = 101
  %ERR__UDT_EQU_OR_ELEMENT_EXPECTED           = 102
                                              
  %ERR__ASSIGNMENT_NOT_SUPPORTED              = 110
                                              
  %ERR__RELATIONAL_EXPECTED                   = 115
                                              
  %ERR__API_LIB_NOT_FOUND                     = 120
  %ERR__API_FUNCTION_NOT_FOUND_IN_LIB         = 121
  %ERR__API_GENERAL_ADDRESS_NOT_PRESENT       = 122
                                              
  %ERR__CALL_NOTSUPPORTEDSTATEMENT            = 130
  %ERR__CALL_FUNCTIONNOTFOUND                 = 131
                                              
  %ERR__FUNCTIONNOTFOUND                      = 133
  %ERR__FUNCTION_EXPECTED_CALLBACK            = 134

  %ERR__EQUATE_ALREADY_DEFINED                = 135  
  %ERR__EQUATE_ALREADY_DEFINED_DIFFERENT      = 136  
                                              
  %ERR__VARIABLE_NAME_DUPLICATE_GLOBAL        = 142
  %ERR__VARIABLE_NAME_DUPLICATE_FUNCTION      = 144
  %ERR__VARIABLE_NAME_DUPLICATE_LOCAL         = 145
  %ERR__VARIABLE_NAME_DUPLICATE               = 146
                                              
  %ERR__FOR_StepShouldBeNegative              = 150
  %ERR__FOR_StepShouldBePositive              = 151
  %ERR__FOR_ExpectedAVariable                 = 152
  %ERR__FOR_StepIsZero                        = 153
  %ERR__FOR_VarMustBeNumeric                  = 154
                                              
  %ERR__ALIAS_CommandNameExpected             = 160
  %ERR__ALIAS_AsExpected                      = 161
  %ERR__ALIAS_UndefNotUndef                   = 162
                                              
  %ERR__TYPE_MISSING_END_UNION                = 169
  %ERR__TYPE_MISSING_END_TYPE                 = 170
  %ERR__TYPE_MISSING_END_CLASS                = 171
  %ERR__TYPE_TYPE_NOT_DEFINED                 = 172
  %ERR__TYPE_MISSING_AS                       = 173
  %ERR__TYPE_NAME_MUSTBE_UNDEFINED            = 174
  %ERR__TYPE_ARRAY_MUSTBE_DIMENSIONED         = 175
  %ERR__TYPE_STRINGS_MUST_HAVE_SIZE           = 176
  %ERR__TYPE_ELEMENT_ALREADY_PRESENT          = 177
  %ERR__TYPE_ELEMENT_ALREADY_PRESENT_INHERIT  = 178
  %ERR__TYPE_DYNSTRING_INSIDE_UNION           = 179

  %ERR__NO_END_RAWTEXT_FOUND                  = 180
                                              
  %ERR__BEGIN_BLOCK_UNSUPORTED                = 190
  %ERR__BEGIN_CONST_MISSING_END               = 191
                                              
  %ERR__FUNCTION_PARAM_UNRECOGNIZED_TYPE      = 220
                                              
  %ERR__DO_WITHOUT_LOOP                       = 245
  %ERR__FUNCTION_MISSING_AS                   = 246
  %ERR__REGEXPR_MISSING_TO                    = 247
  %ERR__REGEXPR_MISSING_IN                    = 248
  %ERR__SELECT_WITHOUT_ENDSELECT              = 249
  %ERR__DUPLICATE_SYMBOL                      = 250
  %ERR__INVALIDNUMERICCHAR                    = 251
  %ERR__INVALIDDELIMITER                      = 252
  %ERR__INVALIDDATATYPE                       = 253
  %ERR__VARIABLE_EXPECTED                     = 254 
  %ERR__VARIABLE_VARIANT_EXPECTED             = 255 
                                              
  %ERR__SELECT_MISSING_CASE                   = 270
  %ERR__SELECT_ERROR_KINDOF_OPERATION         = 271
  %ERR__SELECT_CODE_BETWEEN_SELECT_CASE       = 272
                                              
  %ERR__STRPTR_VARIABLE_NOT_A_DYNSTRING_NUM   = 280
  %ERR__STRPTR_VARIABLE_NOT_A_DYNSTRING_VAR   = 281
  %ERR__STRPTR_VARIABLE_NOT_A_DYNSTRING_UDT   = 282
                                              
  %ERR__APICALL_REF_EXPECTED                  = 300

  %ERR__ARRAY_OUT_OF_BOUND                    = 400
                                              
  %ERR__MODULE_SPECIFIC                       = 500
                                              
  %ERR__PREPARSER_DirectiveNotSupported       = 800
                                              
  %ERR__PREPARSER_ScriptVersionRequest        = 820
                                              
  %ERR__INTERNAL_RETURNMAINTYPE               = 900
                                              
  %ERR__INTERNAL_DECRIPTION                   = 910 
                                              
  %ERR__INTERNAL_UDTBUFFER_SHORT              = 915 
                                              
  %ERR__INTERNAL_RETURNNONE_NOCODEPTR         = 921 
  %ERR__INTERNAL_RETURNNUMBER_NOCODEPTR       = 922 
  %ERR__INTERNAL_RETURNSTRING_NOCODEPTR       = 923 
                                              
  %ERR__CLASS_NEW_NOINDEXALLOWED              = 5010
  %ERR__CLASS_NEW_DIFFERENTCLASS              = 5015
  %ERR__CLASS_NEW_NOCLASS                     = 5020
  %ERR__CLASS_NEW_EXPECTED_NEW                = 5025
  %ERR__CLASS_NOT_INIT_WITH_NEW               = 5030
  %ERR__CLASS_METHODPROPERTY_NOTFOUND         = 5100

  %ERR__TRACE_STOP_BY_USER                    = 11000 
                                              
  %ERR__OBFUSCATION_FILENOTVALID              = 12000
                                              
  %ERR__COM_GENERIC                           = 30000

  

  '----------------------------------------------------------------------------
  'thinBasic_Class_Add
  '----------------------------------------------------------------------------
  ' Add a new class to thinBasic Core Engine
  '----------------------------------------------------------------------------
    Declare Function thinBasic_Class_Add                                _
                            Lib   "thinCore.dll"                        _
                            Alias "thinBasic_Class_Add"                 _
                            (                                           _
                                ByVal sClassName            As String , _     '---Name of the Class to be created
                                ByVal pClassFunc            As Long     _     '---Pointer to a class function that will handle calls to methods
                            ) As Long
  '----------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  'thinBasic_Class_AddMethod
  '----------------------------------------------------------------------------
  ' Add a method to a class previously created with thinBasic_Class_Add
  '----------------------------------------------------------------------------
    Declare Function thinBasic_Class_AddMethod                          _
                            Lib   "thinCore.dll"                        _
                            Alias "thinBasic_Class_AddMethod"           _
                            (                                           _
                                ByVal pClass                As Long   , _     '---Poiter to a class. This is the value returned by thinBasic_Class_Add
                                ByVal sMethodName           As String , _     '---Name of the Class to be created
                                ByVal MethodReturnType      As Long   , _
                                ByVal pMethodFunc           As Long     _     '---Pointer to a method function that will handle method execution
                            ) As Long
  '----------------------------------------------------------------------------

  %Class_Action_None          = 0?
  %Class_Action_Get           = 1?
  %Class_Action_Set           = 2?
  
  %Class_Action_Constructor   = 100?
  %Class_Action_Destructor    = 110?

  '----------------------------------------------------------------------------
  'thinBasic_Class_AddProperty
  '----------------------------------------------------------------------------
  ' Add a property to a class previously created with thinBasic_Class_Add
  '----------------------------------------------------------------------------
    Declare Function thinBasic_Class_AddProperty  _
                            Lib   "thinCore.dll"                        _
                            Alias "thinBasic_Class_AddProperty"         _
                            (                                           _
                                ByVal pClass                As Long   , _     '---Poiter to a class. This is the value returned by thinBasic_Class_Add 
                                ByVal sPropertyName         As String , _     '---Name of the Class to be created                                      
                                ByVal PropertyReturnType    As Long   , _                                                                            
                                ByVal PtrToPropertyFunction As Long     _     '---Pointer to a method function that will handle method execution       
                             ) As Long
 
     