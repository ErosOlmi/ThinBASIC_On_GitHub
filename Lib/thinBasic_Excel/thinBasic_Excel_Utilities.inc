' Cnvert a VT_DATE to a dynamic string
' Variant time is "ccyymmddpart.timepart" so if VARIANT# has a FRAC part it includes
' a time component as well as a date component.
' Tested format of time component 11/5/08
' official ODBC timestamp format is ccyy-mm-dd hh:mm:ss
Function Excel_DateFromVariantDate (vDate As VARIANT, Optional ByVal AddTime As Byte) As String

    Local ST As SYSTEMTIME
    Local szDF As Asciiz * 48,  szDate As Asciiz * 48
    Local szTF As Asciiz * 48,  szTime As Asciiz * 48
    Local vbTime As Double

    vbTime = Variant#(vDate)
    VariantTimeToSystemTime vbTime, St
    ' always format the date
    szDf          =  "yyyy'-'MM'-'dd"
    GetDateFormat    _
                      ByVal 0&, _
                      ByVal 0&, _
                      st, _
                      szDf, _
                      szDate, _
                      SizeOf(szDate)

    ' only add the time string if present
    If Frac (vbTime) And AddTime Then
         szTF     = "HH':'mm':'ss"  ' HH= 24 hour, hh= 12 hour
         GetTimeFormat    _
                          ByVal 0&, _
                          ByVal 0&, _
                          st, _
                          szTF, _
                          szTime, _
                          SizeOf (szTime)
         Function = szDate & $Spc & szTime
     Else
        Function      = szDate
     End If
End Function
