Attribute VB_Name = "modVRNUDF"
Option Explicit

Function IsValidReg(s As String) As Boolean
Static RX As Object
' Valid VRN Function

' Place in a standard module, usage =IsValidReg(A2)

Dim tests, I As Long
  
  tests = Array( _
     "^[A-Z]{2}(5[1-9]|0[2-9]|6[0-9]|1[0-9])[A-Z]{3}$", _
     "^[A-HJ-NP-Y]\d{1,3}[A-Z]{3}$", _
     "^[A-Z]{3}\d{1,3}[A-HJ-NP-Y]$", _
     "^(?:[A-Z]{1,2}\d{1,4}|[A-Z]{3}\d{1,3})$", _
     "^(?:\d{1,4}[A-Z]{1,2}|\d{1,3}[A-Z]{3})$", _
     "(?:[A-Z]{1,3}[0-9]{1,4}$)")
       
  If RX Is Nothing Then
    Set RX = CreateObject("VBScript.RegExp")
  End If
  
  For I = LBound(tests) To UBound(tests)
    RX.Pattern = tests(I)
    If RX.test(s) Then
        IsValidReg = True
        Exit Function
    End If
  Next I

End Function

