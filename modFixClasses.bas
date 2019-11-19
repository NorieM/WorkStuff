Attribute VB_Name = "modFixClasses"
Option Explicit

Function ReplaceClassName(ByVal arrVals As Variant, strMatch As String, strReplace As String) As Variant
Dim idx As Long

    For idx = LBound(arrVals) To UBound(arrVals)
        If arrVals(idx, 1) = strMatch Then
            arrVals(idx, 1) = strReplace
        End If
    Next idx

    ReplaceClassName = arrVals
    
End Function


Sub FixClassNames()
Dim wbSite As Workbook
Dim wsData As Worksheet
Dim rngClasses As Range
Dim arrClasses As Variant

    For Each wbSite In Application.Workbooks
                                    
        If Not wbSite Is ThisWorkbook Then
            Set wsData = wbSite.Sheets("Data")
            
            Set rngClasses = wsData.Range("A1").CurrentRegion.Columns(6)
            
            arrClasses = rngClasses.Value
            
            Debug.Print arrClasses(1, 1)
            
            arrClasses = ReplaceClassName(arrClasses, "MC", "M/C")

            arrClasses = ReplaceClassName(arrClasses, "PC", "P/C")

            rngClasses.Value = arrClasses
            
             wbSite.Close SaveChanges:=True
            
        End If
                                    
    Next wbSite
    
End Sub

