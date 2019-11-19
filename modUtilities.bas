Attribute VB_Name = "modUtilities"
Option Explicit

Function CreateStopSchedule(arrStops As Variant, lngNoShifts As Long) As Variant
Dim arrScheduleStops As Variant
Dim arrSet As Variant
Dim rng As Range
Dim idxShift As Long
Dim idxStop As Long
Dim cnt As Long

    ReDim arrScheduleStops(1 To lngNoShifts * 6)
    
    For idxShift = 1 To lngNoShifts
    
        arrSet = CreateRandomSet(arrStops, 6)
            
        For idxStop = LBound(arrSet) To UBound(arrSet)
            cnt = cnt + 1
            arrScheduleStops(cnt) = arrSet(idxStop)
        Next idxStop
        
    Next idxShift
    
    CreateStopSchedule = arrScheduleStops
    
End Function

Function GetConsecutiveShift(ByVal dicShifts As Object, strHours As Variant) As String
Dim lngStart As Long
Dim lngEnd As Long
Dim ky As Variant

    lngStart = Val(Split(strHours, "-")(0))
    
    lngEnd = Val(Split(strHours, "-")(1))
    
    For Each ky In dicShifts.Keys
        If InStr(ky, "W") Then
            If Val(Split(dicShifts(ky), "-")(0)) >= lngEnd Or Val(Split(dicShifts(ky), "-")(1)) < -lngEnd Then
                GetConsecutiveShift = ky
                Exit Function
            End If
        End If
    Next ky
    
End Function

Function GetPlan(ByVal arrAvailableShifts As Variant, dtStart As Date, lngNoWeeks As Long) As Variant
Dim dicPlan As Object
Dim dicWeeks As Object
Dim dtDay As Date
Dim dtEnd As Date
Dim dtRand As Date
Dim strLine As String
Dim arrShifts As Variant
Dim dicShifts As Object
Dim strShift As String
Dim idx As Long
Dim Res As Variant
Dim boolAddShifts As Boolean
Dim boolWeekDay As Boolean
Dim counter As Long
Dim lngWeek As Long

    ' creates a plan using based on the available shifts and certain rules/guidelines
    ' plan begins on start dates and goes on for lngNoWeeks
    ' Rules/guidelines
    ' - weekdays, maximum 2 shifts per day, preferably consecutive
    ' - no weekend shifts on last weekend
    ' - no more than 8 shifts per week
    
    ' get the end date of the plan
    dtEnd = dtStart + lngNoWeeks * 7 - 3
        
        
    Set dicPlan = CreateObject("Scripting.Dictionary")
    
    Set dicShifts = CreateObject("Scripting.Dictionary")
    
    If TypeName(arrAvailableShifts) = "Range" Then arrAvailableShifts = arrAvailableShifts.Value
    
    For idx = LBound(arrAvailableShifts) To UBound(arrAvailableShifts)
        dicShifts(arrAvailableShifts(idx, 1)) = arrAvailableShifts(idx, 2)
    Next idx
    
    Set dicWeeks = CreateObject("Scripting.Dictionary")
    
    For idx = Application.WeekNum(dtStart, 2) To Application.WeekNum(dtEnd, 2)
        dicWeeks(idx) = 0
    Next idx
            
    Do
    
        Randomize Time
        
        boolAddShifts = False
        
        boolWeekDay = False
        
        ' generate random date between start and end inclusive
        dtRand = Application.RandBetween(dtStart, dtEnd)
        
        ' calculate week for date
        lngWeek = Application.WeekNum(dtRand, 2)
        
        ' randomly selected line to use
        strLine = IIf(Rnd > 0.5, "G", "R")
        
        ' check if date has been already been allocated shifts
        If Not dicPlan.Exists(dtRand) Then
        
            ReDim arrShifts(1 To 2, 1 To 2)
            
            ' check if date weeday, or Saturday/Sunday
            Select Case Weekday(dtRand, vbSaturday)
                Case 1
                
                    ' randomly select first shifr
                    arrShifts(1, 1) = strLine & "SA" & Application.RandBetween(1, 3)
                    
                    ' check shift still available
                    If dicShifts.Exists(arrShifts(1, 1)) Then
                        arrShifts(1, 2) = dicShifts(arrShifts(1, 1))
                    End If
                    
                    ' randonly select second shift
                    arrShifts(2, 1) = strLine & "SA" & Application.RandBetween(Mid(arrShifts(1, 1), 4), 3)
                    
                    ' check shift still available
                    If dicShifts.Exists(arrShifts(2, 1)) Then
                        arrShifts(2, 2) = dicShifts(arrShifts(2, 1))
                    End If
                    
                Case 2 ' Sunday
                
                   ' randomly select first shifr
                    arrShifts(1, 1) = strLine & "SU" & Application.RandBetween(1, 3)
                    
                    ' check shift still available
                    If dicShifts.Exists(arrShifts(1, 1)) Then
                        arrShifts(1, 2) = dicShifts(arrShifts(1, 1))
                    End If
                    
                    ' randomly select first shifr
                    arrShifts(2, 1) = strLine & "SU" & Application.RandBetween(Mid(arrShifts(1, 1), 4), 3)
                    
                    ' check shift still available
                    If dicShifts.Exists(arrShifts(2, 1)) Then
                        arrShifts(2, 2) = dicShifts(arrShifts(2, 1))
                    End If
                    
                Case Else ' weekday
                
                    boolWeekDay = True
                    
                    ' randomly select first shifr
                    arrShifts(1, 1) = strLine & "W" & Application.RandBetween(1, 18)
                    
                    ' check shift still available
                    If dicShifts.Exists(arrShifts(1, 1)) Then
                        arrShifts(1, 2) = dicShifts(arrShifts(1, 1))
                        
                        ' based on first shift get random consecutive shift
                        arrShifts(2, 1) = GetConsecutiveShift(dicShifts, dicShifts(arrShifts(1, 1)))
                        
                        ' check shift avaialable
                        If arrShifts(2, 1) <> "" Then
                            arrShifts(2, 2) = dicShifts(arrShifts(2, 1))
                        End If
                    End If
                  
            End Select
                            
            ' check if selected shifts exist
            For idx = LBound(arrShifts, 1) To UBound(arrShifts, 1)
                If Not dicShifts.Exists(arrShifts(idx, 1)) Then
                    arrShifts(idx, 1) = ""
                End If
            Next idx

            DoEvents
            
            ' if it's a weekday check the 2 selected shift still available
            ' if it's a weekend check at least one of the selected shifts is available
            
            If (boolWeekDay And dicShifts.Exists(arrShifts(1, 1)) And dicShifts.Exists(arrShifts(2, 1))) _
        Or Not boolWeekDay And (dicShifts.Exists(arrShifts(1, 1)) Or dicShifts.Exists(arrShifts(2, 1))) Then
            
                ' if it's a weekday check 2 shifts have been selected and they don't cover the same time slot
                boolAddShifts = (Not boolWeekDay Or (boolWeekDay And arrShifts(1, 1) <> "" And arrShifts(2, 1) <> "")) And _
                                                                                                                       dicShifts(arrShifts(1, 1)) <> dicShifts(arrShifts(2, 1))
            End If
            
            ' if no of available shifts is less than 5 allow single shift
            If boolAddShifts Or (dicShifts.Count < 5 And arrShifts(1, 1) <> "" Or arrShifts(2, 1) <> "") Then
                                                                                                            
                ' check that adding shifts will not push no of shifts for week over 8
                If dicWeeks(lngWeek) <= 6 Then
                
                    ' check for duplicate shifts
                    If arrShifts(1, 1) = arrShifts(2, 1) Then arrShifts(2, 1) = ""
                    
                    ' add shifts
                    dicPlan(dtRand) = arrShifts
                
                    ' remove shifts that have been added from available shifts list
                    For idx = LBound(arrShifts) To UBound(arrShifts)
                        If dicShifts.Exists(arrShifts(idx, 1)) Then
                            DoEvents
                            dicShifts.Remove arrShifts(idx, 1)
                            dicWeeks(lngWeek) = dicWeeks(lngWeek) + 1
                        End If
                    Next idx
                
                End If
            End If
            
        End If
        
        DoEvents
        
        counter = counter + 1
        
        If counter = 500 Then
            dicPlan.RemoveAll
            Set GetPlan = dicPlan
            Exit Function
        End If
        
    Loop Until dicShifts.Count = 0
    
    Set GetPlan = dicPlan
    
End Function

Function CreateRandomSet(ByVal arrValues As Variant, lngNoValues As Long, Optional boolIndex As Boolean) As Variant
Dim dicSet As Object
Dim arrSetValues As Variant
Dim RandNo As Variant
Dim ky As Variant
Dim idx As Long

    ' create a random set of values from array
    ' no of values in set is set by lngNoValues
    ' if boolIndex is True then use indices of array for the randomization, useful when array has repeating values
    
    Set dicSet = CreateObject("Scripting.Dictionary")
    
    Do
        Randomize Time
        
        RandNo = Application.RandBetween(LBound(arrValues), UBound(arrValues))
        
        If Not boolIndex Then
            If Not dicSet.Exists(arrValues(RandNo)) Then
                dicSet.Add arrValues(RandNo), arrValues(RandNo)
            End If
        Else
            If Not dicSet.Exists(RandNo) Then
                dicSet.Add RandNo, RandNo
            End If
        End If
        
        DoEvents
        
    Loop Until dicSet.Count = lngNoValues

    If boolIndex Then
    
        ReDim arrSetValues(LBound(arrValues) To UBound(arrValues))
        
        idx = LBound(arrValues)
        
        For Each ky In dicSet.Keys
            arrSetValues(idx) = arrValues(ky)
            idx = idx + 1
        Next ky
        
        For idx = LBound(arrValues) To UBound(arrValues)
        
        Next idx
        
    Else
        arrSetValues = dicSet.Keys
    End If
    
    CreateRandomSet = arrSetValues
    
End Function

Function SortDicByDate(ByVal dic As Object, ByVal dtStart, ByVal dtEnd) As Object
Dim dicSorted As Object
Dim idxDate As Date

    ' sort a dictionary chronologically by key
    
    Set dicSorted = CreateObject("Scripting.Dictionary")
    
    For idxDate = dtStart To dtEnd
        If dic.Exists(idxDate) Then
            dicSorted(idxDate) = dic(idxDate)
        End If
    Next idxDate
    
    Set SortDicByDate = dicSorted
    
End Function

Function CreateTimeline(lngStartHr As Long, lngEndHr As Long, lngPeriod As Long) As Variant
Dim Res As Variant
Dim NoHrs As Long
Dim NoPeriods As Long

    ' create a timeline for specified start/end hour and period
    
    NoHrs = lngEndHr - lngStartHr
    
    NoPeriods = NoHrs * 60 / lngPeriod
    
    Res = Evaluate("INDEX(TIME(" & lngStartHr & ", " & lngPeriod & "*(ROW(A1:A" & NoPeriods & ")-1),0),,1)")

    CreateTimeline = Res
    
End Function

Function SheetExists(strSheetName, Optional wb As Workbook) As Boolean
' check to see if sheet named strSheetname exists in active workbook

    SheetExists = Evaluate("ISREF('" & strSheetName & "'!A1)")
    
End Function

