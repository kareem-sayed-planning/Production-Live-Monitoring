' ==============================================================================
' CORE SYSTEM: Asynchronous Schedulers & Andon Engine
' Purpose: Runs background tasks seamlessly without interrupting the UI dashboard.
' ==============================================================================
Public RunBlinker As Boolean
Public GlobalMaxDT As Long
Public NextDataUpdate As Date

' --- 1. HEAVY DATA IMPORT SCHEDULER ---
Sub Update_Data_Only_Scheduled()
    On Error GoTo EmergencyReset

    ' Prevent overlap if previous heavy query is still running
    If IsDataMacroRunning Then Exit Sub
    IsDataMacroRunning = True
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Call the core dictionary/array data pipeline
    Call Data_Import_Heavy_Process
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    IsDataMacroRunning = False

Reschedule:
    ' Schedule next run strictly in 10 minutes
    NextDataUpdate = Now + TimeValue("00:10:00")
    Application.OnTime EarliestTime:=NextDataUpdate, Procedure:="Update_Data_Only_Scheduled", Schedule:=True
    Exit Sub

EmergencyReset:
    Err.Clear
    IsDataMacroRunning = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Resume Reschedule
End Sub

' --- 2. ANDON MONITORING ENGINE (2-Min Cycle) ---
Sub Update_Andon_Numbers()
    If Not RunBlinker Then Exit Sub
    
    Dim wsFinal As Worksheet
    Dim currentDT As Long, currentVar As Double, maxDT As Long
    Dim i As Integer, varCol As String, shp As Shape
    
    Set wsFinal = ThisWorkbook.Sheets("Final")
    maxDT = 0
    varCol = IIf(wsFinal.Range("AG1").Value = "1st", "AE", "AF") ' Dynamic Shift Column
    
    ' Identify max downtime ONLY for shops missing their volume targets
    For i = 2 To 5
        currentDT = wsFinal.Cells(i, "Z").Value
        currentVar = wsFinal.Cells(i, varCol).Value
        If currentVar < 0 And currentDT > maxDT Then maxDT = currentDT
    Next i
    
    GlobalMaxDT = maxDT
    
    ' Update Visual UI (Andon Light Status)
    Set shp = wsFinal.Shapes("Circle_Status")
    If Not shp Is Nothing Then
        If GlobalMaxDT >= 30 Then
            shp.Fill.ForeColor.RGB = RGB(255, 0, 0)       ' Critical: Red
        ElseIf GlobalMaxDT >= 20 Then
            shp.Fill.ForeColor.RGB = RGB(255, 192, 0)     ' Warning: Yellow
        Else
            shp.Fill.ForeColor.RGB = RGB(0, 255, 0)       ' Normal: Green
        End If
    End If
    
    ' Schedule next check
    Application.OnTime EarliestTime:=Now + TimeValue("00:02:00"), Procedure:="Update_Andon_Numbers", Schedule:=True
End Sub
