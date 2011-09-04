Attribute VB_Name = "LibPerf"
Option Explicit

Private ScreenUpdateState As Boolean
Private StatusBarState As Boolean
Private CalcState As XlCalculation
Private EventsState As Boolean
Private PageBreaksState As Boolean

Public Sub OptimizePerformance()
    With Application
        ScreenUpdatingState = .ScreenUpdating
        StatusBarState = .DisplayStatusBar
        CalcState = .Calculation
        EventsState = .EnableEvents
        PageBreaksState = ActiveSheet.DisplayPageBreaks
        
        .ScreenUpdating = False
        .DisplayStatusBar = False
        .Calculation = xlCalculationManual
        .EnableEvents = EventsState
        ActiveSheet.DisplayPageBreaks = PageBreaksState
    End With
End Sub

Public Sub RestoreFunctionality()
    With Application
        .ScreenUpdating = ScreenUpdateState
        .DisplayStatusBar = StatusBarState
        .Calculation = CalcState
        .EnableEvents = EventsState
        ActiveSheet.DisplayPageBreaks = PageBreaksState
    End With
End Sub
