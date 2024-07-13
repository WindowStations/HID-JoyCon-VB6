Attribute VB_Name = "modTimer"
Option Explicit
Private Declare Function apiKillTimer Lib "user32" Alias "KillTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Public Timercollection As New Collection
Public colTimers As New Collection
Private tCount As Integer
Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
   On Error Resume Next: Err.Clear
   Dim tmr As Timer
   Dim tmrs As Timers
   Set tmr = Timercollection("id:" & idEvent)
   If tmr Is Nothing Then
      apiKillTimer 0, idEvent
   Else
      If tmr.ParentKeyClass > 0 Then
         Set tmrs = colTimers("key:" & tmr.ParentKeyClass)
         If tmrs Is Nothing Then
            apiKillTimer 0, idEvent
         Else
            tmrs.RaiseTimer_Event tmr.Index
         End If
      Else
         tmr.RaiseTimer_Event
      End If
   End If
   Set tmr = Nothing
End Sub
Public Function RegisterTimers(ByRef tmrs As Timers) As Integer
   On Error Resume Next: Err.Clear
   Dim key As String
   tCount = tCount + 1
   key = "key:" & tCount
   colTimers.Add tmrs, key
   RegisterTimers = tCount
End Function
