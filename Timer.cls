VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const DEBUGMODE As Boolean = False
Private Declare Function apiSetTimer Lib "user32" Alias "SetTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function apiKillTimer Lib "user32" Alias "KillTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private timerId As Long
Private isActive As Boolean
Public Interval As Long
Public Tag As Variant
Public Event Timer()
'Public Event Elapsed()
Private mvarIndex As Long
Private mvarParentKeyClass As Variant
Private Sub Class_Initialize()
   Interval = 2000
End Sub
Private Sub Class_Terminate()
   TimerDestroy
End Sub
Friend Property Let Index(dwIndex As Long)
mvarIndex = dwIndex
End Property
Friend Property Get Index() As Long
Index = mvarIndex
End Property
Friend Property Let ParentKeyClass(key As Long)
mvarParentKeyClass = key
End Property
Friend Property Get ParentKeyClass() As Long
ParentKeyClass = mvarParentKeyClass
End Property
Property Get Enabled() As Boolean
   Enabled = isActive
End Property
Property Let Enabled(ByRef activate As Boolean)
   If activate = True And isActive = True Then
      Reset
   ElseIf activate = True And Not isActive = True Then
      Enable
   ElseIf isActive = True And Not activate = True Then
      Disable
   ElseIf Not isActive = True And Not activate = True Then
   End If
End Property
Sub EnableTimer(mInterval As Long)
   If Interval = mInterval And isActive = True Then TimerDestroy
   Interval = mInterval
   Enable
End Sub
Sub Enable()
   If isActive = True Then TimerDestroy
   TimerCreate
End Sub
Sub Disable()
   TimerDestroy
End Sub
Sub Reset()
   TimerDestroy
   TimerCreate
End Sub
Public Sub RaiseTimer_Event()
   RaiseEvent Timer
   ' RaiseEvent Elapsed
End Sub
Private Function TimerCreate() As Boolean
   If Interval <= 0 Then Exit Function
   timerId = apiSetTimer(0, 0, Interval, AddressOf TimerProc)
   If timerId <> 0 Then
      TimerCreate = True
      Timercollection.Add Me, "id:" & timerId
      isActive = True
   Else
      timerId = 0
      TimerCreate = False
      isActive = False
   End If
End Function
Private Function TimerDestroy() As Long
   Dim i As Integer
   Dim f As Boolean
   If TimerExists() = True Then
      f = apiKillTimer(0, timerId)
      Timercollection.Remove "id:" & timerId
      TimerDestroy = True
      isActive = False
   End If
End Function
Private Function TimerExists() As Boolean
   On Error Resume Next: Err.Clear
   Dim c As Timer
   Set c = Timercollection("id:" & timerId)
   If Err.Number = 0 Then TimerExists = True
   Set c = Nothing
End Function
