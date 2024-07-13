VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Timers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event TimerEvent(ByVal Index As Integer)
Private mvarCount As Integer
Private mvarCollection As New Collection
Private mvarKeyClass As Integer
Private mvarTimer As Timer
Private Sub Class_Initialize()
   mvarKeyClass = RegisterTimers(Me)
   Set mvarTimer = New Timer
End Sub
Private Sub Class_Terminate()
   Set mvarTimer = Nothing
   Dim tmr As Timer
   For Each tmr In mvarCollection
      Set tmr = Nothing
   Next
   Set mvarCollection = Nothing
   colTimers.Remove "key:" & mvarKeyClass
End Sub
Public Property Get Timer() As Timer
   Set Timer = mvarTimer
End Property
Public Property Set Timer(vData As Timer)
   Set mvarTimer = vData
End Property
Property Get Count() As Integer
   Count = mvarCount
End Property
Function Add() As Integer
   Dim tmr As New Timer
   mvarCount = mvarCount + 1
   tmr.Index = mvarCount
   tmr.ParentKeyClass = mvarKeyClass
   mvarCollection.Add tmr
   Add = mvarCount
End Function
'Public Function BulkLoad(num As Integer)
'   Dim i As Long
'   If num > 0 And num < 200 Then
'      For i = 0 To num
'         Add
'      Next
'   Else
'      Err.Raise vbObjectError, "CTimers.BulkLoad", "No more than 200 times allowed per class"
'   End If
'End Function
Public Function Item(ByVal Index As Integer) As Timer
   If Index > Count Then Err.Raise vbObjectError, "CTimers.Item", "Timer Index larger than count not Exist"
   Set Item = mvarCollection(Index)
End Function
Friend Function RaiseTimer_Event(ByVal Index As Integer)
   RaiseEvent TimerEvent(Index)
End Function
