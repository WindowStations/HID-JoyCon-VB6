VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "HID Joy-Con output"
   ClientHeight    =   13350
   ClientLeft      =   255
   ClientTop       =   330
   ClientWidth     =   20895
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   13350
   ScaleWidth      =   20895
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtButtonEvent 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   12960
      Width           =   8895
   End
   Begin VB.CommandButton cmdDetach 
      Caption         =   "Detatch from application"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   21
      ToolTipText     =   "Detatch and close"
      Top             =   10920
      Width           =   3855
   End
   Begin VB.TextBox txtSend 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   19
      Text            =   "00 01 00 00 00 00 00 00 00 00 00 02"
      Top             =   12240
      Width           =   8895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "00 01 00 00 00 00 00 00 00 00 00 Command"
      Top             =   11880
      Width           =   8895
   End
   Begin VB.CommandButton cmdHCI 
      Caption         =   "HCI reboot"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   18
      ToolTipText     =   "Causes the controller to change power state."
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton cmdDeviceInfo 
      Caption         =   "Device info"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   17
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton cmdDisable6axis 
      Caption         =   "Disable 6-Axis"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   16
      Top             =   10080
      Width           =   1815
   End
   Begin VB.CommandButton cmdEnable6axis 
      Caption         =   "Enable 6-Axis"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   15
      Top             =   10080
      Width           =   1815
   End
   Begin VB.CommandButton cmdLightOff 
      Caption         =   "Lights off"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   14
      Top             =   9240
      Width           =   1815
   End
   Begin VB.CommandButton cmdLightsOn 
      Caption         =   "Lights on"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   13
      Top             =   9240
      Width           =   1815
   End
   Begin VB.OptionButton optNFC 
      Caption         =   "NFC/IR"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      Top             =   7560
      Width           =   1215
   End
   Begin VB.OptionButton optFull 
      Caption         =   "Full"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   7560
      Width           =   1215
   End
   Begin VB.OptionButton optBasic 
      Caption         =   "Basic "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   9
      Top             =   7560
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdDevices 
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Top             =   4800
      Width           =   1815
   End
   Begin VB.ListBox lstConnection 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      ItemData        =   "frmMain.frx":15467
      Left            =   240
      List            =   "frmMain.frx":15469
      TabIndex        =   0
      Top             =   1080
      Width           =   5565
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Write report"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16680
      TabIndex        =   6
      Top             =   12120
      Width           =   1815
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear events"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18000
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdPoll 
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18600
      TabIndex        =   2
      Top             =   12120
      Width           =   1815
   End
   Begin VB.TextBox txtState 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10455
      Left            =   6000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1080
      Width           =   14565
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Input report mode:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   12
      Top             =   7080
      Width           =   1950
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Bytes to send:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   8
      Top             =   12240
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Joy-Con events:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6120
      TabIndex        =   5
      Top             =   600
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Connection status:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   1905
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents hid As clsHIDJoyCon
Attribute hid.VB_VarHelpID = -1
Private Declare Function apiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long) As Long
Private isloading As Boolean
Private Sub Form_Load()
   Set hid = New clsHIDJoyCon
   hid.FindHID
   hid.WriteReadDevices VBA.Trim("00 01 00 00 00 00 00 00 00 00 00 03 3F")
End Sub
Private Sub Form_Resize()
   If Me.WindowState = vbNormal Then
      Me.Height = 13980
      Dim w As Long
      w = (Me.Width - txtState.Left) - 300
      If w > 32 Then txtState.Width = w
   End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
   hid.FormUnloading
End Sub
Private Sub cmdDevices_Click()
   If cmdDevices.Caption = "Poll devices" Then  'Enables the user to select 1-time or continuous data transfers.
      cmdDevices.Caption = "Pause"
      hid.EnableFindDevices 'Enable the timer that finds devices
      ChangeList "Device polling started"
      hid.FindHID
   Else
      cmdDevices.Caption = "Poll devices" 'Change the command button to Continuous
      hid.DisableFindDevices  'Disable the timer that finds devices
      ChangeList "Device polling paused"
   End If
End Sub
Private Sub cmdPoll_Click()
   If cmdPoll.Caption = "Poll reports" Then  'Enables the user to select 1-time or continuous data transfers.
      cmdPoll.Caption = "Pause" 'Change the command button to pause
      hid.EnablePoll 'Enable the timer to read and write to the device once/second.
      ChangeList "Event polling started"
   Else
      cmdPoll.Caption = "Poll reports" 'Change the command button to Continuous
      hid.DisablePoll 'Disable the timer that reads and writes to the device once/second.
      ChangeList "Event polling paused"
   End If
End Sub
Private Sub cmdClear_Click()
   txtState.Text = ""
   lstConnection.Clear
End Sub
Private Sub cmdWrite_Click()
   hid.WriteReadDevices VBA.Trim(txtSend.Text)
End Sub
'SendBuffer commands in bytes with bitfields
'00 01 00 00 00 00 00 00 00 00 00 02"Request device info while in 3F mode.  Responds with 82 02 for success."
'00 01 00 00 00 00 00 00 00 00 00 30 00 00 00 00 03 02 01 00 "Set lights off."
'00 01 00 00 00 00 00 00 00 00 00 30 03 02 01 00 00 00 00 00 "Set lights on."
'00 01 00 00 00 00 00 00 00 00 00 06 04  "Set HCI state (disconnect/page/pair/off). Reboots and Reconnects (page mode/HOME mode resets to 3F w/no flashing lights)."
'00 01 00 00 00 00 00 00 00 00 00 03 3F "Set input report mode to Basic HID mode.  Pushes updates with every button press."
'00 01 00 00 00 00 00 00 00 00 00 03 30 "Set input report mode to Standard full mode.  Pushes current state @60Hz."
'00 01 00 00 00 00 00 00 00 00 00 03 31 "Set input report mode to NFC/IR mode. Pushes large packets @60Hz."
'00 01 00 00 00 00 00 00 00 00 00 40 01 "Enable IMU (6-Axis sensor) while in 30 mode."
'00 01 00 00 00 00 00 00 00 00 00 40 00 "Disable IMU (6-Axis sensor) while in 30 mode."
Private Sub optBasic_Click()
   If isloading = False Then
      hid.WriteReadDevices VBA.Trim("00 01 00 00 00 00 00 00 00 00 00 03 3F")
   End If
End Sub
Private Sub optFull_Click()
   If isloading = False Then
      hid.WriteReadDevices VBA.Trim("00 01 00 00 00 00 00 00 00 00 00 03 30")
   End If
End Sub
Private Sub optNFC_Click()
   If isloading = False Then
      hid.WriteReadDevices VBA.Trim("00 01 00 00 00 00 00 00 00 00 00 03 31")
   End If
End Sub
Private Sub cmdDeviceInfo_Click()
   If optBasic.Value = False Then
      optBasic.Value = True
      isloading = True
   End If
   hid.WriteReadDevices VBA.Trim("00 01 00 00 00 00 00 00 00 00 00 03 3F")
   hid.FlushQueue
   DoEvents
   apiSleep 25
   hid.WriteReadDevices VBA.Trim("00 01 00 00 00 00 00 00 00 00 00 02")
   isloading = False
End Sub
Private Sub cmdHCI_Click()
   Dim mbr As VbMsgBoxResult
   mbr = MsgBox("Would you like to turn off the JoyCons and manually reboot them?", vbYesNo, "Causes the controller to change power state and sleep.")
   If mbr = vbYes Then
      hid.WriteReadDevices VBA.Trim("00 01 00 00 00 00 00 00 00 00 00 06 04")
   End If
End Sub
Private Sub cmdLightOff_Click()
   hid.WriteReadDevices VBA.Trim("00 01 00 00 00 00 00 00 00 00 00 30 00")
End Sub
Private Sub cmdLightsOn_Click()
   hid.WriteReadDevices VBA.Trim("00 01 00 00 00 00 00 00 00 00 00 30 01")
End Sub
Private Sub cmdEnable6axis_Click()
   If optFull.Value = False Then
      optFull.Value = True
      isloading = True
   End If
   hid.WriteReadDevices VBA.Trim("00 01 00 00 00 00 00 00 00 00 00 03 30")
   hid.FlushQueue
   DoEvents
   apiSleep 25
   hid.WriteReadDevices VBA.Trim("00 01 00 00 00 00 00 00 00 00 00 40 01")
   isloading = False
End Sub
Private Sub cmdDisable6axis_Click()
   hid.WriteReadDevices VBA.Trim("00 01 00 00 00 00 00 00 00 00 00 40 00")
End Sub
Private Sub cmdDetach_Click()
   'Detatches JoyCon from calling application.  The application will need to be restarted in order to access JoyCons after detatching them.
   hid.WriteReadDevices VBA.Trim("00 01 00 00 00 00 00 00 00 00 00 30 01")
   hid.WriteReadDevices VBA.Trim("00 01 00 00 00 00 00 00 00 00 00 40 00")
   hid.WriteReadDevices VBA.Trim("00 01 00 00 00 00 00 00 00 00 00 48 00")
   hid.WriteReadDevices VBA.Trim("00 01 00 00 00 00 00 00 00 00 00 03 3F")
   hid.WriteReadDevices VBA.Trim("00 01 00 00 00 00 00 00 00 00 00 06 01")
   Shell (App.Path & "\" & App.EXEName)
   Unload Me
End Sub
Private Sub hid_DeviceConnection(ByVal Index As Integer, ByVal connected As Boolean)
   If Index = 1 Then
      If connected = True Then
         ChangeList "Joy-Con (Left) detected"
      Else
         ChangeList "Joy-Con (Left) not detected"
      End If
   ElseIf Index = 2 Then
      If connected = True Then
         ChangeList "Joy-Con (Right) detected"
      Else
         ChangeList "Joy-Con (Right) not detected"
      End If
   End If
End Sub
Private Sub hid_ClickButton(ByVal joycon As Integer, ByVal down As Boolean, ByVal btns As String)
   'txtButtonEvent.Text = "JoyCon " & CStr(joycon) & " " & btns & vbCrLf
End Sub
Private Sub hid_WriteReportResult(ByVal Report As String)
    'txtState.Text = txtState.Text & Report & vbCrLf
End Sub
Public Sub ChangeText(ByVal txt As String)
   Dim s() As String
   s = VBA.Split(txtState.Text, vbCrLf)
   If UBound(s) > 50 Then 'limit the number of lines in the textbox
      Dim x As String
      Dim z As Long
      For z = 1 To 50
         x = x & s(z) & vbCrLf
      Next
      txtState.Text = x
   End If
   txtState.Text = txtState.Text & txt & vbCrLf
   txtState.SelStart = Len(txtState.Text) - 1
   txtState.SelLength = 1 'scroll
End Sub
Public Sub ChangeList(ByVal txt As String)
   lstConnection.AddItem txt & " " & Now
   If lstConnection.ListCount > 30 Then
      Dim z As Long 'trim list count
      For z = 1 To (lstConnection.ListCount - 30)
         lstConnection.RemoveItem z - 1
      Next
   End If
   lstConnection.ListIndex = lstConnection.ListCount - 1 'scroll
End Sub
