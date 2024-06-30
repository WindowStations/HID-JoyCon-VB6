VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HID Joy-Con output"
   ClientHeight    =   12690
   ClientLeft      =   180
   ClientTop       =   255
   ClientWidth     =   12405
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12690
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClear 
      Caption         =   "ClearJoy-Con events"
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
      Left            =   7920
      TabIndex        =   5
      Top             =   11880
      Width           =   2415
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Joy-Con (Right)"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Joy-Con (Left)"
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
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.CommandButton cmdPoll 
      Caption         =   "Poll reports"
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
      Left            =   2160
      TabIndex        =   2
      Top             =   11880
      Width           =   1452
   End
   Begin VB.TextBox Text1 
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
      Height          =   10215
      Left            =   6240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1440
      Width           =   6000
   End
   Begin VB.Timer tmrContinuousDataCollect 
      Enabled         =   0   'False
      Left            =   11880
      Top             =   120
   End
   Begin VB.ListBox lstResults 
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
      Height          =   10140
      ItemData        =   "frmMain.frx":15467
      Left            =   120
      List            =   "frmMain.frx":15469
      TabIndex        =   0
      Top             =   1440
      Width           =   6000
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   960
      Width           =   6015
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   6015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Set these to match the values in the device's firmware and INF file.
'
'Vendor 'Nintendo' &H57E
'Product 'Joy-Con (L)' &H2006
'Product 'Joy-Con (R)' &H2007
'
'Vendor 'Scuf Gaming' &H2E95
'Product 'Instinct Pro' &H504 - 'Xbox Controller' &H7725
'
Private Const VendorId As Long = &H57E
Private Const productIdJoyConL As Long = &H2006
Private Const productIdJoyConR As Long = &H2007
Private Const DIGCF_PRESENT As Long = &H2 'setupapi.h
Private Const DIGCF_DEVICEINTERFACE As Long = &H10
Private Const FILE_FLAG_OVERLAPPED As Long = &H40000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const OPEN_EXISTING As Long = 3
Private Const WAIT_TIMEOUT As Long = &H102&
Private Const WAIT_OBJECT_0 As Long = 0
Private Const HidP_Input As Integer = 0
Private Const HidP_Output As Integer = 1
Private Const HidP_Feature As Integer = 2
'
'LEFT JOYCON bit constants
Private Const VK_PAD_LTHUMB_RIGHT As Long = &H0
Private Const VK_PAD_LTHUMB_DOWNRIGHT As Long = &H1
Private Const VK_PAD_LTHUMB_DOWN As Long = &H2
Private Const VK_PAD_LTHUMB_DOWNLEFT As Long = &H3
Private Const VK_PAD_LTHUMB_LEFT As Long = &H4
Private Const VK_PAD_LTHUMB_UPLEFT As Long = &H5
Private Const VK_PAD_LTHUMB_UP As Long = &H6
Private Const VK_PAD_LTHUMB_UPRIGHT As Long = &H7
Private Const VK_PAD_LTHUMB_DEAD As Long = &H8
Private Const VK_PAD_MINUS As Long = &H1
Private Const VK_PAD_LTHUMB_PRESS As Long = &H4
Private Const VK_PAD_BACK As Long = &H2
Private Const VK_PAD_LSHOULDER As Long = &H4
Private Const VK_PAD_LTRIGGER As Long = &H8
Private Const VK_PAD_DPAD_LEFT As Long = &H1
Private Const VK_PAD_DPAD_DOWN As Long = &H2
Private Const VK_PAD_DPAD_UP As Long = &H4
Private Const VK_PAD_DPAD_RIGHT As Long = &H8
Private Const VK_PAD_SL1 As Long = &H1
Private Const VK_PAD_SR1 As Long = &H2
'
'RIGHT JOYCON bit Private Const ants
Private Const VK_PAD_RTHUMB_LEFT As Long = &H0
Private Const VK_PAD_RTHUMB_UPLEFT As Long = &H1
Private Const VK_PAD_RTHUMB_UP As Long = &H2
Private Const VK_PAD_RTHUMB_UPRIGHT As Long = &H3
Private Const VK_PAD_RTHUMB_RIGHT As Long = &H4
Private Const VK_PAD_RTHUMB_DOWNRIGHT As Long = &H5
Private Const VK_PAD_RTHUMB_DOWN As Long = &H6
Private Const VK_PAD_RTHUMB_DOWNLEFT As Long = &H7
Private Const VK_PAD_RTHUMB_DEAD As Long = &H8
Private Const VK_PAD_PLUS As Long = &H2
Private Const VK_PAD_RTHUMB_PRESS As Long = &H8
Private Const VK_PAD_START As Long = &H1
Private Const VK_PAD_RSHOULDER As Long = &H4
Private Const VK_PAD_RTRIGGER As Long = &H8
Private Const VK_PAD_B As Long = &H1
Private Const VK_PAD_Y As Long = &H2
Private Const VK_PAD_A As Long = &H4
Private Const VK_PAD_X As Long = &H8
Private Const VK_PAD_SL2 As Long = &H1
Private Const VK_PAD_SR2 As Long = &H2
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Private Type HIDD_ATTRIBUTES
   Size As Long
   VendorId As Integer
   ProductID As Integer
   VersionNumber As Integer
End Type
Private Type HIDP_CAPS 'hidpi.h
   Usage As Integer
   UsagePage As Integer
   InputReportByteLength As Integer
   OutputReportByteLength As Integer
   FeatureReportByteLength As Integer
   Reserved(16) As Integer
   NumberLinkCollectionNodes As Integer
   NumberInputButtonCaps As Integer
   NumberInputValueCaps As Integer
   NumberInputDataIndices As Integer
   NumberOutputButtonCaps As Integer
   NumberOutputValueCaps As Integer
   NumberOutputDataIndices As Integer
   NumberFeatureButtonCaps As Integer
   NumberFeatureValueCaps As Integer
   NumberFeatureDataIndices As Integer
End Type
Private Type HidP_Value_Caps
   UsagePage As Integer
   ReportID As Byte
   IsAlias As Long
   BitField As Integer
   LinkCollection As Integer
   LinkUsage As Integer
   LinkUsagePage As Integer
   IsRange As Long 'If IsRange is false, UsageMin is the Usage and UsageMax is unused.
   IsStringRange As Long 'If IsStringRange is false, StringMin is the string index and StringMax is unused.
   IsDesignatorRange As Long 'If IsDesignatorRange is false, DesignatorMin is the designator index and DesignatorMax is unused.
   IsAbsolute As Long
   HasNull As Long
   Reserved As Byte
   BitSize As Integer
   ReportCount As Integer
   Reserved2 As Integer
   Reserved3 As Integer
   Reserved4 As Integer
   Reserved5 As Integer
   Reserved6 As Integer
   LogicalMin As Long
   LogicalMax As Long
   PhysicalMin As Long
   PhysicalMax As Long
   UsageMin As Integer
   UsageMax As Integer
   StringMin As Integer
   StringMax As Integer
   DesignatorMin As Integer
   DesignatorMax As Integer
   DataIndexMin As Integer
   DataIndexMax As Integer
End Type
Private Type OVERLAPPED
   Internal As Long
   InternalHigh As Long
   Offset As Long
   OffsetHigh As Long
   hEvent As Long
End Type
Private Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Long
End Type
Private Type SP_DEVICE_INTERFACE_DATA
   cbSize As Long
   InterfaceClassGuid As GUID
   Flags As Long
   Reserved As Long
End Type
Private Type SP_DEVICE_INTERFACE_DETAIL_DATA
   cbSize As Long
   DevicePath As Byte
End Type
Private Type SP_DEVINFO_DATA
   cbSize As Long
   ClassGuid As GUID
   DevInst As Long
   Reserved As Long
End Type
'Private Declare Function apilstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal dest As String, ByVal source As Long) As String
'Private Declare Function apilstrlen Lib "kernel32" Alias "lstrlenA" (ByVal source As Long) As Long
'Private Declare Function apiSetupDiCreateDeviceInfoList Lib "setupapi.dll" Alias "SetupDiCreateDeviceInfoList" (ByRef ClassGuid As GUID, ByVal hwndParent As Long) As Long
Private Declare Function apiCancelIo Lib "kernel32" Alias "CancelIo" (ByVal hFile As Long) As Long
Private Declare Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private Declare Function apiCreateEvent Lib "kernel32" Alias "CreateEventA" (ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Private Declare Function apiCreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function apiFormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageZId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByVal Arguments As Long) As Long
Private Declare Function apiReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As OVERLAPPED) As Long
Private Declare Function apiResetEvent Lib "kernel32" Alias "ResetEvent" (ByVal hEvent As Long) As Long
Private Declare Function apiRtlMoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal Count As Long) As Long
Private Declare Function apiWaitForSingleObject Lib "kernel32" Alias "WaitForSingleObject" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function apiWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function apiHidD_FreePreparsedData Lib "hid.dll" Alias "HidD_FreePreparsedData" (ByRef PreparsedData As Long) As Long
Private Declare Function apiHidD_GetAttributes Lib "hid.dll" Alias "HidD_GetAttributes" (ByVal HidDeviceObject As Long, ByRef Attributes As HIDD_ATTRIBUTES) As Long
Private Declare Function apiHidD_GetHidGuid Lib "hid.dll" Alias "HidD_GetHidGuid" (ByRef HidGuid As GUID) As Long
Private Declare Function apiHidD_GetPreparsedData Lib "hid.dll" Alias "HidD_GetPreparsedData" (ByVal HidDeviceObject As Long, ByRef PreparsedData As Long) As Long
Private Declare Function apiHidP_GetCaps Lib "hid.dll" Alias "HidP_GetCaps" (ByVal PreparsedData As Long, ByRef Capabilities As HIDP_CAPS) As Long
Private Declare Function apiHidP_GetValueCaps Lib "hid.dll" Alias "HidP_GetValueCaps" (ByVal ReportType As Integer, ByRef ValueCaps As Byte, ByRef ValueCapsLength As Integer, ByVal PreparsedData As Long) As Long
Private Declare Function apiSetupDiDestroyDeviceInfoList Lib "setupapi.dll" Alias "SetupDiDestroyDeviceInfoList" (ByVal DeviceInfoSet As Long) As Long
Private Declare Function apiSetupDiEnumDeviceInterfaces Lib "setupapi.dll" Alias "SetupDiEnumDeviceInterfaces" (ByVal DeviceInfoSet As Long, ByVal DeviceInfoData As Long, ByRef InterfaceClassGuid As GUID, ByVal MemberIndex As Long, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA) As Long
Private Declare Function apiSetupDiGetClassDevs Lib "setupapi.dll" Alias "SetupDiGetClassDevsA" (ByRef ClassGuid As GUID, ByVal Enumerator As String, ByVal hwndParent As Long, ByVal Flags As Long) As Long
Private Declare Function apiSetupDiGetDeviceInterfaceDetail Lib "setupapi.dll" Alias "SetupDiGetDeviceInterfaceDetailA" (ByVal DeviceInfoSet As Long, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA, ByVal DeviceInterfaceDetailData As Long, ByVal DeviceInterfaceDetailDataSize As Long, ByRef RequiredSize As Long, ByVal DeviceInfoData As Long) As Long
Private bAlertable As Long
Private Capabilities As HIDP_CAPS
Private DataString As String
Private DetailData As Long
Private DetailDataBuffer() As Byte
Private DeviceAttributes As HIDD_ATTRIBUTES
Private DevicePathName As String
Private DeviceInfoSet As Long
Private ErrorString As String
Private EventObject As Long
Private HIDHandle As Long
Private HIDOverlapped As OVERLAPPED
Private LastDevice As Boolean
Private MyDeviceDetected As Boolean
Private MyDeviceInfoData As SP_DEVINFO_DATA
Private MyDeviceInterfaceDetailData As SP_DEVICE_INTERFACE_DETAIL_DATA
Private MyDeviceInterfaceData As SP_DEVICE_INTERFACE_DATA
Private Needed As Long
Private OutputReportData(0) As Byte
Private PreparsedData As Long
Private ReadHandle As Long
Private Result As Long
Private Security As SECURITY_ATTRIBUTES
Private Timeout As Boolean
Private unloading As Boolean
Private ProductID As Long
Private Sub Form_Load()
   Timeout = True
   ProductID = productIdJoyConL
   tmrContinuousDataCollect.Interval = 1
   ReadAndWriteToDevices
End Sub
Private Sub Form_Unload(Cancel As Integer)
   unloading = True
   Result = apiCloseHandle(HIDHandle) 'Actions that must execute when the program ends.   'Close the open handles to the device.
   End
End Sub
Private Sub cmdPoll_Click()
   If cmdPoll.Caption = "Poll reports" Then  'Enables the user to select 1-time or continuous data transfers.
      cmdPoll.Caption = "Pause" 'Change the command button to pause
      tmrContinuousDataCollect.Enabled = True    'Enable the timer to read and write to the device once/second.
      ChangeList "Polling started"
   Else
      cmdPoll.Caption = "Poll reports" 'Change the command button to Continuous
      tmrContinuousDataCollect.Enabled = False  'Disable the timer that reads and writes to the device once/second.
      ChangeList "Polling paused"
   End If
End Sub
Private Sub Option1_Click()
   MyDeviceDetected = False
   Dim b As Boolean
   b = tmrContinuousDataCollect.Enabled
   If b = True Then tmrContinuousDataCollect.Enabled = False
   ProductID = productIdJoyConL 'set to the left joycon
   If cmdPoll.Caption = "Poll reports" Then 'not polling so poll once
      ReadAndWriteToDevices
   End If
   If b = True Then tmrContinuousDataCollect.Enabled = True
End Sub
Private Sub Option2_Click()
   MyDeviceDetected = False
   Dim b As Boolean
   b = tmrContinuousDataCollect.Enabled
   If b = True Then tmrContinuousDataCollect.Enabled = False
   ProductID = productIdJoyConR 'set to the right joycon
   If cmdPoll.Caption = "Poll reports" Then
      ReadAndWriteToDevices
   End If
   If b = True Then tmrContinuousDataCollect.Enabled = True
End Sub
Private Sub cmdClear_Click()
   Text1.Text = ""
End Sub
Private Sub tmrContinuousDataCollect_Timer()
   ReadAndWriteToDevices
End Sub
Private Sub ReadAndWriteToDevices()
   If MyDeviceDetected = False Then 'If device hasn't been detected or it timed out
      MyDeviceDetected = FindTheHid
   End If
   If MyDeviceDetected = True Then
      ReadReport 'Read a report from device.
   End If
End Sub
Private Sub ReadReport()
   Dim Count As Integer
   Dim NumberOfBytesRead As Long
   Dim ReadBuffer() As Byte
   Dim UBoundReadBuffer As Integer
   Dim ByteValue As String
   If Capabilities.InputReportByteLength > 0 Then 'The ReadBuffer array begins at 0, so subtract 1 from the number of bytes.
      ReDim ReadBuffer(Capabilities.InputReportByteLength - 1)
   Else
      ReDim ReadBuffer(0)
   End If
   Result = apiReadFile(ReadHandle, ReadBuffer(0), CLng(Capabilities.InputReportByteLength), NumberOfBytesRead, HIDOverlapped)   'Do an overlapped ReadFile. 'The function returns immediately, even if the data hasn't been received yet.
   bAlertable = True
   DoEvents
   Result = apiWaitForSingleObject(EventObject, 400)
   DoEvents
   Select Case Result 'Find out if ReadFile completed or timeout.
   Case WAIT_OBJECT_0
      'ReadFile has completed
   Case WAIT_TIMEOUT
      Result = apiCancelIo(ReadHandle) 'Returns non-zero on success.
      apiCloseHandle HIDHandle  'look for the device on the next attempt.
      apiCloseHandle ReadHandle
      MyDeviceDetected = False
   Case Else
      MyDeviceDetected = False
      ChangeList "wait cancelled"
   End Select
   If UBound(ReadBuffer) > 0 And ReadBuffer(0) > 0 Then
      Dim bv As String
      For Count = 1 To UBound(ReadBuffer)
         If Len(Hex$(ReadBuffer(Count))) < 2 Then  'Add a leading 0 to values 0 - Fh.
            ByteValue = "0" & Hex$(ReadBuffer(Count))
         Else
            ByteValue = Hex$(ReadBuffer(Count))
         End If
         bv = bv & ByteValue
      Next
      ChangeText GetButtonState(bv)
      apiResetEvent EventObject
   End If
End Sub
Function FindTheHid() As Boolean
   Dim Count As Integer 'Makes a series of API calls to locate the desired HID-class device.   'Returns True if the device is detected, False if not detected.
   Dim GUIDString As String
   Dim HidGuid As GUID
   Dim MemberIndex As Long
   LastDevice = False
   MyDeviceDetected = False
   Security.lpSecurityDescriptor = 0
   Security.bInheritHandle = True
   Security.nLength = Len(Security)
   Result = apiHidD_GetHidGuid(HidGuid)
   GUIDString = Hex$(HidGuid.Data1) & "-" & Hex$(HidGuid.Data2) & "-" & Hex$(HidGuid.Data3) & "-"
   For Count = 0 To 7
      If HidGuid.Data4(Count) >= &H10 Then 'Ensure that each of the 8 bytes in the GUID displays two characters.
         GUIDString = GUIDString & Hex$(HidGuid.Data4(Count)) & " "
      Else
         GUIDString = GUIDString & "0" & Hex$(HidGuid.Data4(Count)) & " "
      End If
   Next
   DeviceInfoSet = apiSetupDiGetClassDevs(HidGuid, vbNullString, 0, (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))
   DataString = GetDataString(DeviceInfoSet, 32)
   MemberIndex = 0 'Begin with 0 and increment until no more devices are detected.
   If DeviceInfoSet <> 0 Then
      Do
         MyDeviceInterfaceData.cbSize = LenB(MyDeviceInterfaceData)   'The cbSize element of the MyDeviceInterfaceData structure must be set to the structure's size in bytes. The size is 28 bytes.
         Result = apiSetupDiEnumDeviceInterfaces(DeviceInfoSet, 0, HidGuid, MemberIndex, MyDeviceInterfaceData)
         If Result = 0 Then
            LastDevice = True
         End If
         If Result <> 0 Then 'If a device exists, display the information returned.
            'SetupDiGetDeviceInterfaceDetail   Returns: an SP_DEVICE_INTERFACE_DETAIL_DATA structure containing information about a device.
            'To retrieve the information, call this function twice.  The first time returns the size of the structure in Needed. The second time returns a pointer to the data in DeviceInfoSet.
            'Requires: A DeviceInfoSet returned by SetupDiGetClassDevs and an SP_DEVICE_INTERFACE_DATA structure returned by SetupDiEnumDeviceInterfaces.
            MyDeviceInfoData.cbSize = Len(MyDeviceInfoData)
            Result = apiSetupDiGetDeviceInterfaceDetail(DeviceInfoSet, MyDeviceInterfaceData, 0, 0, Needed, 0)
            DetailData = Needed
            MyDeviceInterfaceDetailData.cbSize = Len(MyDeviceInterfaceDetailData)
            ReDim DetailDataBuffer(Needed) 'Use a byte array to allocate memory for the MyDeviceInterfaceDetailData structure
            apiRtlMoveMemory DetailDataBuffer(0), MyDeviceInterfaceDetailData, 4 'Store cbSize in the first four bytes of the array.
            Result = apiSetupDiGetDeviceInterfaceDetail(DeviceInfoSet, MyDeviceInterfaceData, VarPtr(DetailDataBuffer(0)), DetailData, Needed, 0)
            DevicePathName = CStr(DetailDataBuffer())  'Convert the byte array to a string.
            DevicePathName = StrConv(DevicePathName, vbUnicode)  'Convert to Unicode.
            DevicePathName = Right$(DevicePathName, Len(DevicePathName) - 4) 'Strip cbSize (4 bytes) from the beginning.
            'Requires: The DevicePathName returned by SetupDiGetDeviceInterfaceDetail.
            HIDHandle = apiCreateFile(DevicePathName, GENERIC_READ Or GENERIC_WRITE, (FILE_SHARE_READ Or FILE_SHARE_WRITE), Security, OPEN_EXISTING, 0&, 0)
            'Now we can find out if it's the device we're looking for.
            'HidD_GetAttributes Requests information from the device.  'Requires: The handle returned by apiCreateFile. 'Returns: an HIDD_ATTRIBUTES structure containing
            'the Vendor ID, Product ID, and Product Version Number.  Use this information to determine if the detected device 'is the one we're looking for.
            DeviceAttributes.Size = LenB(DeviceAttributes)
            Result = apiHidD_GetAttributes(HIDHandle, DeviceAttributes)
            If (DeviceAttributes.VendorId = VendorId) And (DeviceAttributes.ProductID = ProductID) Then 'Find out if the device matches the one we're looking for.
               MyDeviceDetected = True     'It's the desired device.
               If ProductID = productIdJoyConL Then
                  ChangeList "Joy-Con (Left) detected"
               ElseIf ProductID = productIdJoyConR Then
                  ChangeList "Joy-Con (Right) detected"
               End If
            Else
               MyDeviceDetected = False
               Result = apiCloseHandle(HIDHandle) 'If it's not the one we want, close its handle.
            End If
         End If
         MemberIndex = MemberIndex + 1 'Keep looking until we find the device or there are no more left to examine.
         DoEvents
         If unloading = True Then
            End
         End If
      Loop Until (LastDevice = True) Or (MyDeviceDetected = True)
      If MyDeviceDetected = False Then
         If ProductID = productIdJoyConL Then
            ChangeList "Joy-Con (Left) not detected"
            If cmdPoll.Caption = "Pause" Then
               cmdPoll_Click
            End If
         ElseIf ProductID = productIdJoyConR Then
            ChangeList "Joy-Con (Right) not detected"
            If cmdPoll.Caption = "Pause" Then
               cmdPoll_Click
            End If
         End If
      End If
      Result = apiSetupDiDestroyDeviceInfoList(DeviceInfoSet) 'Free the memory reserved for the DeviceInfoSet returned by SetupDiGetClassDevs.
   End If
   If MyDeviceDetected = True Then
      FindTheHid = True
      GetDeviceCapabilities  'Learn the capabilities of the device 'Get another handle for the overlapped ReadFiles.
      ReadHandle = apiCreateFile(DevicePathName, (GENERIC_READ Or GENERIC_WRITE), (FILE_SHARE_READ Or FILE_SHARE_WRITE), Security, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, 0)
      PrepareForOverlappedTransfer
   End If
End Function
Private Sub ChangeText(ByVal txt As String)
   Text1.Text = Text1.Text & txt & vbCrLf
   If Len(Text1.Text) > 32767 Then
      Text1.Text = VBA.Left(Text1.Text, 32767)
   End If
   Text1.SelStart = Len(Text1.Text) - 1
   Text1.SelLength = 1 'scroll
End Sub
Private Sub ChangeList(ByVal txt As String)
   lstResults.AddItem txt & " " & Now
   'trim list count
   If lstResults.ListCount > 30 Then
      Dim z As Long
      For z = 1 To (lstResults.ListCount - 30)
         lstResults.RemoveItem z - 1
      Next
   End If
   lstResults.ListIndex = lstResults.ListCount - 1 'scroll
End Sub
Private Sub GetDeviceCapabilities()
   Dim ppData(29) As Byte
   Dim ppDataString As Variant
   Result = apiHidD_GetPreparsedData(HIDHandle, PreparsedData) 'Preparsed Data is a pointer to a routine-allocated buffer.
   Result = apiRtlMoveMemory(ppData(0), PreparsedData, 30) 'Copy the data at PreparsedData into a byte array.
   ppDataString = ppData()
   ppDataString = StrConv(ppDataString, vbUnicode)  'Convert the data to Unicode.
   Result = apiHidP_GetCaps(PreparsedData, Capabilities)
   Dim ValueCaps(1023) As Byte
   Result = apiHidP_GetValueCaps(HidP_Input, ValueCaps(0), Capabilities.NumberInputValueCaps, PreparsedData)
   Result = apiHidD_FreePreparsedData(PreparsedData) 'Free the buffer reserved by HidD_GetPreparsedData
End Sub
Private Sub PrepareForOverlappedTransfer()
   'Creates an event object for the overlapped structure used with ReadFile. Requires a security attributes structure or null,
   'Manual Reset = True (ResetEvent resets the manual reset object to nonsignaled),  Initial state = True (signaled), and event object name (optional)   'Returns a handle to the event object.
   If EventObject = 0 Then
      EventObject = apiCreateEvent(Security, True, True, "")
   End If
   HIDOverlapped.Offset = 0
   HIDOverlapped.OffsetHigh = 0
   HIDOverlapped.hEvent = EventObject
End Sub
Private Function GetDataString(Address As Long, Bytes As Long) As String
   Dim Offset As Integer
   Dim Result As String
   Dim ThisByte As Byte
   For Offset = 0 To Bytes - 1
      apiRtlMoveMemory ByVal VarPtr(ThisByte), ByVal Address + Offset, 1
      If (ThisByte And &HF0) = 0 Then
         Result = Result & "0"
      End If
      Result = Result & Hex$(ThisByte) & " "
   Next
   GetDataString = Result$
End Function
Private Function GetErrorString(ByVal LastError As Long) As String
   Dim Bytes As Long
   Dim ErrorString As String
   ErrorString = String$(129, 0)
   Bytes = apiFormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, LastError, 0, ErrorString$, 128, 0)
   If Bytes > 2 Then 'Subtract two characters from the message to strip the CR and LF.
      GetErrorString = Left$(ErrorString, Bytes - 2)
   End If
End Function
'
Private Sub WriteReport()
   Dim Count As Integer
   Dim NumberOfBytesRead As Long
   Dim NumberOfBytesToSend As Long
   Dim NumberOfBytesWritten As Long
   Dim ReadBuffer() As Byte
   Dim SendBuffer() As Byte
   OutputReportData(0) = &H1  'Injects input ???
   '   OutputReportData(1) = &H0
   '   OutputReportData(2) = &H0
   '   OutputReportData(3) = &H0
   '   OutputReportData(4) = &H0
   '   OutputReportData(5) = &H0
   '   OutputReportData(6) = &H0
   '   OutputReportData(7) = &H0
   If Capabilities.OutputReportByteLength > 0 Then  'The SendBuffer array begins at 0, so subtract 1 from the number of bytes.
      ReDim SendBuffer(Capabilities.OutputReportByteLength - 1)
   Else
      ReDim SendBuffer(0)
   End If
   SendBuffer(0) = 0
   If Capabilities.OutputReportByteLength > 0 Then
      For Count = 0 To Capabilities.OutputReportByteLength - 1
         If Count < UBound(OutputReportData) Then
            SendBuffer(Count) = OutputReportData(Count)
         End If
      Next
   End If
   NumberOfBytesWritten = 0
   Result = apiWriteFile(HIDHandle, SendBuffer(0), CLng(Capabilities.OutputReportByteLength), NumberOfBytesWritten, 0) 'Send data to the device.
   ChangeList "OutputReportByteLength = " & Capabilities.OutputReportByteLength
   ChangeList "NumberOfBytesWritten = " & NumberOfBytesWritten
   ChangeList "Report ID: " & SendBuffer(0)
   ChangeList "Report Data:"
End Sub
'
'Left Joycon bits
'1 SL1
'2 SR1
'01 Dpad Left
'02 Dpad Down
'04 Dpad Up
'08 Dpad Right
'002 View
'004 Left Shoulder Bumper
'008 Left Trigger
'0000 None (buttons)
'0001 Minus
'0004 Left Stick
'000000 Thumb Left
'000001 Thumb Up Left
'000002 Thumb Up
'000003 Thumb Up Right
'000004 Thumb Right
'000005 Thumb Down Right
'000006 Thumb Down
'000007 Thumb Down Left
'000008 Thumb Dead Zone
'_________________________________________
'Right Joycon
'1 SL2
'2 SR2
'01 B
'02 Y
'04 A
'08 X
'001 Menu
'004 Right Shoulder Bumper
'008 Right Trigger
'0000 None (buttons)
'0002 Plus
'0008 Right Stick
'000000 Thumb Right
'000001 Thumb Down Right
'000002 Thumb Down
'000003 Thumb Down Left
'000004 Thumb Left
'000005 Thumb Up Left
'000006 Thumb Up
'000007 Thumb Up Right
'000008 Thumb Dead Zone
Private Function GetButtonState(ByRef bv As String) As String
   If LenB(bv) <= 6 Then Exit Function
   Dim bit1 As String
   Dim bit2 As String
   Dim bit3 As String
   Dim bit4 As String
   Dim bit6 As String
   bit1 = VBA.Mid(bv, 1, 1) 'get first bit
   bit2 = VBA.Mid(bv, 2, 1) 'second bit
   bit3 = VBA.Mid(bv, 3, 1)
   bit4 = VBA.Mid(bv, 4, 1)
   bit6 = VBA.Mid(bv, 6, 1) 'sixth bit (thumb 0-8.  Dead=8)
   '
   If DeviceAttributes.ProductID = productIdJoyConL Then
      'thumb
      If CLng("&H" & bit6) = VK_PAD_LTHUMB_RIGHT Then
         GetButtonState = GetButtonState & " + Left-Thumb(Right) "
      ElseIf CLng("&H" & bit6) = VK_PAD_LTHUMB_DOWNRIGHT Then
         GetButtonState = GetButtonState & " + Left-Thumb(Down-Right) "
      ElseIf CLng("&H" & bit6) = VK_PAD_LTHUMB_DOWN Then
         GetButtonState = GetButtonState & " + Left-Thumb(Down) "
      ElseIf CLng("&H" & bit6) = VK_PAD_LTHUMB_DOWNLEFT Then
         GetButtonState = GetButtonState & " + Left-Thumb(Down-Left) "
      ElseIf CLng("&H" & bit6) = VK_PAD_LTHUMB_LEFT Then
         GetButtonState = GetButtonState & " + Left-Thumb(Left) "
      ElseIf CLng("&H" & bit6) = VK_PAD_LTHUMB_UPLEFT Then
         GetButtonState = GetButtonState & " + Left-Thumb(Up-Left) "
      ElseIf CLng("&H" & bit6) = VK_PAD_LTHUMB_UP Then
         GetButtonState = GetButtonState & " + Left-Thumb(Up) "
      ElseIf CLng("&H" & bit6) = VK_PAD_LTHUMB_UPRIGHT Then
         GetButtonState = GetButtonState & " + Left-Thumb(Up-Right) "
      ElseIf CLng("&H" & bit6) = VK_PAD_LTHUMB_DEAD Then
         GetButtonState = GetButtonState & " + Left-Thumb(Dead-Zone) "
      End If
      '
      If CLng("&H" & bit4) And VK_PAD_MINUS Then
         GetButtonState = GetButtonState & " + Minus "
      End If
      If CLng("&H" & bit4) And VK_PAD_LTHUMB_PRESS Then
         GetButtonState = GetButtonState & " + Left-Stick "
      End If
      '
      If CLng("&H" & bit3) And VK_PAD_BACK Then
         GetButtonState = GetButtonState & " + View "
      End If
      '
      If CLng("&H" & bit3) And VK_PAD_LSHOULDER Then
         GetButtonState = GetButtonState & " + Left-Shoulder "
      End If
      If CLng("&H" & bit3) And VK_PAD_LTRIGGER Then
         GetButtonState = GetButtonState & " + Left-Trigger "
      End If
      '
      'Dpad
      If CLng("&H" & bit2) And VK_PAD_DPAD_LEFT Then
         GetButtonState = GetButtonState & " + Dpad-Left "
      End If
      If CLng("&H" & bit2) And VK_PAD_DPAD_DOWN Then
         GetButtonState = GetButtonState & " + Dpad-Down "
      End If
      If CLng("&H" & bit2) And VK_PAD_DPAD_UP Then
         GetButtonState = GetButtonState & " + Dpad-Up "
      End If
      If CLng("&H" & bit2) And VK_PAD_DPAD_RIGHT Then
         GetButtonState = GetButtonState & " + Dpad-Right "
      End If
      '
      'S buttons
      If CLng("&H" & bit1) And VK_PAD_SL1 Then
         GetButtonState = GetButtonState & "SL1 "
      End If
      If CLng("&H" & bit1) And VK_PAD_SR1 Then
         GetButtonState = GetButtonState & "SR1 "
      End If
   ElseIf DeviceAttributes.ProductID = productIdJoyConR Then
      'thumb
      If CLng("&H" & bit6) = VK_PAD_RTHUMB_LEFT Then
         GetButtonState = GetButtonState & " + Right-Thumb(Left) "
      ElseIf CLng("&H" & bit6) = VK_PAD_RTHUMB_UPLEFT Then
         GetButtonState = GetButtonState & " + Right-Thumb(Up-Left) "
      ElseIf CLng("&H" & bit6) = VK_PAD_RTHUMB_UP Then
         GetButtonState = GetButtonState & " + Right-Thumb(Up) "
      ElseIf CLng("&H" & bit6) = VK_PAD_RTHUMB_UPRIGHT Then
         GetButtonState = GetButtonState & " + Right-Thumb(Up-Right) "
      ElseIf CLng("&H" & bit6) = VK_PAD_RTHUMB_RIGHT Then
         GetButtonState = GetButtonState & " + Right-Thumb(Right) "
      ElseIf CLng("&H" & bit6) = VK_PAD_RTHUMB_DOWNRIGHT Then
         GetButtonState = GetButtonState & " + Right-Thumb(Down-Right) "
      ElseIf CLng("&H" & bit6) = VK_PAD_RTHUMB_DOWN Then
         GetButtonState = GetButtonState & " + Right-Thumb(Down) "
      ElseIf CLng("&H" & bit6) = VK_PAD_RTHUMB_DOWNLEFT Then
         GetButtonState = GetButtonState & " + Right-Thumb(Down-Left) "
      ElseIf CLng("&H" & bit6) = VK_PAD_RTHUMB_DEAD Then
         GetButtonState = GetButtonState & " + Right-Thumb(Dead-Zone) "
      End If
      '
      If CLng("&H" & bit4) And VK_PAD_PLUS Then
         GetButtonState = GetButtonState & " + Plus "
      End If
      If CLng("&H" & bit4) And VK_PAD_RTHUMB_PRESS Then
         GetButtonState = GetButtonState & " + Right-Stick "
      End If
      '
      If CLng("&H" & bit3) And VK_PAD_START Then
         GetButtonState = GetButtonState & " + Start "
      End If
      '
      If CLng("&H" & bit3) And VK_PAD_RSHOULDER Then
         GetButtonState = GetButtonState & " + Right-Shoulder "
      End If
      If CLng("&H" & bit3) And VK_PAD_RTRIGGER Then
         GetButtonState = GetButtonState & " + Right-Trigger "
      End If
      '
      'Action buttons
      If CLng("&H" & bit2) And VK_PAD_A Then
         GetButtonState = GetButtonState & " + A "
      End If
      If CLng("&H" & bit2) And VK_PAD_B Then
         GetButtonState = GetButtonState & " + B "
      End If
      If CLng("&H" & bit2) And VK_PAD_X Then
         GetButtonState = GetButtonState & " + X "
      End If
      If CLng("&H" & bit2) And VK_PAD_Y Then
         GetButtonState = GetButtonState & " + Y "
      End If
      'S buttons
      If CLng("&H" & bit1) And VK_PAD_SL2 Then
         GetButtonState = GetButtonState & "SL2 "
      End If
      If CLng("&H" & bit1) And VK_PAD_SR2 Then
         GetButtonState = GetButtonState & "SR2 "
      End If
   End If
   GetButtonState = Trim(GetButtonState)
   If VBA.Left(GetButtonState, 1) = "+" Then
      GetButtonState = VBA.Right(GetButtonState, Len(GetButtonState) - 1)
   End If
End Function
