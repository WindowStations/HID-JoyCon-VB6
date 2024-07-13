VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHIDJoyCon"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
Private Const VendorID As Long = &H57E
Private Const productIdJoyConL As Long = &H2006
Private Const productIdJoyConR As Long = &H2007
Private Const HIDP_INPUT As Integer = 0
Private Const HIDP_OUTPUT As Integer = 1
Private Const HIDP_FEATURE As Integer = 2
Private Const DIGCF_PRESENT As Long = &H2 'setupapi.h
Private Const DIGCF_DEVICEINTERFACE As Long = &H10
Private Const FILE_FLAG_OVERLAPPED As Long = &H40000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const OPEN_EXISTING As Long = &H3
Private Const WAIT_ABANDONED  As Long = &H80
Private Const WAIT_FAILED  As Long = &HFFFFFFFF
Private Const WAIT_OBJECT_0  As Long = &H0
Private Const WAIT_TIMEOUT  As Long = &H102
Private Const INFINITE  As Long = &HFFFF
Private Const ERROR_SUCCESS As Long = &H0
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
Private Const VK_PAD_LTHUMB_PRESS As Long = &H8
Private Const VK_PAD_BACK As Long = &H2
Private Const VK_PAD_LSHOULDER As Long = &H4
Private Const VK_PAD_LTRIGGER As Long = &H8
'initial mode only
'Private Const VK_PAD_DPAD_LEFT As Long = &H1
'Private Const VK_PAD_DPAD_DOWN As Long = &H2
'Private Const VK_PAD_DPAD_UP As Long = &H4
'Private Const VK_PAD_DPAD_RIGHT As Long = &H8
'Private Const VK_PAD_LTHUMB_PRESS As Long = &H4
'Private Const VK_PAD_SL1 As Long = &H1
'Private Const VK_PAD_SR1 As Long = &H2
'
Private Const VK_PAD_DPAD_DOWN As Long = &H1
Private Const VK_PAD_DPAD_UP As Long = &H2
Private Const VK_PAD_DPAD_RIGHT As Long = &H4
Private Const VK_PAD_DPAD_LEFT As Long = &H8
Private Const VK_PAD_SL1 As Long = &H2
Private Const VK_PAD_SR1 As Long = &H1
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
Private Const VK_PAD_RTHUMB_PRESS As Long = &H4
Private Const VK_PAD_START As Long = &H1
Private Const VK_PAD_RSHOULDER As Long = &H4
Private Const VK_PAD_RTRIGGER As Long = &H8
Private Const VK_PAD_B As Long = &H1
Private Const VK_PAD_Y As Long = &H2
Private Const VK_PAD_A As Long = &H4
Private Const VK_PAD_X As Long = &H8
Private Const VK_PAD_SL2 As Long = &H2
Private Const VK_PAD_SR2 As Long = &H1
Private Type GUID
   data1 As Long
   data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Private Type HIDD_ATTRIBUTES
   Size As Long
   VendorID As Integer
   ProductID As Integer
   VersionNumber As Integer
End Type
Private Type HIDP_CAPS
   usage As Integer
   UsagePage As Integer
   InputReportByteLength As Integer
   OutputReportByteLength As Integer
   FeatureReportByteLength As Integer
   reserved(0 To 16) As Integer
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
Private Type HIDP_VALUE_CAPS
   UsagePage As Integer 'USAGE
   ReportID As Byte
   IsAlias As Byte
   BitField As Integer
   LinkCollection As Integer ' A unique internal index pointer
   LinkUsage As Integer 'USAGE
   LinkUsagePage As Integer 'USAGE
   IsRange As Byte 'If IsRange is false, UsageMin is the Usage and UsageMax is unused.
   IsStringRange As Byte 'If IsStringRange is false, StringMin is the string index and StringMax is unused.
   IsDesignatorRange As Byte 'If IsDesignatorRange is false, DesignatorMin is the designator index and DesignatorMax is unused.
   IsAbsolute As Byte
   HasNull As Byte ' Does this channel have a null report   union
   reserved As Byte
   BitSize As Integer ' How many bits are devoted to this value?
   ReportCount As Integer ' See Note below.  Usually set to 1.
   Reserved2(0 To 4) As Integer
   UnitsExp As Long 'Specifies the usage's exponent, as described by the USB HID standard.
   units As Long 'Specifies the usage's units, as described by the USB HID Standard.
   LogicalMin As Long
   LogicalMax As Long
   PhysicalMin As Long
   PhysicalMax As Long
   UsageOrUsageMin As Integer
   UsageMax As Integer
   StringMinOrIndex As Integer
   StringMax As Integer
   DesignatorMinOrIndex As Integer
   DesignatorMax As Integer
   DataIndexMinOrIndex As Integer
   DataIndexMax As Integer
End Type
Private Type OVERLAPPED
   Internal As Long
   InternalHigh As Long
   offset As Long
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
   reserved As Long
End Type
Private Type SP_DEVICE_INTERFACE_DETAIL_DATA
   cbSize As Long
   DevicePath As Byte
End Type
Private Type SP_DEVINFO_DATA
   cbSize As Long
   ClassGuid As GUID
   devInst As Long
   reserved As Long
End Type
Private Declare Function apiIIDFromString Lib "ole32.dll" Alias "IIDFromString" (ByVal lpsz As Long, ByRef lpiid As GUID) As Long
Private Declare Function apiCancelIo Lib "kernel32" Alias "CancelIo" (ByVal hFile As Long) As Long
Private Declare Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private Declare Function apiCreateEvent Lib "kernel32" Alias "CreateEventA" (ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Private Declare Function apiCreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function apiFormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageZId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByVal Arguments As Long) As Long
Private Declare Function apiReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As OVERLAPPED) As Long
Private Declare Function apiResetEvent Lib "kernel32" Alias "ResetEvent" (ByVal hEvent As Long) As Long
Private Declare Function apiRtlMoveMemoryByteStruct Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Byte, ByRef src As SP_DEVICE_INTERFACE_DETAIL_DATA, ByVal Count As Long) As Long
Private Declare Function apiRtlMoveMemoryByteLong Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Byte, ByRef src As Long, ByVal Count As Long) As Long
Private Declare Function apiRtlMoveMemoryLongLong Lib "kernel32" Alias "RtlMoveMemory" (ByRef des As Long, ByRef src As Long, ByVal Count As Long) As Long
Private Declare Function apiCopyMemoryBYTELONG Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Byte, ByVal Source As Long, ByVal Length As Long) As Long
Private Declare Function apiCopyMemoryByteLongLong Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Byte, ByVal Source As Long, ByVal Bytes As Long) As Long
Private Declare Function apiRtlMoveMemoryByteValueCaps Lib "kernel32" Alias "RtlMoveMemory" (ByRef des As HIDP_VALUE_CAPS, ByVal src As Byte, ByVal Count As Long) As Long
Private Declare Function apiWaitForSingleObjectEx Lib "kernel32" Alias "WaitForSingleObjectEx" (ByVal hHandle As Long, ByVal dwMilliseconds As Long, ByVal balertable As Boolean) As Long
Private Declare Function apiWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Long) As Long
Private Declare Function apiWideCharToMultiByte Lib "kernel32" Alias "WideCharToMultiByte" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function apilstrlenA Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function apilstrlenW Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
'
Private Declare Function apiHidD_FreePreparsedData Lib "hid.dll" Alias "HidD_FreePreparsedData" (ByRef PreparsedData As Long) As Long 'ByVal PreparsedData
Private Declare Function apiHidD_GetAttributes Lib "hid.dll" Alias "HidD_GetAttributes" (ByVal HidDeviceObject As Long, ByRef Attributes As HIDD_ATTRIBUTES) As Long
Private Declare Function apiHidP_GetCaps Lib "hid.dll" Alias "HidP_GetCaps" (ByVal PreparsedData As Long, ByRef capabilities As HIDP_CAPS) As Long
Private Declare Function apiHidD_GetHidGuid Lib "hid.dll" Alias "HidD_GetHidGuid" (ByRef HidGuid As GUID) As Long
Private Declare Function apiHidD_GetPreparsedData Lib "hid.dll" Alias "HidD_GetPreparsedData" (ByVal HidDeviceObject As Long, ByRef PreparsedData As Long) As Long
Private Declare Function apiHidP_GetValueCaps Lib "hid.dll" Alias "HidP_GetValueCaps" (ByVal ReportType As Integer, ByRef valuecaps As Byte, ByRef ValueCapsLength As Integer, ByVal PreparsedData As Long) As Long
''Private Const FACILITY_HID_ERROR_CODE = &H11
'? extras untested
Private Declare Function apiHidD_FlushQueue Lib "hid.dll" Alias "HidD_FlushQueue" (ByVal HidDeviceObject As Long) As Byte
''Private Declare Function apiHidD_GetNumInputBuffers Lib "hid.dll" Alias "HidD_GetNumInputBuffers" (ByVal HidDeviceObject As Long, ByRef NumberBuffers As Long) As Byte
''Private Declare Function apiHidD_SetNumInputBuffers Lib "hid.dll" Alias "HidD_SetNumInputBuffers" (ByVal HidDeviceObject As Long, ByVal NumberBuffers As Long) As Byte
''Private Declare Function apiHidD_GetFeature Lib "hid.dll" Alias "HidD_GetFeature" (ByVal HidDeviceObject As Long, ByRef ReportBuffer As Byte, ByVal ReportBufferLength As Long) As Byte
''Private Declare Function apiHidD_SetFeature Lib "hid.dll" Alias "HidD_SetFeature" (ByVal HidDeviceObject As Long, ByRef ReportBuffer As Byte, ByVal ReportBufferLength As Long) As Byte
''Private Declare Function apiHidD_GetInputReport Lib "hid.dll" Alias "HidD_GetInputReport" (ByVal HidDeviceObject As Long, ByRef ReportBuffer As Byte, ByVal ReportBufferLength As Long) As Byte
''Private Declare Function apiHidD_SetOutputReport Lib "hid.dll" Alias "HidD_SetOutputReport" (ByVal HidDeviceObject As Long, ByRef ReportBuffer As Byte, ByVal ReportBufferLength As Long) As Byte
Private Declare Function apiSetupDiDestroyDeviceInfoList Lib "setupapi.dll" Alias "SetupDiDestroyDeviceInfoList" (ByVal DeviceInfoSet As Long) As Long
Private Declare Function apiSetupDiGetClassDevs Lib "setupapi.dll" Alias "SetupDiGetClassDevsA" (ByRef ClassGuid As GUID, ByVal Enumerator As String, ByVal hwndParent As Long, ByVal Flags As Long) As Long
Private Declare Function apiSetupDiEnumDeviceInterfaces Lib "setupapi.dll" Alias "SetupDiEnumDeviceInterfaces" (ByVal DeviceInfoSet As Long, ByVal DeviceInfoData As Long, ByRef InterfaceClassGuid As GUID, ByVal MemberIndex As Long, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA) As Long
Private Declare Function apiSetupDiGetDeviceInterfaceDetail Lib "setupapi.dll" Alias "SetupDiGetDeviceInterfaceDetailA" (ByVal DeviceInfoSet As Long, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA, ByVal DeviceInterfaceDetailData As Long, ByVal DeviceInterfaceDetailDataSize As Long, ByRef RequiredSize As Long, ByVal DeviceInfoData As Long) As Long
Private ReadHandleL As Long
Private ReadHandleR As Long
Public bAlertableL As Long
Public bAlertableR As Long
Private CapabilitiesL As HIDP_CAPS
Private CapabilitiesR As HIDP_CAPS
Private EventObjectL As Long
Private EventObjectR As Long
Private hidHandleL As Long
Private hidHandleR As Long
Private HIDOverlappedL As OVERLAPPED
Private HIDOverlappedR As OVERLAPPED
Private mvarDeviceDetectedL As Boolean
Private mvarDeviceDetectedR As Boolean
Private DevicePathNameL As String
Private DevicePathNameR As String
'
Public Event DeviceConnection(ByVal Index As Integer, ByVal connected As Boolean)
Public Event ClickButton(ByVal joycon As Integer, ByVal down As Boolean, ByVal btns As String)
Public Event WriteReportResult(ByVal Report As String)
Private LTHUMB_RIGHT As Boolean
Private LTHUMB_DOWNRIGHT As Boolean
Private LTHUMB_DOWN As Boolean
Private LTHUMB_DOWNLEFT As Boolean
Private LTHUMB_LEFT As Boolean
Private LTHUMB_UP As Boolean
Private LTHUMB_UPLEFT As Boolean
Private LTHUMB_UPRIGHT As Boolean
Private LTHUMB_DEAD As Boolean
Private JOYCON1_MINUS As Boolean
Private JOYCON1_VIEW As Boolean
Private JOYCON1_LSHOULDER As Boolean
Private JOYCON1_LTRIGGER As Boolean
Private JOYCON1_LSTICK As Boolean
Private JOYCON1_DPAD_UP As Boolean
Private JOYCON1_DPAD_DOWN As Boolean
Private JOYCON1_DPAD_LEFT As Boolean
Private JOYCON1_DPAD_RIGHT As Boolean
Private JOYCON1_SL As Boolean
Private JOYCON1_SR As Boolean
'
Private RTHUMB_RIGHT As Boolean
Private RTHUMB_DOWNRIGHT As Boolean
Private RTHUMB_DOWN As Boolean
Private RTHUMB_DOWNLEFT As Boolean
Private RTHUMB_LEFT As Boolean
Private RTHUMB_UPLEFT As Boolean
Private RTHUMB_UP As Boolean
Private RTHUMB_UPRIGHT As Boolean
Private RTHUMB_DEAD As Boolean
Private JOYCON2_PLUS As Boolean
Private JOYCON2_MENU As Boolean
Private JOYCON2_RSHOULDER As Boolean
Private JOYCON2_RTRIGGER As Boolean
Private JOYCON2_RSTICK As Boolean
Private JOYCON2_A As Boolean
Private JOYCON2_B As Boolean
Private JOYCON2_X As Boolean
Private JOYCON2_Y As Boolean
Private JOYCON2_SL As Boolean
Private JOYCON2_SR As Boolean
Private WithEvents tmrPoll As Timer
Attribute tmrPoll.VB_VarHelpID = -1
Private WithEvents tmrFindDevices As Timer
Attribute tmrFindDevices.VB_VarHelpID = -1
Public Sub Class_Initialize()
   Set tmrFindDevices = New Timer
   Set tmrPoll = New Timer
   tmrFindDevices.Interval = 3000
   tmrPoll.Interval = 1
   tmrFindDevices.Enabled = True
   tmrPoll.Enabled = True
End Sub
Public Sub FormUnloading()
   If hidHandleL <> 0 Then apiCloseHandle hidHandleL    'Actions that must execute when the program ends.   'Close the open handles to the device.
    If hidHandleR <> 0 Then apiCloseHandle hidHandleR
End Sub
Private Sub tmrFindDevices_Timer()
   FindHID
End Sub
Private Sub tmrPoll_Timer()
   ReadDevices
End Sub
Public Sub FindHID()
   FindJoyConHIDs
End Sub
Public Sub FlushQueue()
   If hidHandleL <> 0 Then apiHidD_FlushQueue hidHandleL
   If hidHandleR <> 0 Then apiHidD_FlushQueue hidHandleR
End Sub
Public Sub ReadDevices()
   'Read a report from device.
   If mvarDeviceDetectedL = True And mvarDeviceDetectedR = True Then
      ReadReport ReadHandleL, CapabilitiesL, EventObjectL, bAlertableL, HIDOverlappedL
      ReadReport ReadHandleR, CapabilitiesR, EventObjectR, bAlertableR, HIDOverlappedR
   ElseIf mvarDeviceDetectedL = True And mvarDeviceDetectedR = False Then
      ReadReport ReadHandleL, CapabilitiesL, EventObjectL, bAlertableL, HIDOverlappedL
   ElseIf mvarDeviceDetectedL = False And mvarDeviceDetectedR = True Then
      ReadReport ReadHandleR, CapabilitiesR, EventObjectR, bAlertableR, HIDOverlappedR
   End If
End Sub
Public Sub WriteReadDevices(ByVal lpData As String)
   FindJoyConHIDs
   If mvarDeviceDetectedL = True Then
      If hidHandleL <> 0 Then
         WriteReport lpData, hidHandleL, CapabilitiesL
         ReadReport ReadHandleL, CapabilitiesL, EventObjectL, bAlertableL, HIDOverlappedL
      End If
   End If
   If mvarDeviceDetectedR = True Then
      If hidHandleR <> 0 Then
         WriteReport lpData, hidHandleR, CapabilitiesR
         ReadReport ReadHandleR, CapabilitiesR, EventObjectR, bAlertableR, HIDOverlappedR
      End If
   End If
End Sub
Private Sub FindJoyConHIDs()
   Dim hidHandle As Long
   Dim SECURITY As SECURITY_ATTRIBUTES
   Dim b As Byte
   If VBA.Trim(DevicePathNameL) <> "" Then
      'if we got a name from previous detail call, we may not need to find again
      hidHandle = apiCreateFile(DevicePathNameL, GENERIC_READ Or GENERIC_WRITE, (FILE_SHARE_READ Or FILE_SHARE_WRITE), SECURITY, OPEN_EXISTING, 0&, 0)
      If hidHandle <> 0 Then
         mvarDeviceDetectedL = True
         apiCloseHandle hidHandle
         b = b + 1
      Else
         hidHandleL = 0
         mvarDeviceDetectedL = False
      End If
   End If
   If VBA.Trim(DevicePathNameR) <> "" Then
      'if we got a name from previous detail call
      hidHandle = apiCreateFile(DevicePathNameR, GENERIC_READ Or GENERIC_WRITE, (FILE_SHARE_READ Or FILE_SHARE_WRITE), SECURITY, OPEN_EXISTING, 0&, 0)
      If hidHandle <> 0 Then
         mvarDeviceDetectedR = True
         apiCloseHandle hidHandle
         b = b + 1
      Else
          hidHandleR = 0
         mvarDeviceDetectedR = False
      End If
   End If
   If b = 2 Then Exit Sub 'We already have device handles
   Dim i As Integer
   Dim GUIDString As String
   Dim HidGuid As GUID
   Dim MemberIndex As Long
   Dim ret As Long
   Dim MyDeviceInterfaceData As SP_DEVICE_INTERFACE_DATA
   Dim MyDeviceInfoData As SP_DEVINFO_DATA
   Dim MyDeviceInterfaceDetailData As SP_DEVICE_INTERFACE_DETAIL_DATA
   Dim Needed As Long
   Dim DataString As String
   Dim DetailData As Long
   Dim DetailDataBuffer() As Byte
   Dim DeviceAttributes As HIDD_ATTRIBUTES
   Dim DevicePathName As String
   Dim DeviceInfoSet As Long
   Dim LastDevice As Boolean
   Dim SECURITYL As SECURITY_ATTRIBUTES
   Dim SECURITYR As SECURITY_ATTRIBUTES
   Dim PreparsedData As Long
   SECURITYL.lpSecurityDescriptor = 0
   SECURITYL.bInheritHandle = True
   SECURITYL.nLength = Len(SECURITYL)
   SECURITYR.lpSecurityDescriptor = 0
   SECURITYR.bInheritHandle = True
   SECURITYR.nLength = Len(SECURITYR)
   ret = apiHidD_GetHidGuid(HidGuid)
   GUIDString = VBA.Hex$(HidGuid.data1) & "-" & VBA.Hex$(HidGuid.data2) & "-" & VBA.Hex$(HidGuid.Data3) & "-"
   For i = 0 To 7
      If HidGuid.Data4(i) >= &H10 Then 'Ensure that each of the 8 bytes in the GUID displays two characters
         GUIDString = GUIDString & VBA.Hex$(HidGuid.Data4(i)) & " "
      Else
         GUIDString = GUIDString & "0" & VBA.Hex$(HidGuid.Data4(i)) & " "
      End If
   Next
   DeviceInfoSet = apiSetupDiGetClassDevs(HidGuid, vbNullString, 0, (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))
   DataString = GetDataString(DeviceInfoSet, 32)
   MemberIndex = 0
   If DeviceInfoSet <> 0 Then ' if device info set can be obtained
      Do
         MyDeviceInterfaceData.cbSize = LenB(MyDeviceInterfaceData)   'The cbSize element of the MyDeviceInterfaceData structure must be set to the structure's size in bytes. The size is 28 bytes.
         ret = apiSetupDiEnumDeviceInterfaces(DeviceInfoSet, 0, HidGuid, MemberIndex, MyDeviceInterfaceData)
         If ret = 0 Then
            LastDevice = True ' no more devices to loop through
         Else 'If a device exists during enumeration
            MyDeviceInfoData.cbSize = Len(MyDeviceInfoData)
            ret = apiSetupDiGetDeviceInterfaceDetail(DeviceInfoSet, MyDeviceInterfaceData, 0, 0, Needed, 0)
            If ret = ERROR_SUCCESS Then 'if we obtained the device's interface details
               DetailData = Needed
               MyDeviceInterfaceDetailData.cbSize = Len(MyDeviceInterfaceDetailData)
               ReDim DetailDataBuffer(Needed) 'Use a byte array to allocate memory for the MyDeviceInterfaceDetailData structure
               apiRtlMoveMemoryByteStruct DetailDataBuffer(0), MyDeviceInterfaceDetailData, 4 'Store cbSize in the first four bytes of the array
               ret = apiSetupDiGetDeviceInterfaceDetail(DeviceInfoSet, MyDeviceInterfaceData, VarPtr(DetailDataBuffer(0)), DetailData, Needed, 0)
               DevicePathName = CStr(DetailDataBuffer())  'Convert the byte array to a string.
               DevicePathName = StrConv(DevicePathName, vbUnicode)  'Convert to Unicode.
               DevicePathName = Right$(DevicePathName, Len(DevicePathName) - 4) 'Strip cbSize (4 bytes) from the beginning.
               If VBA.Trim(DevicePathName) <> "" Then 'if we got a name from detail call
                  hidHandle = apiCreateFile(DevicePathName, GENERIC_READ Or GENERIC_WRITE, (FILE_SHARE_READ Or FILE_SHARE_WRITE), SECURITY, OPEN_EXISTING, 0&, 0)
                  If hidHandle <> 0 Then 'if we found a handle to the device
                     DeviceAttributes.Size = LenB(DeviceAttributes)
                     ret = apiHidD_GetAttributes(hidHandle, DeviceAttributes)
                     If ret <> 0 Then 'If we got the device's attributes
                        If DeviceAttributes.VendorID = VendorID Then
                           If DeviceAttributes.ProductID = productIdJoyConL Then   'If the device matches
                              mvarDeviceDetectedL = True 'Signal a flag for the desired device
                              hidHandleL = hidHandle
                              SECURITYL = SECURITY
                              DevicePathNameL = DevicePathName
                              RaiseEvent DeviceConnection(1, True) 'expose the class object event to application developers
                           ElseIf DeviceAttributes.ProductID = productIdJoyConR Then
                              mvarDeviceDetectedR = True
                              hidHandleR = hidHandle
                              SECURITYR = SECURITY
                              DevicePathNameR = DevicePathName
                              RaiseEvent DeviceConnection(2, True)
                           Else
                              apiCloseHandle hidHandle 'If not matching, close handle
                           End If
                        Else
                           apiCloseHandle hidHandle 'If not matching, close handle
                        End If
                     Else
                        apiCloseHandle hidHandle
                     End If
                  End If
               End If
            End If
         End If
         If LastDevice = True Then Exit Do
         If mvarDeviceDetectedL = True And mvarDeviceDetectedR = True Then Exit Do
         MemberIndex = MemberIndex + 1 'Loop through devices
      Loop
      If mvarDeviceDetectedL = False Then
         RaiseEvent DeviceConnection(1, False)
      End If
      If mvarDeviceDetectedR = False Then
         RaiseEvent DeviceConnection(2, False)
      End If
      apiSetupDiDestroyDeviceInfoList DeviceInfoSet  'Free the memory reserved for the DeviceInfoSet returned by SetupDiGetClassDevs
   End If
   Dim ppData(29) As Byte
   Dim ppDataString As Variant
   Dim valuecaps() As Byte 'Dim valuecaps As HIDP_VALUE_CAPS
   If mvarDeviceDetectedL = True Then
      ' Learn the capabilities of the device 'Get another handle for the overlapped ReadFiles.
      If hidHandleL <> 0 Then
         ret = apiHidD_GetPreparsedData(hidHandleL, PreparsedData) 'Preparsed Data is a pointer to a routine-allocated buffer.
         If ret <> 0 Then
            ret = apiRtlMoveMemoryByteLong(ppData(0), PreparsedData, 30) 'Copy the data at PreparsedData into a byte array.
            If ret <> 0 Then
               ppDataString = ppData()
               ppDataString = StrConv(ppDataString, vbUnicode)  'Convert the data to Unicode.
               ret = apiHidP_GetCaps(PreparsedData, CapabilitiesL)
               If ret <> 0 Then
                  ReDim valuecaps(4096)
                  apiHidP_GetValueCaps HIDP_INPUT, valuecaps(0), CapabilitiesL.NumberInputValueCaps, PreparsedData 'To use this data, copy the byte array into an array of structures.
                  ''apiRtlMoveMemoryByteValueCaps vc, valuecaps, Len(vc)
                  '               Debug.Print Capabilities.FeatureReportByteLength '0
                  '               Debug.Print Capabilities.InputReportByteLength '362
                  '               Debug.Print Capabilities.NumberFeatureButtonCaps '0
                  '               Debug.Print Capabilities.NumberFeatureDataIndices '0
                  '               Debug.Print Capabilities.NumberFeatureValueCaps '0
                  '               Debug.Print Capabilities.NumberInputButtonCaps '1
                  '               Debug.Print Capabilities.NumberInputDataIndices '26
                  '               Debug.Print Capabilities.NumberInputValueCaps '10
                  '               Debug.Print Capabilities.NumberLinkCollectionNodes '1
                  '               Debug.Print Capabilities.NumberOutputButtonCaps '0
                  '               Debug.Print Capabilities.NumberOutputDataIndices '4
                  '               Debug.Print Capabilities.NumberOutputValueCaps '4
                  '               Debug.Print Capabilities.OutputReportByteLength '49
                  '               '''Debug.Print Capabilities.Reserved'NO
                  '               Debug.Print Capabilities.usage '5
                  '               Debug.Print Capabilities.UsagePage '1
                  '               Debug.Print "Value Caps= " & GetDataString((VarPtr(valuecaps(0))), 180)
                  apiHidD_FreePreparsedData PreparsedData  'Free the buffer reserved by HidD_GetPreparsedData
               End If
            End If
         End If
      End If
      If VBA.Trim(DevicePathNameL) <> "" Then
         ReadHandleL = apiCreateFile(DevicePathNameL, (GENERIC_READ Or GENERIC_WRITE), (FILE_SHARE_READ Or FILE_SHARE_WRITE), SECURITY, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, 0)
      End If
      If EventObjectL = 0 Then 'Creates an event object for the overlapped structure used with ReadFile. Requires a security attributes structure or null,
         EventObjectL = apiCreateEvent(SECURITY, True, True, "")
      End If 'Manual Reset = True (ResetEvent resets the manual reset object to nonsignaled),  Initial state = True (signaled), and event object name (optional)
      HIDOverlappedL.offset = 0
      HIDOverlappedL.OffsetHigh = 0
      HIDOverlappedL.hEvent = EventObjectL
   End If
   If mvarDeviceDetectedR = True Then
      ' Learn the capabilities of the device 'Get another handle for the overlapped ReadFiles.
      If hidHandleR <> 0 Then
         ret = apiHidD_GetPreparsedData(hidHandleR, PreparsedData) 'Preparsed Data is a pointer to a routine-allocated buffer.
         If ret <> 0 Then
            'Dim ppData(29) As Byte
            ret = apiRtlMoveMemoryByteLong(ppData(0), PreparsedData, 30) 'Copy the data at PreparsedData into a byte array.
            If ret <> 0 Then
               'Dim ppDataString As Variant
               ppDataString = ppData()
               ppDataString = StrConv(ppDataString, vbUnicode)  'Convert the data to Unicode.
               ret = apiHidP_GetCaps(PreparsedData, CapabilitiesR)
               If ret <> 0 Then
                  'Dim valuecaps() As Byte 'Dim valuecaps As HIDP_VALUE_CAPS
                  ReDim valuecaps(4096)
                  apiHidP_GetValueCaps HIDP_INPUT, valuecaps(0), CapabilitiesR.NumberInputValueCaps, PreparsedData 'To use this data, copy the byte array into an array of structures.
                  apiHidD_FreePreparsedData PreparsedData  'Free the buffer reserved by HidD_GetPreparsedData
               End If
            End If
         End If
      End If
      If VBA.Trim(DevicePathNameR) <> "" Then
         ReadHandleR = apiCreateFile(DevicePathNameR, (GENERIC_READ Or GENERIC_WRITE), (FILE_SHARE_READ Or FILE_SHARE_WRITE), SECURITY, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, 0)
      End If
      If EventObjectR = 0 Then 'Creates an event object for the overlapped structure used with ReadFile. Requires a security attributes structure or null,
         EventObjectR = apiCreateEvent(SECURITY, True, True, "")
      End If 'Manual Reset = True (ResetEvent resets the manual reset object to nonsignaled),  Initial state = True (signaled), and event object name (optional)
      HIDOverlappedR.offset = 0
      HIDOverlappedR.OffsetHigh = 0
      HIDOverlappedR.hEvent = EventObjectR
   End If
End Sub
Private Sub ReadReport(ByRef readHandle As Long, ByRef capabilities As HIDP_CAPS, ByRef eventobject As Long, ByVal balertable As Boolean, ByRef hidoverlapped As OVERLAPPED)
   Dim i As Integer
   Dim NumberOfBytesRead As Long
   Dim ReadBuffer() As Byte
   Dim UBoundReadBuffer As Integer
   Dim ByteValue As String
   Dim ret As Long
   If capabilities.InputReportByteLength > 0 Then 'The ReadBuffer array begins at 0, so subtract 1 from the number of bytes.
      ReDim ReadBuffer(capabilities.InputReportByteLength - 1)
   Else
      ReDim ReadBuffer(0)
   End If
   ret = apiReadFile(readHandle, ReadBuffer(0), CLng(capabilities.InputReportByteLength), NumberOfBytesRead, hidoverlapped)   'Do an overlapped ReadFile. 'The function returns immediately, even if the data hasn't been received yet.
   If ret = ERROR_SUCCESS Then
      'balertable = True
      ret = apiWaitForSingleObjectEx(eventobject, 0, balertable)
      apiResetEvent eventobject
      balertable = False
      Select Case ret 'Find out if ReadFile completed or timeout.
      Case WAIT_OBJECT_0
         'ReadFile has completed success
      Case WAIT_TIMEOUT
         ret = apiCancelIo(readHandle) 'Returns non-zero on success.
      Case WAIT_ABANDONED
         If readHandle = ReadHandleL Then
            mvarDeviceDetectedL = False
         ElseIf readHandle = ReadHandleR Then
            mvarDeviceDetectedR = False
         End If
      Case WAIT_FAILED
         If readHandle = ReadHandleL Then
            mvarDeviceDetectedL = False
         ElseIf readHandle = ReadHandleR Then
            mvarDeviceDetectedR = False
         End If
      Case Else
         If readHandle = ReadHandleL Then
            mvarDeviceDetectedL = False
         ElseIf readHandle = ReadHandleR Then
            mvarDeviceDetectedR = False
         End If
      End Select
   End If
   If UBound(ReadBuffer) > 0 And ReadBuffer(0) > 0 Then
      Dim bv As String
      Dim bv2 As String
      For i = 0 To UBound(ReadBuffer)
         If Len(VBA.Hex$(ReadBuffer(i))) < 2 Then 'Add a leading 0 to values 0 - Fh.
            ByteValue = "0" & VBA.Hex$(ReadBuffer(i))
         Else
            ByteValue = VBA.Hex$(ReadBuffer(i))
         End If
         bv = bv & ByteValue
         If i < 50 Then
            bv2 = bv2 & ByteValue & " "
         End If
         If i > 361 Then '361 seems to be the limit-ish for a pair of joycons. 49-361 (ID x31) NFC/IR data input report. Max 313 bytes.
            Exit For
         End If
      Next
      If readHandle = ReadHandleL Then
         RaiseEvents bv, "2006", bv2
      ElseIf readHandle = ReadHandleR Then
         RaiseEvents bv, "2007", bv2
      End If
   End If
End Sub
Private Function ConvertByteToBit(ByRef byt As Byte) As String
    'Bit  ==  32103210
    '&HF0 == "11110000"
    '&H0A == "00001010"
    Dim e As Integer
    Dim z As Integer
    Dim valbyte As Integer
    Dim bit As String
    valbyte = VBA.Val(CStr(byt))
    For e = 0 To 7
        z = 2 ^ e
        If valbyte And z Then
            bit = "1" & bit
        Else
            bit = "0" & bit
        End If
    Next
    ConvertByteToBit = bit
End Function
Private Sub RaiseEvents(ByVal bv As String, ByVal mvarProductId As String, ByVal bv2 As String)
   If VBA.Trim(bv) = "" Then Exit Sub ' no events to raise
   Dim x As String
   Dim y As String
   Dim flag1 As Long
   Dim flag2 As Long
   Dim flag3 As Long
   Dim flag4 As Long
   Dim flag5 As Long
   Dim flag6 As Long
   Dim flag7 As Long
   Dim flag8 As Long
   Dim flag9 As Long
   Dim flag10 As Long
   Dim flag11 As Long
   Dim flag12 As Long
   Dim accel_x As Long ' IMU (6-Axis sensor)
   Dim accel_y As Long
   Dim accel_z As Long
   Dim gyro_1 As Long
   Dim gyro_2 As Long
   Dim gyro_3 As Long
   Dim imu As String
   Dim bState As String
   Dim ss As String
   Dim bys() As String
   flag1 = CLng("&H" & VBA.Mid(bv, 1, 1))
   flag2 = CLng("&H" & VBA.Mid(bv, 2, 1))
   flag3 = CLng("&H" & VBA.Mid(bv, 3, 1))
   flag4 = CLng("&H" & VBA.Mid(bv, 4, 1))
   flag5 = CLng("&H" & VBA.Mid(bv, 5, 1))
   flag6 = CLng("&H" & VBA.Mid(bv, 6, 1))
   flag7 = CLng("&H" & VBA.Mid(bv, 7, 1))
   flag8 = CLng("&H" & VBA.Mid(bv, 8, 1)) 'eight bit (thumb 0-8.  Dead=8) 3F
   flag9 = CLng("&H" & VBA.Mid(bv, 9, 1)) 'View button
   flag10 = CLng("&H" & VBA.Mid(bv, 10, 1)) 'Minus-button=1 and Left-stick button=8
   flag11 = CLng("&H" & VBA.Mid(bv, 11, 1)) 'SR=1,SL=2, Left-shoulder=4 and Left-trigger=8
   flag12 = CLng("&H" & VBA.Mid(bv, 12, 1)) 'Dpad=2  Left JoyCon
   'If 6-axis sensor is enabled, the IMU data in an 0x30, 0x31, 0x32 and 0x33 input report is packaged
   'like this (assuming the packet ID is located at byte 0):
   '25-48   The data is repeated 2 more times. Each with 5ms ?t sampling.
   accel_x = CLng("&H" & VBA.Mid(bv, 27, 4))
   accel_y = CLng("&H" & VBA.Mid(bv, 29, 4))
   accel_z = CLng("&H" & VBA.Mid(bv, 31, 4))
   gyro_1 = CLng("&H" & VBA.Mid(bv, 33, 4))
   gyro_2 = CLng("&H" & VBA.Mid(bv, 35, 4))
   gyro_3 = CLng("&H" & VBA.Mid(bv, 37, 4))
   If mvarProductId = "2006" Then 'Left JoyCon
      If VBA.Left(bv, 2) = "3F" Then 'Basic mode
         bv = "(JoyCon L) (Basic Input Mode: " & VBA.Left(bv, 2) & ") (Buttons: " & VBA.Mid(bv, 3, 6) & ")"
         bv = bv & "  (Raw: " & bv2 & ")" & vbCrLf
      Else 'fast polling mode 30, 31, 32 and 33
         'Nintendo has the order scrambled encoded.  Unscramble and convert to decimal.  Messy. why did they do this?
         x = VBA.Mid(bv, 16, 1) & VBA.Mid(bv, 13, 1) & VBA.Mid(bv, 14, 1) ' & VBA.Mid(bv, 14, 1)
         y = VBA.Mid(bv, 17, 1) & VBA.Mid(bv, 18, 1) & VBA.Mid(bv, 15, 1) '& VBA.Mid(bv, 15, 1)
         x = CStr(VBA.Round(1.5 * CLng("&H" & x & "0")) - 50000)
         y = CStr(VBA.Round(1.5 * CLng("&H" & y & "0")) - 50000)
         If VBA.Abs(CLng(x)) >= 32767 Then
            If VBA.Abs(CLng(x)) = CLng(x) Then
               x = CStr(32767)
            Else
               x = CStr(-32767)
            End If
         End If
         If VBA.Abs(CLng(y)) >= 32767 Then
            If VBA.Abs(CLng(y)) = CLng(y) Then
               y = CStr(32767)
            Else
               y = CStr(-32767)
            End If
         End If
         If accel_x <> 0 And accel_y <> 0 And accel_z <> 0 And gyro_1 <> 0 And gyro_2 <> 0 And gyro_3 <> 0 Then
            If VBA.Left(bv, 2) = "21" Then
               bys = Split(bv2, " ")
               ss = "(Firmware Version: " & bys(15) & "." & bys(16) & ") " & "(JoyCon: " & bys(17) & ") "
               ss = ss & "(MAC address: " & bys(19) & bys(20) & bys(21) & bys(22) & bys(23) & bys(24)
               '
            Else
               imu = "IMU 6-axis Mode: " & "(Accel X: " & CStr(accel_x) & ")  (Accel Y: " & CStr(accel_y) & ")  (Accel Z: " & CStr(accel_z) & ")  (Gyro 1: " & CStr(gyro_1) & ")  (Gyro 2: " & CStr(gyro_2) & ")  (Gyro 3: " & CStr(gyro_3) & ")"
            End If
         End If
         bv = "(JoyCon L) (Input Mode: " & VBA.Left(bv, 2) & ") (Packet: " & VBA.Mid(bv, 3, 2) & ") (Battery: " & VBA.Mid(bv, 5, 1) & ") (Connection:" & VBA.Mid(bv, 6, 1) & ") (Buttons: " & VBA.Mid(bv, 9, 4) & ") (X,Y: " & x & "," & y & ") (L/R: " & VBA.Mid(bv, 25, 1) & "/" & VBA.Mid(bv, 26, 1) & ")"
         If ss <> "" Then bv = ss & " " & bv
         If imu <> "" Then bv = imu & " " & bv
         bv = bv & "  (Raw: " & bv2 & ")"
      End If
   ElseIf mvarProductId = "2007" Then 'Right JoyCon
      If VBA.Left(bv, 2) = "3F" Then
         bv = "(JoyCon R) (Basic Input Mode: " & VBA.Left(bv, 2) & ") (Buttons: " & VBA.Mid(bv, 3, 6) & ")"
         bv = bv & "  (Raw: " & bv2 & ")" & vbCrLf
      Else
         'Nintendo scrambled the order.  Unscramble and convert  Ugly
         x = VBA.Mid(bv, 22, 1) & VBA.Mid(bv, 19, 1) & VBA.Mid(bv, 21, 1)
         y = VBA.Mid(bv, 23, 1) & VBA.Mid(bv, 24, 1) & VBA.Mid(bv, 20, 1)
         
         x = CStr(VBA.Round(1.5 * CLng("&H" & x & "0")) - 50000)
         y = CStr(VBA.Round(1.5 * CLng("&H" & y & "0")) - 50000)
         If VBA.Abs(CLng(x)) >= 32767 Then
            If VBA.Abs(CLng(x)) = CLng(x) Then
               x = CStr(32767)
            Else
               x = CStr(-32767)
            End If
         End If
         If VBA.Abs(CLng(y)) >= 32767 Then
            If VBA.Abs(CLng(y)) = CLng(y) Then
               y = CStr(32767)
            Else
               y = CStr(-32767)
            End If
         End If
         If accel_x <> 0 And accel_y <> 0 And accel_z <> 0 And gyro_1 <> 0 And gyro_2 <> 0 And gyro_3 <> 0 Then
            If VBA.Left(bv, 2) = "21" Then
               bys = Split(bv2, " ")
               ss = "(Firmware Version: " & bys(15) & "." & bys(16) & ") " & "(JoyCon: " & bys(17) & ") "
               ss = ss & "(MAC address: " & bys(19) & bys(20) & bys(21) & bys(22) & bys(23) & bys(24)
            Else
               imu = "IMU 6-axis Mode: " & "(Accel X: " & CStr(accel_x) & ")  (Accel Y: " & CStr(accel_y) & ")  (Accel Z: " & CStr(accel_z) & ")  (Gyro 1: " & CStr(gyro_1) & ")  (Gyro 2: " & CStr(gyro_2) & ")  (Gyro 3: " & CStr(gyro_3) & ")"
            End If
         End If
         bv = "(JoyCon R) (Input Mode: " & VBA.Left(bv, 2) & ") (Packet: " & VBA.Mid(bv, 3, 2) & ") (Battery: " & VBA.Mid(bv, 5, 1) & ") (Connection:" & VBA.Mid(bv, 6, 1) & ") (Buttons: " & VBA.Mid(bv, 7, 4) & ") (X,Y: " & x & "," & y & ") (L/R: " & VBA.Mid(bv, 25, 1) & "/" & VBA.Mid(bv, 26, 1) & ")"
         If ss <> "" Then bv = ss & " " & bv
         If imu <> "" Then bv = imu & " " & bv
         bv = bv & "  (Raw: " & bv2 & ")"
      End If
   End If
   frmMain.ChangeText bv
End Sub
Public Function ProductIdJoyConLeft() As Long
   ProductIdJoyConLeft = productIdJoyConL
End Function
Public Function ProductIdJoyConRight() As Long
   ProductIdJoyConRight = productIdJoyConR
End Function
Public Sub EnablePoll()
   tmrPoll.Enabled = True
End Sub
Public Sub DisablePoll()
   tmrPoll.Enabled = False
End Sub
Public Sub EnableFindDevices()
   tmrFindDevices.Enabled = True
End Sub
Public Sub DisableFindDevices()
   tmrFindDevices.Enabled = False
End Sub
Private Function GetDataString(ByRef Address As Long, Bytes As Long) As String
   Dim offset As Integer
   Dim ret As String
   Dim ThisByte As Byte
   For offset = 0 To Bytes - 1
      apiRtlMoveMemoryLongLong ByVal VarPtr(ThisByte), ByVal Address + offset, 1
      If (ThisByte And &HF0) = 0 Then
         ret = ret & "0"
      End If
      ret = ret & VBA.Hex$(ThisByte) & " "
   Next
   GetDataString = ret
End Function
Private Function GetStringFromLP(ByVal SPtr As Long) As String
   Dim b As Byte
   Dim txt As String
   Do
      apiCopyMemoryBYTELONG b, SPtr, 1         ' Get the byte/character that StrPtr is pointing to.
      If b = 0 Then Exit Do  ' If you've found a null character, then you're done.
      txt = txt & VBA.Chr(b)  ' Get the character for the byte's value'Add it to the string
      SPtr = SPtr + 1  ' Increment the pointer to next byte/char
   Loop
   GetStringFromLP = txt
End Function
Public Function PtrToStringA(ByRef lpszA As Long) As String
   If lpszA = 0 Then Exit Function
   Dim nLen As Long
   nLen = apilstrlenA(ByVal lpszA)
   If nLen = 0 Then Exit Function
   Dim buff() As Byte
   ReDim buff(0 To (nLen - 1)) As Byte
   apiCopyMemoryByteLongLong buff(0), lpszA, nLen
   PtrToStringA = StrConv(buff, vbUnicode)
End Function
Public Function PtrToStringW(ByRef lpszW As Long) As String
   Dim s As String
   Const CP_ACP As Long = 0&
   s = String$(apilstrlenW(ByVal lpszW) * 2&, vbNullChar)
   apiWideCharToMultiByte CP_ACP, 0&, ByVal lpszW, -1, ByVal s, Len(s), 0&, 0&
   PtrToStringW = Left$(s, apilstrlenW(StrPtr(s)))
End Function
Private Sub WriteReport(ByVal lpData As String, ByRef hidHandle As Long, ByRef capabilities As HIDP_CAPS)
   If capabilities.OutputReportByteLength < 11 Then Exit Sub
   Dim i As Integer
   Dim NumberOfBytesRead As Long
   Dim NumberOfBytesToSend As Long
   Dim NumberOfBytesWritten As Long
   Dim ReadBuffer() As Byte
   Dim SendBuffer() As Byte
   ReDim SendBuffer(capabilities.OutputReportByteLength)
   Dim byts() As String
   lpData = VBA.Replace(lpData, "   ", " ")
   lpData = VBA.Replace(lpData, "  ", " ")
   byts = VBA.Split(lpData, " ")
   For i = 0 To UBound(byts)
      If VBA.Trim(byts(i)) <> "" Then
         SendBuffer(i) = CByte(CLng("&H" & VBA.Trim(byts(i))))
      End If
   Next
   NumberOfBytesWritten = 0
   Dim nLen As Long
   nLen = UBound(SendBuffer) - LBound(SendBuffer) + 1
   apiWriteFile hidHandle, SendBuffer(1), nLen, NumberOfBytesWritten, 0  'Send data to the device.
   'Dim Report As String
   ' Report = Report & "NumberOfBytesWritten = " & NumberOfBytesWritten
   ' RaiseEvent WriteReportResult(Report)
End Sub
'' 3F mode
''Left Joycon bits
''1 SL1
''2 SR1
''01 Dpad Left
''02 Dpad Down
''04 Dpad Up
''08 Dpad Right
''002 View
''004 Left Shoulder Bumper
''008 Left Trigger
''0000 None (buttons)
''0001 Minus
''0004 Left Stick
''000000 Thumb Left
''000001 Thumb Up Left
''000002 Thumb Up
''000003 Thumb Up Right
''000004 Thumb Right
''000005 Thumb Down Right
''000006 Thumb Down
''000007 Thumb Down Left
''000008 Thumb Dead Zone
''_________________________________________
''Right Joycon
''1 SL2
''2 SR2
''01 B
''02 Y
''04 A
''08 X
''001 Menu
''004 Right Shoulder Bumper
''008 Right Trigger
''0000 None (buttons)
''0002 Plus
''0008 Right Stick
''000000 Thumb Right
''000001 Thumb Down Right
''000002 Thumb Down
''000003 Thumb Down Left
''000004 Thumb Left
''000005 Thumb Up Left
''000006 Thumb Up
''000007 Thumb Up Right
''000008 Thumb Dead Zone
'
''Private Const HIDP_LINK_COLLECTION_ROOT  As Integer = (-1)
''Private Const HIDP_LINK_COLLECTION_UNSPECIFIED  As Integer = (0)
''Private Const HIDP_STATUS_SUCCESS = &H110000
''Private Const HIDP_STATUS_NULL = &H80110001
''Private Const HIDP_STATUS_INVALID_PREPARSED_DATA = &HC0110001
''Private Const HIDP_STATUS_INVALID_REPORT_TYPE = &HC0110002
''Private Const HIDP_STATUS_INVALID_REPORT_LENGTH = &HC0110003
''Private Const HIDP_STATUS_USAGE_NOT_FOUND = &HC0110004
''Private Const HIDP_STATUS_VALUE_OUT_OF_RANGE = &HC0110005
''Private Const HIDP_STATUS_BAD_LOG_PHY_VALUES = &HC0110006
''Private Const HIDP_STATUS_BUFFER_TOO_SMALL = &HC0110007
''Private Const HIDP_STATUS_INTERNAL_ERROR = &HC0110008
''Private Const HIDP_STATUS_I8042_TRANS_UNKNOWN = &HC0110009
''Private Const HIDP_STATUS_INCOMPATIBLE_REPORT_ID = &HC011000A
''Private Const HIDP_STATUS_NOT_VALUE_ARRAY = &HC011000B
''Private Const HIDP_STATUS_IS_VALUE_ARRAY = &HC011000C
''Private Const HIDP_STATUS_DATA_INDEX_NOT_FOUND = &HC011000D
''Private Const HIDP_STATUS_DATA_INDEX_OUT_OF_RANGE = &HC011000E
''Private Const HIDP_STATUS_BUTTON_NOT_PRESSED = &HC011000F
''Private Const HIDP_STATUS_REPORT_DOES_NOT_EXIST = &HC0110010
''Private Const HIDP_STATUS_NOT_IMPLEMENTED = &HC0110020
''Private Const HIDP_STATUS_NOT_BUTTON_ARRAY = &HC0110021
''Private Const DIGCF_ALLCLASSES As Long = &H4
''Private Const DIGCF_PROFILE As Long = &H8
''Private Type HIDP_BUTTON_CAPS
''   UsagePage As Integer
''   ReportID As Byte
''   IsAlias As Byte
''   BitField As Integer
''   LinkCollection As Integer
''   LinkUsage As Integer
''   LinkUsagePage As Integer
''   IsRange As Byte
''   IsStringRange As Byte
''   IsDesignatorRange As Byte
''   IsAbsolute As Byte
''   ReportCount As Integer
''   Reserved2 As Integer
''   reserved(0 To 8) As Long
''   UsageOrUsageMin As Integer
''   UsageMax As Integer
''   StringMinOrIndex As Integer
''   StringMax As Integer
''   DesignatorMinOrIndex As Integer
''   DesignatorMax As Integer
''   DataIndexMinOrIndex As Integer
''   DataIndexMax As Integer
''End Type
''Private Type HIDP_LINK_COLLECTION_NODE
''   LinkUsage As Integer
''   LinkUsagePage As Integer
''   Parent As Integer
''   NumberOfChildren As Integer
''   NextSibling As Integer
''   FirstChild As Integer
''   CollectionTypeAndIsAlias As Long
''   UserContext As Long
''End Type
''Private Type HIDP_DATA
''   DataIndex As Integer
''   reserved As Integer
''   RawValue As Long
''End Type
''Private Type HIDP_UNKNOWN_TOKEN
''   Token As Byte
''   reserved(0 To 2) As Byte
''   BitField As Long
''End Type
''Private Type HIDD_CONFIGURATION
''   cookie As Long
''   Size As Long
''   RingBuffersize As Long
''End Type
''Private Type HIDP_EXTENDED_ATTRIBUTES
''   NumGlobalUnknowns As Byte
''   reserved(0 To 2) As Byte
''   GlobalUnknowns As Long
''   Data(0 To 4096) As Long
''End Type
''Private Type HIDP_EXTENDED_ATTRIBUTES_SA
''   NumGlobalUnknowns As Byte
''   reserved(0 To 2) As Byte
''   GlobalUnknowns As Long
''   Data() As Long
''End Type
''Private Type HIDP_BUTTON_ARRAY_DATA
''   ArrayIndex As Integer
''   On As Byte
''End Type
''Private Type HIDP_KEYBOARD_MODIFIER_STATE
''   ul As Long
''End Type
''Private Type SP_CLASSINSTALL_HEADER
''   cbSize As Long
''   InstallFunction As Long
''End Type
''Private Type SP_PROPCHANGE_PARAMS
''   ClassInstallHeader As SP_CLASSINSTALL_HEADER
''   StateChange As Long
''   Scope As Long
''   HwProfile As Long
''End Type
''Private Type USAGE_AND_PAGE
''   Usage As Integer
''   UsagePage As Integer
''End Type
''Private Enum HIDP_KEYBOARD_DIRECTION
''   HidP_Keyboard_Break
''   HidP_Keyboard_Make
''End Enum
''Private Type HDEVINFO 'DeviceInfoSet
''   Value As Long
''End Type
''Private Declare Function apiSysAllocStringByteLen Lib "oleaut32" Alias "SysAllocStringByteLen" (ByVal Ptr As Long, ByVal Length As Long) As String
''Private Declare Function apilstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal dest As String, ByVal source As Long) As String
''Private Declare Function apilstrlen Lib "kernel32" Alias "lstrlenA" (ByVal source As Long) As Long
''Unsure untested
''Private Declare Function apiHidD_GetConfiguration Lib "hid.dll" Alias "HidD_GetConfiguration" (ByVal HidDeviceObject As Long, ByRef Configuration As HIDD_CONFIGURATION, ByVal ConfigurationLength As Long) As Byte
''Private Declare Function apiHidD_SetConfiguration Lib "hid.dll" Alias "HidD_SetConfiguration" (ByVal HidDeviceObject As Long, ByRef Configuration As HIDD_CONFIGURATION, ByVal ConfigurationLength As Long) As Byte
''Private Declare Function apiHidD_GetPhysicalDescriptor Lib "hid.dll" Alias "HidD_GetPhysicalDescriptor" (ByVal HidDeviceObject As Long, ByRef buffer As Any, ByVal BufferLength As Long) As Byte
''Private Declare Function apiHidD_GetManufacturerString Lib "hid.dll" Alias "HidD_GetManufacturerString" (ByVal HidDeviceObject As Long, ByRef buffer As Long, ByVal BufferLength As Long) As Byte
''Private Declare Function apiHidD_GetProductString Lib "hid.dll" Alias "HidD_GetProductString" (ByVal HidDeviceObject As Long, ByRef buffer As Byte, ByVal BufferLength As Long) As Byte
''Private Declare Function apiHidD_GetIndexedString Lib "hid.dll" Alias "HidD_GetIndexedString" (ByVal HidDeviceObject As Long, ByVal StringIndex As Long, ByRef buffer As Any, ByVal BufferLength As Long) As Byte
''Private Declare Function apiHidD_GetSerialNumberString Lib "hid.dll" Alias "HidD_GetSerialNumberString" (ByVal HidDeviceObject As Long, ByRef buffer() As Byte, ByVal BufferLength As Long) As Byte
''Private Declare Function apiHidD_GetMsGenreDescriptor Lib "hid.dll" Alias "HidD_GetMsGenreDescriptor" (ByVal HidDeviceObject As Long, ByRef buffer As Any, ByVal BufferLength As Long) As Byte
''Private Declare Function apiHidP_GetLinkCollectionNodes Lib "hid.dll" Alias "HidP_GetLinkCollectionNodes" (ByRef LinkCollectionNodes As HIDP_LINK_COLLECTION_NODE, ByRef LinkCollectionNodesLength As Long, ByVal PreparsedData As Long) As Long
''Private Declare Function apiHidP_GetSpecificButtonCaps Lib "hid.dll" Alias "HidP_GetSpecificButtonCaps" (ByVal ReportType As Long, ByVal UsagePage As Integer, ByVal LinkCollection As Integer, ByVal Usage As Integer, ByRef ButtonCaps As HIDP_BUTTON_CAPS, ByRef ButtonCapsLength As Integer, ByVal PreparsedData As Long) As Long
''Private Declare Function apiHidP_GetButtonCaps Lib "hid.dll" Alias "HidP_GetButtonCaps" (ByVal ReportType As Long, ByRef ButtonCaps As HIDP_BUTTON_CAPS, ByRef ButtonCapsLength As Integer, ByVal PreparsedData As Long) As Long
''Private Declare Function apiHidP_GetSpecificValueCaps Lib "hid.dll" Alias "HidP_GetSpecificValueCaps" (ByVal ReportType As Long, ByVal UsagePage As Integer, ByVal LinkCollection As Integer, ByVal Usage As Integer, ByRef valuecaps As HIDP_VALUE_CAPS, ByRef ValueCapsLength As Integer, ByVal PreparsedData As Long) As Long
''Private Declare Function apiHidP_GetExtendedAttributes Lib "hid.dll" Alias "HidP_GetExtendedAttributes" (ByVal ReportType As Long, ByVal DataIndex As Integer, ByVal PreparsedData As Long, ByRef Attributes As Any, ByRef LengthAttributes As Long) As Long
''Private Declare Function apiHidP_InitializeReportForID Lib "hid.dll" Alias "HidP_InitializeReportForID" (ByVal ReportType As Long, ByVal ReportID As Byte, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long
''Private Declare Function apiHidP_MaxDataListLength Lib "hid.dll" Alias "HidP_MaxDataListLength" (ByVal ReportType As Long, ByVal PreparsedData As Long) As Long
''Private Declare Function apiHidP_SetData Lib "hid.dll" Alias "HidP_SetData" (ByVal ReportType As Long, ByRef DataList As HIDP_DATA, ByRef DataLength As Long, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long
''Private Declare Function apiHidP_GetData Lib "hid.dll" Alias "HidP_GetData" (ByVal ReportType As Long, ByRef DataList As HIDP_DATA, ByRef DataLength As Long, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long
''Private Declare Function apiHidP_SetUsages Lib "hid.dll" Alias "HidP_SetUsages" (ByVal ReportType As Long, ByVal UsagePage As Integer, ByVal LinkCollection As Integer, ByRef UsageList As Integer, ByRef UsageLength As Long, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long    ' #define HidP_SetButtons(Rty, Up, Lco, ULi, ULe, Ppd, Rep, Rle) \' HidP_SetUsages(Rty, Up, Lco, ULi, ULe, Ppd, Rep, Rle)
''Private Declare Function apiHidP_SetButtons Lib "hid.dll" Alias "HidP_SetUsages" (ByVal ReportType As Long, ByVal UsagePage As Integer, ByVal LinkCollection As Integer, ByRef UsageList As Integer, ByRef UsageLength As Long, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long
''Private Declare Function apiHidP_UnsetUsages Lib "hid.dll" Alias "HidP_UnsetUsages" (ByVal ReportType As Long, ByVal UsagePage As Integer, ByVal LinkCollection As Integer, ByRef UsageList As Integer, ByRef UsageLength As Long, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long    ' #define HidP_UnsetButtons(Rty, Up, Lco, ULi, ULe, Ppd, Rep, Rle) \' HidP_UnsetUsages(Rty, Up, Lco, ULi, ULe, Ppd, Rep, Rle)
''Private Declare Function apiHidP_UnsetButtons Lib "hid.dll" Alias "HidP_UnsetUsages" (ByVal ReportType As Long, ByVal UsagePage As Integer, ByVal LinkCollection As Integer, ByRef UsageList As Integer, ByRef UsageLength As Long, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long
''Private Declare Function apiHidP_GetUsages Lib "hid.dll" Alias "HidP_GetUsages" (ByVal ReportType As Long, ByVal UsagePage As Integer, ByVal LinkCollection As Integer, ByRef UsageList As Integer, ByRef UsageLength As Long, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long    '' #define HidP_GetButtons(Rty, UPa, LCo, ULi, ULe, Ppd, Rep, RLe) \' HidP_GetUsages(Rty, UPa, LCo, ULi, ULe, Ppd, Rep, RLe)
''Private Declare Function apiHidP_GetButtons Lib "hid.dll" Alias "HidP_GetUsages" (ByVal ReportType As Long, ByVal UsagePage As Integer, ByVal LinkCollection As Integer, ByRef UsageList As Integer, ByRef UsageLength As Long, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long
''Private Declare Function apiHidP_GetUsagesEx Lib "hid.dll" Alias "HidP_GetUsagesEx" (ByVal ReportType As Long, ByVal LinkCollection As Integer, ByRef ButtonList As USAGE_AND_PAGE, ByRef UsageLength As Long, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long    '' #define HidP_GetButtonsEx(Rty, LCo, BLi, ULe, Ppd, Rep, RLe)  \' HidP_GetUsagesEx(Rty, LCo, BLi, ULe, Ppd, Rep, RLe)
''Private Declare Function apiHidP_GetButtonsEx Lib "hid.dll" Alias "HidP_GetUsagesEx" (ByVal ReportType As Long, ByVal LinkCollection As Integer, ByRef ButtonList As USAGE_AND_PAGE, ByRef UsageLength As Long, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long
''Private Declare Function apiHidP_MaxUsageListLength Lib "hid.dll" Alias "HidP_MaxUsageListLength" (ByVal ReportType As Long, ByVal UsagePage As Integer, ByVal PreparsedData As Long) As Long
''Private Declare Function apiHidP_SetUsageValue Lib "hid.dll" Alias "HidP_SetUsageValue" (ByVal ReportType As Long, ByVal UsagePage As Integer, ByVal LinkCollection As Integer, ByVal Usage As Integer, ByVal UsageValue As Long, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long
''Private Declare Function apiHidP_SetScaledUsageValue Lib "hid.dll" Alias "HidP_SetScaledUsageValue" (ByVal ReportType As Long, ByVal UsagePage As Integer, ByVal LinkCollection As Integer, ByVal Usage As Integer, ByVal UsageValue As Long, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long
''Private Declare Function apiHidP_SetUsageValueArray Lib "hid.dll" Alias "HidP_SetUsageValueArray" (ByVal ReportType As Long, ByVal UsagePage As Integer, ByVal LinkCollection As Integer, ByVal Usage As Integer, ByRef UsageValue As Any, ByVal UsageValueByteLength As Integer, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long
''Private Declare Function apiHidP_GetUsageValue Lib "hid.dll" Alias "HidP_GetUsageValue" (ByVal ReportType As Long, ByVal UsagePage As Integer, ByVal LinkCollection As Integer, ByVal Usage As Integer, ByRef UsageValue As Long, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long
''Private Declare Function apiHidP_GetScaledUsageValue Lib "hid.dll" Alias "HidP_GetScaledUsageValue" (ByVal ReportType As Long, ByVal UsagePage As Integer, ByVal LinkCollection As Integer, ByVal Usage As Integer, ByRef UsageValue As Long, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long
''Private Declare Function apiHidP_GetUsageValueArray Lib "hid.dll" Alias "HidP_GetUsageValueArray" (ByVal ReportType As Long, ByVal UsagePage As Integer, ByVal LinkCollection As Integer, ByVal Usage As Integer, ByRef UsageValue As Any, ByVal UsageValueByteLength As Integer, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long
''Private Declare Function apiHidP_UsageListDifference Lib "hid.dll" Alias "HidP_UsageListDifference" (PreviousUsageList As Integer, ByRef CurrentUsageList As Integer, ByRef BreakUsageList As Integer, ByRef MakeUsageList As Integer, ByVal UsageListLength As Long) As Long
''Private Declare Function apiHidP_UsageAndPageListDifference Lib "hid.dll" Alias "HidP_UsageAndPageListDifference" (PreviousUsageList As USAGE_AND_PAGE, ByRef CurrentUsageList As USAGE_AND_PAGE, ByRef BreakUsageList As USAGE_AND_PAGE, ByRef MakeUsageList As USAGE_AND_PAGE, ByVal UsageListLength As Long) As Long
''Private Declare Function apiHidP_TranslateUsageAndPagesToI8042ScanCodes Lib "hid.dll" Alias "HidP_TranslateUsageAndPagesToI8042ScanCodes" (ByRef ChangedUsageList As USAGE_AND_PAGE, ByVal UsageListLength As Long, ByVal KeyAction As HIDP_KEYBOARD_DIRECTION, ByRef ModifierState As HIDP_KEYBOARD_MODIFIER_STATE, ByVal InsertCodesProcedure As Long, Optional ByVal InsertCodesContext As Long) As Long
''Private Declare Function apiHidP_TranslateUsagesToI8042ScanCodes Lib "hid.dll" Alias "HidP_TranslateUsagesToI8042ScanCodes" (ByRef ChangedUsageList As Integer, ByVal UsageListLength As Long, ByVal KeyAction As HIDP_KEYBOARD_DIRECTION, ByRef ModifierState As HIDP_KEYBOARD_MODIFIER_STATE, ByVal InsertCodesProcedure As Long, Optional ByVal InsertCodesContext As Long) As Long
''Private Declare Function apiHidP_GetButtonArray Lib "hid.dll" Alias "HidP_GetButtonArray" (ByVal ReportType As Long, ByVal UsagePage As Integer, ByVal LinkCollection As Integer, ByVal Usage As Integer, ByRef ButtonData As HIDP_BUTTON_ARRAY_DATA, ByRef ButtonDataLength As Integer, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long
''Private Declare Function apiHidP_SetButtonArray Lib "hid.dll" Alias "HidP_SetButtonArray" (ByVal ReportType As Long, ByVal UsagePage As Integer, ByVal LinkCollection As Integer, ByVal Usage As Integer, ByRef ButtonData As HIDP_BUTTON_ARRAY_DATA, ByRef ButtonDataLength As Integer, ByVal PreparsedData As Long, ByRef Report As Any, ByVal ReportLength As Long) As Long
''
''Private Declare Function apiSetupDiCreateDeviceInfoList Lib "setupapi.dll" Alias "SetupDiCreateDeviceInfoList" (ByRef ClassGuid As GUID, ByVal hwndParent As Long) As Long
''Private Declare Function apiSetupDiSetClassInstallParamsA Lib "setupapi.dll" Alias "SetupDiSetClassInstallParamsA" (ByVal DeviceInfoSet As Long, ByRef DeviceInfoData As SP_DEVINFO_DATA, ByRef ClassInstallParams As SP_CLASSINSTALL_HEADER, ByVal ClassInstallParamsSize As Long) As Long
''Private Declare Function apiSetupDiChangeState Lib "setupapi.dll" Alias "SetupDiChangeState" (ByVal DeviceInfoSet As Long, ByRef DeviceInfoData As SP_DEVINFO_DATA) As Long
''Private Declare Function apiSetupDiEnumDeviceInfo Lib "setupapi.dll" Alias "SetupDiEnumDeviceInfo" (ByVal DeviceInfoSet As Long, ByVal MemberIndex As Long, ByRef DeviceInfoData As SP_DEVINFO_DATA) As Long
''Private Declare Function apiSetupDiGetDeviceRegistryPropertyA Lib "setupapi.dll" Alias "SetupDiGetDeviceRegistryPropertyA" (ByVal DeviceInfoSet As Long, ByRef DeviceInfoData As SP_DEVINFO_DATA, ByVal Property As Long, ByRef PropertyRegDataType As Long, ByVal PropertyBuffer As Long, ByVal PropertyBufferSize As Long, ByRef RequiredSize As Long) As Long
''Private Declare Function apiSetupDiGetDeviceInstanceId Lib "setupapi.dll" Alias "SetupDiGetDeviceInstanceIdA" (ByVal DeviceInfoSet As Long, ByRef DeviceInfoData As SP_DEVINFO_DATA, ByVal DeviceInstanceId As String, ByVal DeviceInstanceIdSize As Long, ByRef RequiredSize As Long) As Long
''Private Declare Function apiSetupDiGetClassDescription Lib "setupapi.dll" Alias "SetupDiGetClassDescriptionA" (ByRef ClassGuid As GUID, ByVal ClassDescription As String, ByVal ClassDescriptionSize As Long, ByRef RequiredSize As Long) As Long
''Private Declare Function apiSetupDiEnumDeviceInterfaces Lib "setupapi.dll" Alias "SetupDiEnumDeviceInterfaces" (ByVal DeviceInfoSet As Long, ByRef DeviceInfoData As Any, ByRef InterfaceClassGuid As GUID, ByVal MemberIndex As Long, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA) As Long
''Private Declare Function apiSetupDiGetDeviceInterfaceDetail Lib "setupapi.dll" Alias "SetupDiGetDeviceInterfaceDetailA" (ByVal DeviceInfoSet As Long, ByRef DeviceInterfaceData As Any, ByRef DeviceInterfaceDetailData As Any, ByVal DeviceInterfaceDetailDataSize As Long, ByRef RequiredSize As Long, ByRef DeviceInfoData As Any) As Long
''Private Declare Function apiSetupDiGetDeviceRegistryPropertyW Lib "setupapi.dll" Alias "SetupDiGetDeviceRegistryPropertyW" (ByVal DeviceInfoSet As Long, ByRef DeviceInfoData As SP_DEVINFO_DATA, ByVal Property As Long, ByRef PropertyRegDataType As Long, ByVal PropertyBuffer As Long, ByVal PropertyBufferSize As Long, ByRef RequiredSize As Long) As Long
''Private Declare Function apiSetupDiSetClassInstallParamsW Lib "setupapi.dll" Alias "SetupDiSetClassInstallParamsW" (ByVal DeviceInfoSet As Long, ByRef DeviceInfoData As SP_DEVINFO_DATA, ByRef ClassInstallParams As SP_CLASSINSTALL_HEADER, ByVal ClassInstallParamsSize As Long) As Long
''Private Function GetErrorString(ByVal LastError As Long) As String
''   Dim Bytes As Long
''   Dim ErrorString As String
''   ErrorString = VBA.String$(129, 0)
''   Bytes = apiFormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, LastError, 0, ErrorString$, 128, 0)
''   If Bytes > 2 Then 'Subtract two characters from the message to strip the CR and LF.
''      GetErrorString = VBA.Left$(ErrorString, Bytes - 2)
''   End If
''End Function
''Private Function GetFeatureReport(ByVal hidHandle As Long, ByRef inFeatureReportBuffer() As Byte) As Boolean
''   Dim success As Byte
''   Dim nLen As Long
''   nLen = UBound(inFeatureReportBuffer) - LBound(inFeatureReportBuffer) + 1
''   success = apiHidD_GetFeature(hidHandle, inFeatureReportBuffer(0), nLen)
''   GetFeatureReport = CBool(success)
''End Function
''
''Private Function SetNumberOfInputBuffers(ByVal HidDeviceObject As Long, ByVal NumberBuffers As Long) As Boolean
''   Dim success As Byte 'Sets the number of Input reports the host can store.  If the buffer is full and another report arrives, the host drops the oldest report.
''   apiHidD_SetNumInputBuffers HidDeviceObject, NumberBuffers
''   SetNumberOfInputBuffers = CBool(success)
''End Function
''
''Private Function GetNumberOfInputBuffers(ByVal HidDeviceObject As Long, ByRef numberOfInputBuffers As Long) As Boolean
''   Dim success As Byte 'Retrieves the number of Input reports the host can store. If the buffer is full and another report arrives, the host drops the oldest report.
''   success = apiHidD_GetNumInputBuffers(HidDeviceObject, numberOfInputBuffers)
''   GetNumberOfInputBuffers = CBool(success)
''End Function
''
''Friend Function GetHidUsage(ByRef MyCapabilities As HIDP_CAPS) As String
''   Dim usage As Long 'Creates a 32-bit Usage from the Usage Page and Usage ID. Determines whether the Usage is a system mouse or keyboard.  Can be modified to detect other Usages.
''   Dim usageDescription As String
''   usage = MyCapabilities.UsagePage * 256 + MyCapabilities.usage ' Create32-bit Usage from Usage Page and Usage ID.
''   If usage = &H102 Then usageDescription = "mouse"
''   If usage = &H106 Then usageDescription = "keyboard"
''   GetHidUsage = usageDescription
''End Function
''
''
''Private Function SendOutputReportViaControlTransfer(ByVal hidHandle As Long, ByRef outputReportBuffer() As Byte) As Boolean 'ByRef outputReportBuffer As Byte ?
''   Dim success As Byte 'Writes an Output report to the device using a control transfer.
''   Dim nLen As Long
''   nLen = UBound(outputReportBuffer) - LBound(outputReportBuffer) + 1
''   success = apiHidD_SetOutputReport(hidHandle, outputReportBuffer(0), nLen)
''   SendOutputReportViaControlTransfer = CBool(success)
''End Function
''
''Private Function SendFeatureReport(ByVal hidHandle As Long, ByRef outFeatureReportBuffer() As Byte) As Boolean 'ByRef outFeatureReportBuffer As Byte ?
''   Dim success As Byte 'Writes a Feature report to the device.
''   Dim nLen As Long
''   nLen = UBound(outFeatureReportBuffer) - LBound(outFeatureReportBuffer) + 1
''   success = apiHidD_SetFeature(hidHandle, outFeatureReportBuffer(0), nLen)
''   SendFeatureReport = CBool(success)
''End Function
''
''
''Friend Function GetInputReportViaControlTransfer(ByVal hidHandle As Long, ByRef inputReportBuffer() As Byte) As Boolean
''   Dim success As Byte  'Reads an Input report from the device using a control transfer.
''   Dim nLen As Long
''   nLen = UBound(inputReportBuffer) - LBound(inputReportBuffer) + 1
''   success = apiHidD_GetInputReport(hidHandle, inputReportBuffer(0), nLen)
''   GetInputReportViaControlTransfer = CBool(success)
''End Function
''
''
''Public Function IsByteArraysEqual(barray1() As Byte, barray2() As Byte) As Boolean
''   Dim i As Long
''   IsByteArraysEqual = False
''   If UBound(barray1) <> UBound(barray2) Then Exit Function
''   For i = LBound(barray1) To UBound(barray1)
''      If barray1(i) <> barray2(i) Then Exit Function
''   Next
''   IsByteArraysEqual = True
''End Function
'''Dimensioning byte arrays in VB6 is counter-intuitive. An array of n bytes is dimensioned as Dim abArray(n - 1)
'''An n-byte array [abArray(0), abArray(1), ..., abArray(n-1)].
'''Use indexes For i = 0 To n-1, but dimension the size of the array different.
'''If you use Dim abArray(n) you have an extra byte at the end. A more intuitive way is to use the full, formal expression
'''Dim abArray(0 To n - 1)
''
''
''         If flag10 And VK_PAD_MINUS Then
''            If JOYCON1_MINUS = False Then
''               JOYCON1_MINUS = True
''               RaiseEvent ClickButton(1, True, "Press Minus")
''            End If
''         Else
''            If JOYCON1_MINUS = True Then
''               JOYCON1_MINUS = False
''               RaiseEvent ClickButton(1, False, "Release Minus")
''            End If
''         End If
''         If flag10 And VK_PAD_LTHUMB_PRESS Then
''            bState = bState & " + Left-Stick "
''            If JOYCON1_LSTICK = False Then
''               JOYCON1_LSTICK = True
''               RaiseEvent ClickButton(1, True, "Press Left Stick")
''            End If
''         Else
''            If JOYCON1_LSTICK = True Then
''               JOYCON1_LSTICK = False
''               RaiseEvent ClickButton(1, False, "Release Left Stick")
''            End If
''         End If
''         If flag9 And VK_PAD_BACK Then
''            bState = bState & " + View "
''            If JOYCON1_VIEW = False Then
''               JOYCON1_VIEW = True
''               RaiseEvent ClickButton(1, True, "Press View")
''            End If
''         Else
''            If JOYCON1_VIEW = True Then
''               JOYCON1_VIEW = False
''               RaiseEvent ClickButton(1, False, "Release View")
''            End If
''         End If
''         If flag11 And VK_PAD_LSHOULDER Then
''            bState = bState & " + Left-Shoulder "
''            If JOYCON1_LSHOULDER = False Then
''               JOYCON1_LSHOULDER = True
''               RaiseEvent ClickButton(1, True, "Press Left Shoulder")
''            End If
''         Else
''            If JOYCON1_LSHOULDER = True Then
''               JOYCON1_LSHOULDER = False
''               RaiseEvent ClickButton(1, False, "Release Left Shoulder")
''            End If
''         End If
''         If flag11 And VK_PAD_LTRIGGER Then
''            bState = bState & " + Left-Trigger "
''            If JOYCON1_LTRIGGER = False Then
''               JOYCON1_LTRIGGER = True
''               RaiseEvent ClickButton(1, True, "Press Left Trigger")
''            End If
''         Else
''            If JOYCON1_LTRIGGER = True Then
''               JOYCON1_LTRIGGER = False
''               RaiseEvent ClickButton(1, False, "Release Left Trigger")
''            End If
''         End If
''         If flag12 And VK_PAD_DPAD_LEFT Then
''            bState = bState & " + Dpad-Left "
''            If JOYCON1_DPAD_LEFT = False Then
''               JOYCON1_DPAD_LEFT = True
''               RaiseEvent ClickButton(1, True, "Press Dpad Left")
''            End If
''         Else
''            If JOYCON1_DPAD_LEFT = True Then
''               JOYCON1_DPAD_LEFT = False
''               RaiseEvent ClickButton(1, False, "Release Dpad Left")
''            End If
''         End If
''         If flag12 And VK_PAD_DPAD_DOWN Then
''            bState = bState & " + Dpad-Down "
''            If JOYCON1_DPAD_DOWN = False Then
''               JOYCON1_DPAD_DOWN = True
''               RaiseEvent ClickButton(1, True, "Press Dpad Down")
''            End If
''         Else
''            If JOYCON1_DPAD_DOWN = True Then
''               JOYCON1_DPAD_DOWN = False
''               RaiseEvent ClickButton(1, False, "Release Dpad Down")
''            End If
''         End If
''         If flag12 And VK_PAD_DPAD_UP Then
''            bState = bState & " + Dpad-Up "
''            If JOYCON1_DPAD_UP = False Then
''               JOYCON1_DPAD_UP = True
''               RaiseEvent ClickButton(1, True, "Press Dpad Up")
''            End If
''         Else
''            If JOYCON1_DPAD_UP = True Then
''               JOYCON1_DPAD_UP = False
''               RaiseEvent ClickButton(1, False, "Release Dpad Up")
''            End If
''         End If
''         If flag12 And VK_PAD_DPAD_RIGHT Then
''            bState = bState & " + Dpad-Right "
''            If JOYCON1_DPAD_RIGHT = False Then
''               JOYCON1_DPAD_RIGHT = True
''               RaiseEvent ClickButton(1, True, "Press Dpad Right")
''            End If
''         Else
''            If JOYCON1_DPAD_RIGHT = True Then
''               JOYCON1_DPAD_RIGHT = False
''               RaiseEvent ClickButton(1, False, "Release Dpad Right")
''            End If
''         End If
''         '
''         'S buttons
''         If flag11 And VK_PAD_SL1 Then
''            bState = bState & "SL1 "
''            If JOYCON1_SL = False Then
''               JOYCON1_SL = True
''               RaiseEvent ClickButton(1, True, "Press Left SL")
''            End If
''         Else
''            If JOYCON1_SL = True Then
''               JOYCON1_SL = False
''               RaiseEvent ClickButton(1, False, "Release Left SL")
''            End If
''         End If
''         If flag11 And VK_PAD_SR1 Then
''            bState = bState & "SR1 "
''            If JOYCON1_SR = False Then
''               JOYCON1_SR = True
''               RaiseEvent ClickButton(1, True, "Press Left SR")
''            End If
''         Else
''            If JOYCON1_SR = True Then
''               JOYCON1_SR = False
''               RaiseEvent ClickButton(1, False, "Release Left SR")
''            End If
''         End If
''         If flag10 And VK_PAD_PLUS Then
''            If JOYCON2_PLUS = False Then
''               JOYCON2_PLUS = True
''               RaiseEvent ClickButton(2, True, "Press Plus")
''            End If
''         Else
''            If JOYCON2_PLUS = True Then
''               JOYCON2_PLUS = False
''               RaiseEvent ClickButton(2, False, "Release Plus")
''            End If
''         End If
''         If flag10 And VK_PAD_RTHUMB_PRESS Then
''            If JOYCON2_RSTICK = False Then
''               JOYCON2_RSTICK = True
''               RaiseEvent ClickButton(2, True, "Press Right Stick")
''            End If
''         Else
''            If JOYCON2_RSTICK = True Then
''               JOYCON2_RSTICK = False
''               RaiseEvent ClickButton(2, False, "Release Right Stick")
''            End If
''         End If
''         If flag9 And VK_PAD_START Then
''            If JOYCON2_MENU = False Then
''               JOYCON2_MENU = True
''               RaiseEvent ClickButton(2, True, "Press Menu")
''            End If
''         Else
''            If JOYCON2_MENU = True Then
''               JOYCON2_MENU = False
''               RaiseEvent ClickButton(2, False, "Release Menu")
''            End If
''         End If
''         If flag7 And VK_PAD_RSHOULDER Then
''            If JOYCON2_RSHOULDER = False Then
''               JOYCON2_RSHOULDER = True
''               RaiseEvent ClickButton(2, True, "Press Right Shoulder")
''            End If
''         Else
''            If JOYCON2_RSHOULDER = True Then
''               JOYCON2_RSHOULDER = False
''               RaiseEvent ClickButton(2, False, "Release Right Shoulder")
''            End If
''         End If
''         If flag7 And VK_PAD_RTRIGGER Then
''            If JOYCON2_RTRIGGER = False Then
''               JOYCON2_RTRIGGER = True
''               RaiseEvent ClickButton(2, True, "Press Right Trigger")
''            End If
''         Else
''            If JOYCON2_RTRIGGER = True Then
''               JOYCON2_RTRIGGER = False
''               RaiseEvent ClickButton(2, False, "Release Right Trigger")
''            End If
''         End If
''         If flag8 And VK_PAD_A Then '  Action buttons
''            If JOYCON2_A = False Then
''               JOYCON2_A = True
''               RaiseEvent ClickButton(2, True, "Press A")
''            End If
''         Else
''            If JOYCON2_A = True Then
''               JOYCON2_A = False
''               RaiseEvent ClickButton(2, False, "Release A")
''            End If
''         End If
''         If flag8 And VK_PAD_B Then
''            If JOYCON2_B = False Then
''               JOYCON2_B = True
''               RaiseEvent ClickButton(2, True, "Press B")
''            End If
''         Else
''            If JOYCON2_B = True Then
''               JOYCON2_B = False
''               RaiseEvent ClickButton(2, False, "Release B")
''            End If
''         End If
''         If flag8 And VK_PAD_X Then
''            If JOYCON2_X = False Then
''               JOYCON2_X = True
''               RaiseEvent ClickButton(2, True, "Press X")
''            End If
''         Else
''            If JOYCON2_X = True Then
''               JOYCON2_X = False
''               RaiseEvent ClickButton(2, False, "Release X")
''            End If
''         End If
''         If flag8 And VK_PAD_Y Then
''            If JOYCON2_Y = False Then
''               JOYCON2_Y = True
''               RaiseEvent ClickButton(2, True, "Press Y")
''            End If
''         Else
''            If JOYCON2_Y = True Then
''               JOYCON2_Y = False
''               RaiseEvent ClickButton(2, False, "Release Y")
''            End If
''         End If
''         If flag7 And VK_PAD_SL2 Then 'S buttons
''            If JOYCON2_SL = False Then
''               JOYCON2_SL = True
''               RaiseEvent ClickButton(2, True, "Press Right SL")
''            End If
''         Else
''            If JOYCON2_SL = True Then
''               JOYCON2_SL = False
''               RaiseEvent ClickButton(2, False, "Release Right SL")
''            End If
''         End If
''         If flag7 And VK_PAD_SR2 Then
''            If JOYCON2_SR = False Then
''               JOYCON2_SR = True
''               RaiseEvent ClickButton(2, True, "Press Right SR")
''            End If
''         Else
''            If JOYCON2_SR = True Then
''               JOYCON2_SR = False
''               RaiseEvent ClickButton(2, False, "Release Right SR")
''            End If
''         End If
