Attribute VB_Name = "APIcalls"
'   This module is common to all of the Example programs
'   It declares the data types and system calls required to access the Windows operating system
'
Public Type Security_Attributes: nLength As Long: lpSecurityDescriptor As Long: bInheritHandle As Long: End Type

Public Type Guid: Data(3) As Long: End Type 'A GUID is 16 bytes long
    
Public Type Device_Interface_Data
cbsize As Long: InterfaceClassGuid As Guid: Flags As Long: ReservedPtr As Long: End Type
    
Public Type Device_Interface_Detail: cbsize As Long: DataPath(256) As Byte: End Type

Public Type ConfigurationDataType: cookie As Long: cbsize As Long: RingBuffersize As Long: End Type

Public Type HidD_Attributes
cbsize As Long: VendorID(1) As Byte: ProductID(1) As Byte: VersionNumber(1) As Byte: Pad(10) As Byte: End Type

'   Declare the functions from Dan Appleman's "Programmers Guide to the Win32 API"
Declare Function AddressFor Lib "apigid32.dll" Alias "agGetAddressForObject" (PassedByReference As Any) As Long
Declare Sub CopyBuffer Lib "apigid32.dll" Alias "agCopyData" (ByVal SourcePtr&, ByVal DestPtr&, ByVal ByteCount&)
'
'   Declare the API calls that I am using
Declare Sub HidD_GetHidGuid Lib "HID.dll" (GuidPtr&)

Declare Function SetupDiGetClassDevs Lib "setupapi.dll" Alias "SetupDiGetClassDevsA" _
(GuidPtr&, ByVal EnumPtr&, ByVal HwndParent&, ByVal Flags&) As Long

Declare Function SetupDiDestroyDeviceInfoList Lib "setupapi.dll" (ByVal DeviceInfoSet&) As Boolean

Declare Function SetupDiEnumDeviceInterfaces Lib "setupapi.dll" _
(ByVal Handle&, ByVal InfoPtr&, GuidPtr&, ByVal MemberIndex&, InterfaceDataPtr&) As Boolean

Declare Function SetupDiGetDeviceInterfaceDetail Lib "setupapi.dll" Alias "SetupDiGetDeviceInterfaceDetailA" _
(ByVal Handle&, InterfaceDataPtr&, InterfaceDetailPtr&, ByVal DetailLength&, _
ReturnedLengthPtr&, ByVal DevInfoDataPtr&) As Boolean
 
Declare Function GetLastError Lib "kernel32" () As Long

Declare Function ReadFile Lib "kernel32" _
(ByVal Handle&, ByVal BufferPtr&, ByVal ByteCount&, BytesReturnedPtr&, ByVal OverlappedPtr&) As Long

Declare Function DeviceIoControl Lib "kernel32" _
(ByVal hDevice&, ByVal dwIoControlCode&, lpInBuffer&, ByVal nInBufferSize&, _
 lpOutBuffer&, ByVal nOutBufferSize&, lpBytesReturned&, lpOverlapped&) As Long
 
Declare Function HidD_GetPreparsedData Lib "HID.dll" (ByVal Handle&, ByVal BufferPtr&) As Long
Declare Function HidD_GetAttributes Lib "HID.dll" (ByVal Handle&, BufferPtr&) As Long
Declare Function HidD_GetManufacturerString Lib "HID.dll" (ByVal Handle&, ByVal BufferPtr&, ByVal Length&) As Long
Declare Function HidD_GetProductString Lib "HID.dll" (ByVal Handle&, ByVal BufferPtr&, ByVal Length&) As Long
Declare Function HidD_GetSerialNumberString Lib "HID.dll" (ByVal Handle&, ByVal BufferPtr&, ByVal Length&) As Long
Declare Function HidD_GetIndexedString Lib "HID.dll" (ByVal Handle&, ByVal Index&, ByVal BufferPtr&, ByVal Length&) As Long
Declare Function HidD_GetConfiguration Lib "HID.dll" (ByVal Handle&, ByVal BufferPtr&, ByVal Length&) As Long
Declare Function HidD_SetConfiguration Lib "HID.dll" (ByVal Handle&, ByVal BufferPtr&, ByVal Length&) As Long
Declare Function HidD_GetPhysicalDescriptor Lib "HID.dll" (ByVal Handle&, ByVal BufferPtr&, ByVal Length&) As Long

