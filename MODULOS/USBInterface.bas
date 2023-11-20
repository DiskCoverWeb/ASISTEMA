Attribute VB_Name = "USBdefinitions"
'
'   Declare all of the USB Data Structures
'
'   Note that most of these Data Structures MUST be defined a BYTES.
'   This prevents Visual Basic "helpfully" aligning variables on their natural byte boundaries.
'   Little Endian is assumed. ie If Byte(3)= Long,  Then byte(0) = LSB
'
Public Type UNameType: Length As Long: UnicodeName(256) As Byte: End Type
Public Type UNodeType: ConnectionIndex As Long: Length As Long: UnicodeName(256) As Byte: End Type

Public Type SetupPacket
    RequestType As Byte: Request As Byte: wValueLo As Byte: wValueHi As Byte: wIndex As Integer: wLength As Integer: End Type

Public Type DescriptorRequest
    ConnectionIndex As Long: PacketData As SetupPacket: ConfigurationDescriptor(2000) As Byte: End Type
    
Public Type DeviceDescriptor
     Contents(17) As Byte: End Type
' Defined as a Byte Array to make later data movement simpler
'    Length As Byte:    DescriptorType As Byte:    USBSpec(1) As Byte:    Class As Byte
'    SubClass As Byte:  Protocol As Byte:          MaxEP0Size As Byte:    VendorID(1) As Byte
'    ProductID(1) As Byte:       DeviceRevision(1) As Byte:       ManufacturerStringIndex As Byte
'    ProductStringIndex As Byte: SerialNumberStringIndex As Byte: ConfigurationCount As Byte: End Type

Public Type HubDescriptor
    Length As Byte:       HubType As Byte:    PortCount As Byte:     Characteristics(1) As Byte
    PowerOn2Good As Byte: MaxCurrent As Byte: PowerMask(63) As Byte: End Type

Public Type EndPointDescriptor
    Length As Byte:     DescriptorType As Byte:   EndpointAddress As Byte
    Attributes As Byte: MaxPacketSize(1) As Byte: PollingInterval As Byte: End Type

Public Type NodeInformation
    NodeType As Long: NodeDescriptor As HubDescriptor: HubIsBusPowered As Byte: End Type

Public Type NodeConnectionInformation
    ConnectionIndex As Long: ThisDevice As DeviceDescriptor: CurrentConfiguration As Byte
    LowSpeed As Byte: DeviceIsHub As Byte: DeviceAddress(1) As Byte: NumberOfOpenEndPoints(3) As Byte
    ThisConnectionStatus(3) As Byte: MyEndPoints(29) As EndPointDescriptor: End Type

'   I keep all of the IO Device information I collect in a big table
'   Most USB installations will only fill part of this table
Public Type CollectedDeviceData
    DeviceType As Long: DeviceHandle As Long: ConnectionData As NodeConnectionInformation
    NodeData As NodeInformation: End Type
Public DeviceData(200) As CollectedDeviceData
'
'   All Descriptors are concatenated here once a device is selected
Public DescriptorData(2000) As Byte
'
'   I need to send Requests to USB devices
Public PCHostRequest As DescriptorRequest
'
Public ConnectionStatus(6) As String
'
'   Declare my support sub-routines
Public Function DataIndex()
' All writes to the DeviceData table are done to entry DataIndex
' Need to keep DeviceData and IODevice_Display in sync
DataIndex = Collect_Data.Device_Display.ListCount
End Function

Public Function OpenConnection(Name$)
Dim SA As Security_Attributes
Handle& = CreateFile("\\.\" & Name$, &HC0000000, 3, SA, 3, 0, 0)
If Handle& = 0 Then ErrorExit ("Could not open a connection to " & Name$)
OpenConnection = Handle&
End Function

Public Sub GetNodeInformation(Handle&)
'   Get the node information
Dim BytesReturned&, Status&
Status& = DeviceIoControl(Handle&, &H220408, DeviceData(DataIndex).NodeData.NodeType, 256, DeviceData(DataIndex).NodeData.NodeType, 256, BytesReturned&, 0)
If Status& = 0 Then ErrorExit ("Could not get node information")
If BytesReturned& > 256 Then ErrorExit ("DeviceIOControl returned >256 bytes of data")
End Sub

Public Sub GetNodeConnectionData(Handle&, PortIndex&)
Dim BytesReturned&, Status&
DeviceData(DataIndex).ConnectionData.ConnectionIndex = PortIndex&
Status& = DeviceIoControl(Handle&, &H22040C, DeviceData(DataIndex).ConnectionData.ConnectionIndex, 256, DeviceData(DataIndex).ConnectionData.ConnectionIndex, 256, BytesReturned&, 0)
If Status& = 0 Then ErrorExit ("Could not get Node Connection Data")
If BytesReturned& > 256 Then ErrorExit ("DeviceIOControl returned >256 bytes of data")
End Sub

Function GetNameOf$(DeviceName$, DeviceHandle&, API_ID&)
Dim NameBuffer As UNameType
'
'   First need to get the length of the name string
Status& = DeviceIoControl(DeviceHandle&, API_ID&, 0, 0, NameBuffer.Length, 260, BytesReturned&, 0)
If Status& = 0 Then ErrorExit ("Could not get LENGTH of " & DeviceName$ & " Name")
If NameBuffer.Length > 256 Then ErrorExit (Name$ & " Name > 256 Characters")
'
'   . . . and then the string. It will be returned in UNICODE format
Status& = DeviceIoControl(DeviceHandle&, API_ID&, NameBuffer.Length, NameBuffer.Length, NameBuffer.Length, NameBuffer.Length, BytesReturned&, 0)
If Status& = 0 Then ErrorExit ("Could not get TEXT of " & DeviceName$ & " Name")
temp$ = "": I = 0   'A simple unicode to basic string conversion
Do While NameBuffer.UnicodeName(I) <> 0: temp$ = temp$ & Chr(NameBuffer.UnicodeName(I)): I = I + 2: Loop
GetNameOf$ = temp$ 'StrConv(NameBuffer.Length, vbFromUnicode)
End Function

Function GetExternalHubName$(ConnectionIndex&, DeviceHandle&)
Dim NameBuffer As UNodeType
'
'   First need to get the length of the name string
NameBuffer.ConnectionIndex = ConnectionIndex
Status& = DeviceIoControl(DeviceHandle&, &H220414, NameBuffer.ConnectionIndex, 260, NameBuffer.ConnectionIndex, 260, BytesReturned&, CNull)
If Status& = 0 Then ErrorExit ("Could not get LENGTH of External Hub Name")
If NameBuffer.Length > 256 Then ErrorExit ("External Hub Name > 256 Characters")
'
'   . . . and then the string. It will be returned in UNICODE format
NameBuffer.ConnectionIndex = ConnectionIndex
Status& = DeviceIoControl(DeviceHandle&, &H220414, NameBuffer.ConnectionIndex, NameBuffer.Length, NameBuffer.ConnectionIndex, NameBuffer.Length, BytesReturned&, 0)
If Status& = 0 Then ErrorExit ("Could not get TEXT of External Hub Name")
temp$ = "": I = 0
Do While NameBuffer.UnicodeName(I) <> 0: temp$ = temp$ & Chr(NameBuffer.UnicodeName(I)): I = I + 2: Loop
GetExternalHubName$ = temp$
End Function

