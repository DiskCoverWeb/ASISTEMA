Attribute VB_Name = "SubWinZip"
Option Explicit
'Variable to communicate to the FrmDialogo form

Public ExtractDialogCanceled As Boolean

'Constants to determine characteristics of Zip Open Dialog
Public Const OpenZip = 0
Public Const NewZip = 1
Public Const TestZip = 2
Public Const FixZip = 3
Public Const DeleteZip = 4
Public Const SelectBin = 5
Public Const SelectAll = 6
Public Const OpenFile = 7
Public Const SaveFile = 8

'Constants to determine command executed by ExecuteSelFilesCmd
Public Const SF_Delete = 0
Public Const SF_Extract = 1

Public Type ZIPUSERFUNCTIONS
   DLLPrnt As Long
   DLLPassword As Long
   DLLComment As Long
   DLLService As Long
End Type

Public Type UNZIPUSERFUNCTION
   UNZIPPrntFunction As Long
   UNZIPSndFunction As Long
   UNZIPReplaceFunction  As Long
   UNZIPPassword As Long
   UNZIPMessage  As Long
   UNZIPService  As Long
   TotalSizeComp As Long
   TotalSize As Long
   CompFactor As Long
   NumFiles As Long
   Comment As Integer
End Type

Public Type ZPOPT
   fSuffix As Long
   fEncrypt As Long
   fSystem As Long
   fVolume As Long
   fExtra As Long
   fNoDirEntries As Long
   fExcludeDate As Long
   fIncludeDate As Long
   fVerbose As Long
   fQuiet As Long
   fCRLF_LF As Long
   fLF_CRLF As Long
   fJunkDir As Long
   fRecurse As Long
   fGrow As Long
   fForce As Long
   fMove As Long
   fDeleteEntries As Long
   fUpdate As Long
   fFreshen As Long
   fJunkSFX As Long
   fLatestTime As Long
   fComment As Long
   fOffsets As Long
   FPrivilege As Long
   fEncryption As Long
   fRepair As Long
   flevel As Byte
   date As String
   szRootDir As String
End Type

Private Type UNZIPOPTIONS
    ExtractOnlyNewer  As Long
    SpaceToUnderScore As Long
    PromptToOverwrite As Long
    fQuiet As Long
    ncflag As Long
    ntflag As Long
    nvflag As Long
    nUflag As Long
    nzflag As Long
    ndflag As Long
    noflag As Long
    naflag As Long
    nZIflag As Long
    C_flag As Long
    FPrivilege As Long
    Zip As String
    ExtractDir As String
End Type

Public Type ZIPnames
    S(0 To 254) As String
End Type

Public Type CBChar
    ch(4096) As Byte
End Type

Global NombreArchivoZip As String
Global NombresFicherosZip As ZIPnames

'===================== Funciones del ZIP32.DLL =======================================================
Public Declare Function ZpInit Lib "zip32.dll" (ByRef Zipfun As ZIPUSERFUNCTIONS) As Long
Public Declare Function ZpSetOptions Lib "zip32.dll" (ByRef Opts As ZPOPT) As Long
Public Declare Function ZpArchive Lib "zip32.dll" (ByVal argc As Long, ByVal funame As String, ByRef argv As ZIPnames) As Long

'===================== Funciones del UNZIP32.DLL =======================================================
Public Declare Function Wiz_SingleEntryUnzip Lib "unzip32.dll" (ByVal ifnc As Long, ByRef ifnv As ZIPnames, ByVal xfnc As Long, ByRef xfnv As ZIPnames, dcll As UNZIPOPTIONS, Userf As UNZIPUSERFUNCTION) As Long

Function FuncionParaProcesarPassword(ByRef B1 As Byte, L As Long, ByRef B2 As Byte, ByRef B3 As Byte) As Long
    FuncionParaProcesarPassword = 0
End Function

Function FuncionParaProcesarServicios(ByRef fname As CBChar, ByVal X As Long) As Long
    FuncionParaProcesarServicios = 0
End Function

Function FuncionParaProcesarMensajes(ByRef fname As CBChar, ByVal X As Long) As Long
    FuncionParaProcesarMensajes = 0
End Function

Function FuncionParaProcesarComentarios(Comentario As CBChar) As CBChar
    Comentario.ch(0) = vbNullString
    FuncionParaProcesarComentarios = Comentario
End Function

Public Function DevolverDireccionMemoria(Direccion As Long) As Long
On Error GoTo err_DevolverDireccionMemoria

    DevolverDireccionMemoria = Direccion

Exit Function
err_DevolverDireccionMemoria:
    MsgBox "DevolverDireccionMemoria: " + Err.Description, vbExclamation
    Err.Clear
End Function


Private Function UNFuncionParaProcesarMensajes(ByRef fname As CBChar, ByVal X As Long) As Long
On Error GoTo err_UNFuncionParaProcesarMensajes

    UNFuncionParaProcesarMensajes = 0

Exit Function
err_UNFuncionParaProcesarMensajes:
    MsgBox "UNFuncionParaProcesarMensajes: " + Err.Description, vbExclamation
    Err.Clear
End Function

Private Function UNFuncionReplaceOptions(ByRef p As CBChar, ByVal L As Long, ByRef m As CBChar, ByRef Name As CBChar) As Integer
On Error GoTo err_UNFuncionReplaceOptions

    UNFuncionReplaceOptions = 0
    'UNFuncionParaProcesarPassword = 0

Exit Function
err_UNFuncionReplaceOptions:
    MsgBox "UNFuncionParaProcesarPassword: " + Err.Description, vbExclamation
    Err.Clear
End Function

Public Sub UnZip(Zip As String, ExtractDir As String)
On Error GoTo err_Unzip

Dim Resultado As Long
Dim intContadorFicheros As Integer

Dim FuncionesUnZip As UNZIPUSERFUNCTION
Dim OpcionesUnZip As UNZIPOPTIONS

Dim NombresFicherosZip As ZIPnames, NombresFicheros2Zip As ZIPnames

NombresFicherosZip.S(0) = vbNullChar
NombresFicheros2Zip.S(0) = vbNullChar
FuncionesUnZip.UNZIPMessage = 0&
FuncionesUnZip.UNZIPPassword = 0&
FuncionesUnZip.UNZIPPrntFunction = DevolverDireccionMemoria(AddressOf UNFuncionParaProcesarMensajes)
FuncionesUnZip.UNZIPReplaceFunction = DevolverDireccionMemoria(AddressOf UNFuncionReplaceOptions)
FuncionesUnZip.UNZIPService = 0&
FuncionesUnZip.UNZIPSndFunction = 0&
OpcionesUnZip.C_flag = 1
OpcionesUnZip.fQuiet = 2
OpcionesUnZip.noflag = 1
OpcionesUnZip.Zip = Zip
OpcionesUnZip.ExtractDir = ExtractDir

Resultado = Wiz_SingleEntryUnzip(0, NombresFicherosZip, 0, NombresFicheros2Zip, OpcionesUnZip, FuncionesUnZip)

Exit Sub
err_Unzip:
    MsgBox "Unzip: " + Err.Description, vbExclamation
    Err.Clear
End Sub

Public Function SelectZipFile(DlgSelectZip As CommonDialog, DialogType As Integer) As String
   With DlgSelectZip
       .Flags = cdlOFNFileMustExist + cdlOFNNoChangeDir + cdlOFNHideReadOnly
       .Filter = "Zip archives (*.zip)|*.zip|Self-extracting Zip archives (*.exe)|*.exe|All files (*.*)|*.*"
       If .Filename = "" Then .Filename = "Seleccione un Archivo"
        Select Case DialogType
          Case OpenZip
              '.FileName = ""
              .DialogTitle = "Open Archive"
              .Action = 1
          Case NewZip
              .Flags = cdlOFNOverwritePrompt + cdlOFNNoChangeDir + cdlOFNHideReadOnly
              .DialogTitle = "New Archive"
              .Action = 2
             ' Pretend we are saving file. A new archive is really created when Adding files.
          Case TestZip
              '.FileName = ""
              .DialogTitle = "Test Archive"
              .Action = 1
          Case FixZip
              '.FileName = ""
              .DialogTitle = "Fix Archive"
              .Action = 1
          Case DeleteZip
              '.FileName = ""
              .DialogTitle = "Delete Archive"
              .Action = 1
          Case SelectBin
              '.FileName = ""
              .Filter = "Self-extractor binary (*.bin)|*.bin|All files (*.*)|*.*"
              .DialogTitle = "Select self-extractor binary"
              .Action = 1
          Case SelectAll
              .Filter = "All files (*.*)|*.*"
              .DialogTitle = "Seleccionar el directorio de respaldo"
              .Action = 2
          Case OpenFile
              .Filter = "All files (*.*)|*.*"
              .DialogTitle = "Open Archive"
              .Action = 1
        End Select
        SelectZipFile = .Filename
   End With
End Function

