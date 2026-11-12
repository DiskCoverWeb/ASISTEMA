Attribute VB_Name = "modDialogos"
Option Explicit

' =========================================================================
' M√≥dulo de selecci√≥n de archivos mediante API nativa de Windows
' Evita depender de controles externos como Microsoft Common Dialog Control (COMDLG32.OCX)
' =========================================================================

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_EXPLORER = &H80000

Public Function AbrirDialogoArchivo(ByVal hWnd As Long, ByVal Titulo As String, ByVal Filtro As String, ByVal DirInicial As String) As String
    Dim ofn As OPENFILENAME
    Dim sFiltroFormateado As String
    Dim res As Long
    
    ' El filtro de la API de Windows requiere terminar estrictamente con doble caracter nulo
    sFiltroFormateado = Replace(Filtro, "|", Chr$(0))
    If Right$(sFiltroFormateado, 1) <> Chr$(0) Then sFiltroFormateado = sFiltroFormateado & Chr$(0)
    sFiltroFormateado = sFiltroFormateado & Chr$(0)
    
    With ofn
        .lStructSize = 76 ' <-- TamaÒo est·ndar compatible para VB6 en Windows 10/11
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        .lpstrTitle = Titulo
        .lpstrFilter = sFiltroFormateado
        .nFilterIndex = 1
        ' El b˙fer debe inicializarse relleno de nulos para no bloquear el di·logo
        .lpstrFile = String$(1024, 0)
        .nMaxFile = 1024
        .lpstrFileTitle = String$(512, 0)
        .nMaxFileTitle = 512
        .lpstrInitialDir = DirInicial
        .flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_EXPLORER
    End With
    
    res = GetOpenFileName(ofn)
    
    If res <> 0 Then
        Dim posNulo As Long
        posNulo = InStr(ofn.lpstrFile, Chr$(0))
        If posNulo > 0 Then
            AbrirDialogoArchivo = Left$(ofn.lpstrFile, posNulo - 1)
        Else
            AbrirDialogoArchivo = Trim$(ofn.lpstrFile)
        End If
    Else
        AbrirDialogoArchivo = ""
    End If
End Function

Public Function GuardarDialogoArchivo(ByVal hWnd As Long, ByVal Titulo As String, ByVal Filtro As String, ByVal DirInicial As String, ByVal DefExt As String) As String
    Dim ofn As OPENFILENAME
    Dim res As Long
    
    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = hWnd
        .lpstrTitle = Titulo
        .lpstrFilter = Replace(Filtro, "|", Chr$(0)) & Chr$(0) & Chr$(0)
        .lpstrFile = Space$(260)
        .nMaxFile = 260
        .lpstrInitialDir = DirInicial
        .lpstrDefExt = DefExt
        .flags = OFN_PATHMUSTEXIST Or OFN_OVERWRITEPROMPT Or OFN_EXPLORER
    End With
    
    res = GetSaveFileName(ofn)
    If res <> 0 Then
        GuardarDialogoArchivo = Trim$(Left$(ofn.lpstrFile, InStr(ofn.lpstrFile, Chr$(0)) - 1))
    Else
        GuardarDialogoArchivo = ""
    End If
End Function
