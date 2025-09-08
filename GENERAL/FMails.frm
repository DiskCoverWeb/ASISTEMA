VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form FEnviarMails 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ENVIOS DE MAILS EN FORMA MASIVA"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11160
   Icon            =   "FMails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF8080&
      Caption         =   "&Salir"
      Height          =   750
      Left            =   1680
      Picture         =   "FMails.frx":0696
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   105
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "&Emails >>"
      Height          =   750
      Left            =   105
      Picture         =   "FMails.frx":0F60
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   105
      Width           =   1485
   End
   Begin VB.TextBox TxtMemo 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2115
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   2520
      Width           =   10935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "E&xaminar"
      Height          =   330
      Left            =   9975
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1785
      Width           =   1065
   End
   Begin VB.TextBox TxtArchivoAdjunto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2205
      TabIndex        =   9
      Top             =   1785
      Width           =   7680
   End
   Begin MSDataGridLib.DataGrid DGPara 
      Bindings        =   "FMails.frx":182A
      Height          =   2325
      Left            =   105
      TabIndex        =   14
      ToolTipText     =   "<Ctrl+M>Enviar Indivudualmente, <Ctrl+F1> Genera Reporte por Excel"
      Top             =   4935
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   4101
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoPara 
      Height          =   330
      Left            =   945
      Top             =   5985
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Para"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox TxtAsunto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2205
      TabIndex        =   7
      Top             =   1365
      Width           =   8835
   End
   Begin MSAdodcLib.Adodc AdoMemoNo 
      Height          =   330
      Left            =   945
      Top             =   6300
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "MemoNo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoMemo 
      Height          =   330
      Left            =   945
      Top             =   6615
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Memo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CDialogDir 
      Left            =   9870
      Top             =   7350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " D E T A L L E   D E L   C O R R E O"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   105
      TabIndex        =   11
      Top             =   2205
      Width           =   10935
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   9030
      TabIndex        =   3
      Top             =   525
      Width           =   2010
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ARCHIVO ADJUNTO:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   105
      TabIndex        =   8
      Top             =   1785
      Width           =   2115
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   10500
      Top             =   7350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FMails.frx":1840
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FMails.frx":1B5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FMails.frx":1E74
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FMails.frx":20B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ASUNTO MAIL:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Top             =   1365
      Width           =   2115
   End
   Begin VB.Label LblEmailDe 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2205
      TabIndex        =   5
      Top             =   945
      Width           =   8835
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ENVIO &No."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   7665
      TabIndex        =   2
      Top             =   525
      Width           =   1380
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ATENCION:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   9030
      TabIndex        =   1
      Top             =   105
      Width           =   2010
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   7665
      TabIndex        =   0
      Top             =   105
      Width           =   1380
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " L I S T A D O   D E   P E R S O N A S"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   105
      TabIndex        =   13
      Top             =   4620
      Width           =   10935
   End
   Begin VB.Label LblDe 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   105
      TabIndex        =   16
      Top             =   7665
      Width           =   6000
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Atentamente,"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   105
      TabIndex        =   15
      Top             =   7350
      Width           =   1905
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " REMITENTE:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   945
      Width           =   2115
   End
End
Attribute VB_Name = "FEnviarMails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  TxtArchivoAdjunto = ""
  RutaOrigen = SelectDialogFile(CDialogDir)
  If RutaOrigen <> "" Then TxtArchivoAdjunto = RutaOrigen
End Sub

Private Sub Command2_Click()
   Envio_Masivo_Correos
End Sub

Private Sub Command3_Click()
  Unload FEnviarMails
End Sub

' Este correo es una prueba de verificacion del registro de su correo, reenviar para confirmar que estan correctos, o llamar a confirmar al 022238218
Private Sub DGPara_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ParaAux As String
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyM Then
     Label11.Caption = Year(FechaSistema) & "-" & Month(FechaSistema) & "-" & Format(ReadSetDataNum("Envio No", True, True), "00000000")
     With AdoPara.Recordset
      If .RecordCount > 0 Then
          TMail.Adjunto = ""
          TMail.Asunto = ""
          TMail.Mensaje = ""
          TMail.para = ""
         'Enviar Mail Individual
          TMail.Adjunto = TxtArchivoAdjunto
          TMail.Asunto = TxtAsunto
          TMail.Mensaje = LblEmailDe.Caption & " " & vbCrLf _
                        & "Código de Envio: " & Label11.Caption & " " & vbCrLf _
                        & TxtMemo
          Envio_Individual .Fields("Email")
          Envio_Individual .Fields("Email2")
      End If
     End With
  End If
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGPara.Visible = False
     GenerarDataTexto FEnviarMails, AdoPara
     DGPara.Visible = True
  End If
End Sub

Private Sub Form_Activate()
  Label11.Caption = Year(FechaSistema) & "-" & Month(FechaSistema) & "-" & Format(ReadSetDataNum("Envio No", True, False), "00000000")
  LblEmailDe.Caption = Empresa & " (" & NombreUsuario & ")"
  sSQL = "SELECT Cliente,CI_RUC,Email,Email2,Grupo,Representante " _
       & "FROM Clientes " _
       & "WHERE (LEN(Email)+LEN(Email2)) > 2 " _
       & "ORDER BY Cliente,Representante "
  LblDe.Caption = NombreUsuario
  Label8.Caption = FechaSistema
  Select_Adodc_Grid DGPara, AdoPara, sSQL
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FEnviarMails
  ConectarAdodc AdoPara
  ConectarAdodc AdoMemo
  ConectarAdodc AdoMemoNo
End Sub

Private Sub TxtAsunto_GotFocus()
  MarcarTexto TxtAsunto
End Sub

Private Sub TxtAsunto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtAsunto_LostFocus()
  TextoValido TxtAsunto, , True
End Sub

Private Sub TxtMemo_GotFocus()
  MarcarTexto TxtMemo
End Sub

Public Sub Envio_Individual(CorreoElectronico As String)
Dim ParaAux As String
    If Len(CorreoElectronico) > 1 Then
       ParaAux = Sin_Signos_Especiales(LCase(CorreoElectronico))
       ParaAux = Replace(ParaAux, "|", "")
       TMail.Credito_No = Ninguno
       Do While Len(ParaAux) > 0
          If InStr(ParaAux, ";") > 0 Then
             TMail.para = MidStrg(ParaAux, 1, InStr(ParaAux, ";") - 1)
             If Len(TMail.para) > 1 And EsUnEmail(TMail.para) Then FEnviarCorreos.Show 1
             ParaAux = TrimStrg(MidStrg(ParaAux, InStr(ParaAux, ";") + 1, Len(ParaAux)))
          ElseIf InStr(ParaAux, ",") > 0 Then
             TMail.para = MidStrg(ParaAux, 1, InStr(ParaAux, ",") - 1)
             If Len(TMail.para) > 1 And EsUnEmail(TMail.para) Then FEnviarCorreos.Show 1
             ParaAux = TrimStrg(MidStrg(ParaAux, InStr(ParaAux, ",") + 1, Len(ParaAux)))
          Else
             TMail.para = TrimStrg(ParaAux)
             ParaAux = ""
          End If
       Loop
       If Len(TMail.para) > 1 And EsUnEmail(TMail.para) Then FEnviarCorreos.Show 1
    End If
End Sub

Public Sub Envio_Masivo_Correos()
Dim LosMails As String
    DGPara.Visible = False
    TMail.ListaMail = 0
    With AdoPara.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
         TMail.Adjunto = TxtArchivoAdjunto
         TMail.Asunto = TxtAsunto
         Do While Not .EOF
           'Enviar Mail Individual
            Label11.Caption = Year(FechaSistema) & "-" & Month(FechaSistema) & "-" & Format(ReadSetDataNum("Envio No", True, True), "00000000")
            TMail.Mensaje = LblEmailDe.Caption & " " & vbCrLf _
                          & "Código de Envio: " & Label11.Caption & " " & vbCrLf _
                          & TxtMemo _
                          & " " & vbCrLf _
                          & " " & vbCrLf _
                          & String(40, "_") & vbCrLf _
                          & "Este mensaje fue enviado a Clientes de DiskCover System, " _
                          & "si no desea recibir mas correos envie un mail con las " _
                          & "palabras: NO RECIBIR MAS " & vbCrLf
            LosMails = ""
            If Len(.Fields("Email")) > 1 Then LosMails = LosMails & TrimStrg(.Fields("Email")) & ";"
            If Len(.Fields("Email2")) > 1 Then LosMails = LosMails & TrimStrg(.Fields("Email2")) & ";"
            LosMails = TrimStrg(Replace(LosMails, " ", ""))
            Envio_Individual LosMails
           .MoveNext
         Loop
     End If
    End With
    DGPara.Visible = True
    RatonNormal
    MsgBox "Proceso terminado exitosamente"
End Sub
