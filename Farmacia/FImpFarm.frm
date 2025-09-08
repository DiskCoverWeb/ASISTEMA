VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Begin VB.Form FImpFarm 
   Caption         =   "IMPORTACION DE CLIENTES DE FARMACIA"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet URLinet 
      Left            =   4725
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ListBox LstAfiliados 
      Height          =   3375
      Left            =   105
      TabIndex        =   2
      Top             =   735
      Width           =   8520
   End
   Begin MSAdodcLib.Adodc AdoAfiliados 
      Height          =   330
      Left            =   105
      Top             =   3150
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Adodc1"
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
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   540
      Left            =   1995
      TabIndex        =   1
      Top             =   105
      Width           =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Importar Archivo Plano"
      Height          =   540
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1800
   End
   Begin MSComDlg.CommonDialog CDialogDir 
      Left            =   3885
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "FImpFarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim NumFile As Long
Dim IdAfi As Integer
Dim InsAfi As Boolean
  RutaOrigen = SelectDialogFile(RutaSysBases)
  If RutaOrigen <> "" Then
     RatonReloj
     Contador = 0
     NumFile = FreeFile
     Open RutaOrigen For Input As #NumFile
     Do While Not EOF(NumFile)
        Line Input #NumFile, Cadena
        Contador = Contador + 1
     Loop
     Close #NumFile
     IdAfi = 0
     LstAfiliados.Clear
     NumFile = FreeFile
     Open RutaOrigen For Input As #NumFile
     Do While Not EOF(NumFile)
        Line Input #NumFile, Cadena
        CodigoCli = Trim(Mid(Cadena, 1, 10))
        NombreCliente = Trim(Mid(Cadena, 11, 60))
        sSQL = "SELECT * " _
             & "FROM Clientes " _
             & "WHERE CI_RUC = '" & CodigoCli & "' "
        Select_Adodc AdoAfiliados, sSQL
        If AdoAfiliados.Recordset.RecordCount <= 0 Then
           DigVerif = Digito_Verificador(CodigoCli)
           Caracter = Mid(CodigoCli, 10, 1)
           Select Case Tipo_RUC_CI.Tipo_Beneficiario
             Case "C", "P", "R"
                  SetAdoAddNew "Clientes"
                  SetAdoFields "T", Normal
                  SetAdoFields "FA", True
                  SetAdoFields "Codigo", Tipo_RUC_CI.Codigo_RUC_CI
                  SetAdoFields "TD", Tipo_RUC_CI.Tipo_Beneficiario
                  SetAdoFields "CI_RUC", CodigoCli
                  SetAdoFields "Cliente", UCase(NombreCliente)
                  SetAdoFields "Fecha", FechaSistema
                  SetAdoFields "Direccion", "SD"
                  SetAdoFields "DirNumero", "SN"
                  SetAdoFields "Ciudad", "QUITO"
                  SetAdoFields "Prov", "17"
                  SetAdoFields "Pais", "593"
                  SetAdoFields "Grupo", NumEmpresa
                  SetAdoFields "CodigoU", CodigoUsuario
                  SetAdoUpdate
                  
                 'Matriculados
                  SetAdoAddNew "Clientes_Matriculas"
                  SetAdoFields "T", Normal
                  SetAdoFields "Codigo", Tipo_RUC_CI.Codigo_RUC_CI
                  SetAdoFields "TD", "R"
                  SetAdoFields "Representante", "CENTRO MEDICO MATERNAL PAEZ ALMEIDA Y NARANJO"
                  SetAdoFields "Cedula_R", "1790764575001"
                  SetAdoFields "Lugar_Trabajo_R", "GARCIA MORENO Y ESMERALDAS"
                  SetAdoFields "Telefono_R", "022282950"
                  SetAdoFields "Grupo_No", NumEmpresa
                  SetAdoFields "Item", NumEmpresa
                  SetAdoFields "Periodo", Periodo_Contable
                  SetAdoFields "CodigoU", CodigoUsuario
                  SetAdoUpdate
             
             Case Else
                  LstAfiliados.AddItem CodigoCli & " " & NombreCliente
           End Select
        End If
        IdAfi = IdAfi + 1
        FImpFarm.Caption = "Afiliados - " & RutaOrigen & ": " & Format(IdAfi / Contador, "#0.00%") & " Procesando..."
     Loop
     Close #NumFile
     IdAfi = 0
     NumFile = FreeFile
     Open RutaOrigen For Input As #NumFile
     Do While Not EOF(NumFile)
        Line Input #NumFile, Cadena
        CodigoCli = Trim(Mid(Cadena, 91, 10))
        NombreCliente = Trim(Mid(Cadena, 101, 60))
        sSQL = "SELECT * " _
             & "FROM Clientes " _
             & "WHERE CI_RUC = '" & CodigoCli & "' "
        Select_Adodc AdoAfiliados, sSQL
        If AdoAfiliados.Recordset.RecordCount <= 0 Then
           DigVerif = Digito_Verificador(CodigoCli)
           Caracter = Mid(CodigoCli, 10, 1)
           Select Case Tipo_RUC_CI.Tipo_Beneficiario
             Case "C", "P", "R"
                  SetAdoAddNew "Clientes"
                  SetAdoFields "T", Normal
                  SetAdoFields "FA", True
                  SetAdoFields "Codigo", Tipo_RUC_CI.Codigo_RUC_CI
                  SetAdoFields "TD", Tipo_RUC_CI.Tipo_Beneficiario
                  SetAdoFields "CI_RUC", CodigoCli
                  SetAdoFields "Cliente", UCase(NombreCliente)
                  SetAdoFields "Fecha", FechaSistema
                  SetAdoFields "Direccion", "SD"
                  SetAdoFields "DirNumero", "SN"
                  SetAdoFields "Ciudad", "QUITO"
                  SetAdoFields "Prov", "17"
                  SetAdoFields "Pais", "593"
                  SetAdoFields "Grupo", NumEmpresa
                  SetAdoFields "CodigoU", CodigoUsuario
                  SetAdoUpdate
                  
                 'Matriculados
                  SetAdoAddNew "Clientes_Matriculas"
                  SetAdoFields "T", Normal
                  SetAdoFields "Codigo", Tipo_RUC_CI.Codigo_RUC_CI
                  SetAdoFields "TD", "R"
                  SetAdoFields "Representante", "CENTRO MEDICO MATERNAL PAEZ ALMEIDA Y NARANJO"
                  SetAdoFields "Cedula_R", "1790764575001"
                  SetAdoFields "Lugar_Trabajo_R", "GARCIA MORENO Y ESMERALDAS"
                  SetAdoFields "Telefono_R", "022282950"
                  SetAdoFields "Grupo_No", NumEmpresa
                  SetAdoFields "Item", NumEmpresa
                  SetAdoFields "Periodo", Periodo_Contable
                  SetAdoFields "CodigoU", CodigoUsuario
                  SetAdoUpdate
             
             Case Else
                  LstAfiliados.AddItem CodigoCli & " " & NombreCliente
           End Select
        End If
        IdAfi = IdAfi + 1
        FImpFarm.Caption = "Beneficiarios - " & RutaOrigen & ": " & Format(IdAfi / Contador, "#0.00%") & " Procesando..."
     Loop
     LstAfiliados.Text = LstAfiliados.List(0)
     Close #NumFile
  End If
  RatonNormal
  MsgBox "Proceo Terminado"
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
   RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FImpFarm
  ConectarAdodc AdoAfiliados
End Sub
