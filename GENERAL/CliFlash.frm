VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form FClientesFlash 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NUEVO ITEM DE PROCESO"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9570
   Icon            =   "CliFlash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar TBarCliente 
      Height          =   660
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Módulo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Beneficiario"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   9
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CxC"
            Object.ToolTipText     =   "Asignar a Cuentas por Cobrar en Contabilidad"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CxP"
            Object.ToolTipText     =   "Asignar a Cuentas por Pagar en Contabilidad"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   11
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Ahorros"
            Object.ToolTipText     =   "Asignar Cuenta de Ahorros"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "RolPago"
            Object.ToolTipText     =   "Asignar a Rol de Pago"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Facturacion"
            Object.ToolTipText     =   "Asignar Cliente de Facturación"
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&S"
      Height          =   330
      Left            =   9135
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   735
      Width           =   330
   End
   Begin VB.TextBox TxtEmail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      MaxLength       =   50
      TabIndex        =   13
      Top             =   2520
      Width           =   7995
   End
   Begin VB.ComboBox CProvincia 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3150
      TabIndex        =   19
      Text            =   "PICHINCHA"
      Top             =   3255
      Width           =   3165
   End
   Begin VB.ComboBox CNacion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   17
      Text            =   "ECUADOR"
      Top             =   3255
      Width           =   2955
   End
   Begin VB.ComboBox CCiudadS 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6405
      TabIndex        =   21
      Top             =   3255
      Width           =   3060
   End
   Begin VB.TextBox TxtGrupo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      MaxLength       =   8
      TabIndex        =   7
      Top             =   1785
      Width           =   1275
   End
   Begin VB.TextBox TxtTelefonoS 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8190
      MaxLength       =   10
      TabIndex        =   15
      Top             =   2520
      Width           =   1275
   End
   Begin VB.TextBox TxtNumero 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7665
      MaxLength       =   10
      TabIndex        =   11
      Top             =   1785
      Width           =   1800
   End
   Begin VB.TextBox TxtDirS 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1470
      MaxLength       =   50
      TabIndex        =   9
      Top             =   1785
      Width           =   6105
   End
   Begin VB.TextBox TxtApellidosS 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1785
      MaxLength       =   180
      TabIndex        =   5
      Top             =   1050
      Width           =   7680
   End
   Begin VB.TextBox TxtCI_RUC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      MaxLength       =   13
      TabIndex        =   3
      ToolTipText     =   "<Alt+F2> Codigo Automático"
      Top             =   1050
      Width           =   1590
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   210
      Top             =   3675
      Visible         =   0   'False
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
      Caption         =   "Aux"
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
   Begin MSAdodcLib.Adodc AdoListCtas 
      Height          =   330
      Left            =   2205
      Top             =   3675
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "ListCtas"
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
   Begin VB.Label LblSRI 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4110
      Left            =   105
      TabIndex        =   24
      Top             =   3675
      Width           =   9360
   End
   Begin VB.Label LblCodigo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "XXXXXXXXXX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7560
      TabIndex        =   1
      Top             =   315
      Width           =   1905
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CODIGO BENEFIC."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   7560
      TabIndex        =   0
      Top             =   0
      Width           =   1905
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " EMAIL PRINCIPAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   105
      TabIndex        =   12
      Top             =   2205
      Width           =   7995
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   9765
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   19
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":0EFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":1218
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":1532
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":184C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":1B66
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":1E80
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":219A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":24B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":27CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":2AE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":1187A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":11B94
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":11EAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":121C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":1237A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":12B94
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CliFlash.frx":12DD2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CIUDAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   6405
      TabIndex        =   20
      Top             =   2940
      Width           =   3060
   End
   Begin VB.Label Label35 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PROVINCIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   3150
      TabIndex        =   18
      Top             =   2940
      Width           =   3165
   End
   Begin VB.Label Label34 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NACIONALIDAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   105
      TabIndex        =   16
      Top             =   2940
      Width           =   2955
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " GRUPO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Top             =   1470
      Width           =   1275
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TELEFONO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   8190
      TabIndex        =   14
      Top             =   2205
      Width           =   1275
   End
   Begin VB.Label Label38 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DIR. GEOGRAFICA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   7665
      TabIndex        =   10
      Top             =   1470
      Width           =   1800
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DIRECCION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   1470
      TabIndex        =   8
      Top             =   1470
      Width           =   6105
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " APELLIDOS Y NOMBRES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   1785
      TabIndex        =   4
      Top             =   735
      Width           =   7365
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " C.I./R.U.C."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Top             =   735
      Width           =   1590
   End
End
Attribute VB_Name = "FClientesFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Temp_Cliente As String
Dim Temp_CI_RUC As String

Public Sub GrabarClienteFlash()
  Nuevo = False
  T = Normal
  TextoValido TxtCI_RUC, , True
  TextoValido TxtApellidosS, , True
  TextoValido TxtDirS, , True
  TextoValido TxtNumero, , True
  TextoValido TxtTelefonoS, , True
  TextoValido TxtGrupo, , True
  TextoValido TxtEmail, , True
  If TxtCI_RUC.Text = Ninguno Then
     MsgBox "No se puede grabar, La C.I./R.U.C. Deben tener valores"
  Else
     Mensajes = "Esta seguro de Grabar"
     Titulo = "Pregunta de Grabación"
     If BoxMensaje = vbYes Then
        Control_Procesos Normal, "Insertar Clientes"
        RatonReloj
        Codigo = Ninguno
        CodigoCli = Ninguno
        CodigoCliente = Ninguno
        sSQL = "SELECT CI_RUC, Codigo " _
             & "FROM Clientes  " _
             & "WHERE CI_RUC = '" & TxtCI_RUC & "' "
        Select_Adodc AdoListCtas, sSQL
        If AdoListCtas.Recordset.RecordCount <= 0 Then
           Codigo = Tipo_RUC_CI.Codigo_RUC_CI
           
           sSQL = "SELECT " & Full_Fields("Clientes") & " " _
                & "FROM Clientes " _
                & "WHERE Codigo = '" & Codigo & "' "
           Select_Adodc AdoListCtas, sSQL
           If AdoListCtas.Recordset.RecordCount <= 0 Then
             'MsgBox TxtApellidosS
              SetAddNew AdoListCtas
              SetFields AdoListCtas, "T", T
              SetFields AdoListCtas, "Codigo", Codigo
              SetFields AdoListCtas, "Cliente", TxtApellidosS
              SetFields AdoListCtas, "CI_RUC", TxtCI_RUC
              SetFields AdoListCtas, "Direccion", TxtDirS
              SetFields AdoListCtas, "Telefono", Format$(TxtTelefonoS, "000000000")
              SetFields AdoListCtas, "DirNumero", TxtNumero
              SetFields AdoListCtas, "Ciudad", CCiudadS
              SetFields AdoListCtas, "Email", TxtEmail
              SetFields AdoListCtas, "TD", TipoBenef
              SetFields AdoListCtas, "Prov", "00"
              SetFields AdoListCtas, "Pais", "593"
              SetFields AdoListCtas, "Grupo", TxtGrupo
              SetFields AdoListCtas, "Parte_Relacionada", "NO"
              Select Case TipoBenef
                Case "C", "R"
                     SetFields AdoListCtas, "Tipo_Pasaporte", "00"
                Case Else
                     SetFields AdoListCtas, "Tipo_Pasaporte", "01"
              End Select
              Select Case Modulo
                Case "FACTURACION", "FARMACIA"
                     SetFields AdoListCtas, "FA", adTrue
              End Select
              SetUpdate AdoListCtas
              LblCodigo.Caption = Codigo
              FA.CodigoC = Codigo
              CodigoCli = Codigo
              CodigoCliente = Codigo
              NombreCliente = TxtApellidosS
              MsgBox "Grabacion Exitosa"
              Select Case Modulo
                Case "FACTURACION", "FARMACIA"
                     Unload FClientesFlash
                Case Else
                     TBarCliente.buttons("Grabar").Enabled = False
                     Command1.SetFocus
                     TxtTelefonoS.SetFocus
              End Select
           Else
              MsgBox "No se puede crear el Beneficiaro, ya existe"
              Unload FClientesFlash
           End If
        Else
           MsgBox "No se puede crear el Beneficiaro, ya existe"
           Unload FClientesFlash
        End If
     End If
  End If
End Sub

Private Sub CCiudadS_GotFocus()
  MarcarTexto CCiudadS
End Sub

Private Sub CNacion_GotFocus()
  MarcarTexto CNacion
End Sub

Private Sub CNacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CNacion_LostFocus()
  CProvincia.Clear
  sSQL = "SELECT * " _
       & "FROM Tabla_Naciones " _
       & "WHERE CProvincia <> '00' " _
       & "AND TR = 'P' " _
       & "AND CPais = '" & SinEspaciosIzq(CNacion) & "' " _
       & "ORDER BY CProvincia "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     CProvincia.Text = AdoAux.Recordset.Fields("CProvincia") & " " & AdoAux.Recordset.Fields("Descripcion_Rubro")
     Do While Not AdoAux.Recordset.EOF
        CProvincia.AddItem AdoAux.Recordset.Fields("CProvincia") & " " & AdoAux.Recordset.Fields("Descripcion_Rubro")
        AdoAux.Recordset.MoveNext
     Loop
  Else
     CProvincia.AddItem "99 OTRO"
     CProvincia.Text = "99 OTRO"
  End If
End Sub

Private Sub Command1_Click()
  Unload FClientesFlash
End Sub

Private Sub CProvincia_GotFocus()
  MarcarTexto CProvincia
End Sub

Private Sub CProvincia_LostFocus()
  CCiudadS.Clear
  sSQL = "SELECT * " _
       & "FROM Tabla_Naciones " _
       & "WHERE CProvincia <> '00' " _
       & "AND TR = 'C' " _
       & "AND CPais = '" & SinEspaciosIzq(CNacion) & "' " _
       & "AND CProvincia = '" & SinEspaciosIzq(CProvincia) & "' " _
       & "ORDER BY CCiudad "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     CCiudadS.Text = AdoAux.Recordset.Fields("Descripcion_Rubro")
     Do While Not AdoAux.Recordset.EOF
        CCiudadS.AddItem AdoAux.Recordset.Fields("Descripcion_Rubro")
        AdoAux.Recordset.MoveNext
     Loop
  Else
     CCiudadS.AddItem "OTRO"
     CCiudadS.Text = "OTRO"
  End If
End Sub

Private Sub Form_Activate()
  FClientesFlash.Caption = "CREACION DE CLIENTE NUEVO"
  If Modulo = "FACTURACION" Then
     TBarCliente.buttons("CxC").Enabled = False
     TBarCliente.buttons("CxP").Enabled = False
     TBarCliente.buttons("Ahorros").Enabled = False
     TBarCliente.buttons("RolPago").Enabled = False
  End If
  NombreCliente = UCaseStrg(NombreCliente)
  LblCodigo.Caption = "Ninguno"
  TxtApellidosS = ""
  TxtCI_RUC.Text = ""
  TxtDirS.Text = ""
  TxtNumero.Text = ""
  TxtApellidosS.Enabled = True
  'If NivelNo = "" Then NivelNo = NumEmpresa
  CNacion.Clear
  sSQL = "SELECT * " _
       & "FROM Tabla_Naciones " _
       & "WHERE TR = 'N' " _
       & "ORDER BY CPais,Descripcion_Rubro "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     CNacion.Text = "593 ECUADOR"
     Do While Not AdoAux.Recordset.EOF
        CNacion.AddItem AdoAux.Recordset.Fields("CPais") & " " & AdoAux.Recordset.Fields("Descripcion_Rubro")
        AdoAux.Recordset.MoveNext
     Loop
  End If
  CNacion.AddItem "999 OTRO"
  CProvincia.Clear
  sSQL = "SELECT * " _
       & "FROM Tabla_Naciones " _
       & "WHERE CProvincia <> '00' " _
       & "AND TR = 'P' " _
       & "AND CPais = '" & SinEspaciosIzq(CNacion) & "' " _
       & "ORDER BY CProvincia "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     CProvincia.Text = AdoAux.Recordset.Fields("CProvincia") & " " & AdoAux.Recordset.Fields("Descripcion_Rubro")
     Do While Not AdoAux.Recordset.EOF
        CProvincia.AddItem AdoAux.Recordset.Fields("CProvincia") & " " & AdoAux.Recordset.Fields("Descripcion_Rubro")
        AdoAux.Recordset.MoveNext
     Loop
  Else
     CProvincia.AddItem "99 OTRO"
     CProvincia.Text = "99 OTRO"
  End If
  sSQL = "SELECT Cliente, CI_RUC, Codigo " _
       & "FROM Clientes  "
  If IsNumeric(NombreCliente) Then
     TxtCI_RUC = NombreCliente
     Temp_Cliente = ""
     Temp_CI_RUC = NombreCliente
     sSQL = sSQL & "WHERE CI_RUC = '" & NombreCliente & "' "
  Else
     TxtApellidosS = NombreCliente
     Temp_Cliente = NombreCliente
     Temp_CI_RUC = ""
     sSQL = sSQL & "WHERE Cliente = '" & NombreCliente & "' "
  End If
  Select_Adodc AdoListCtas, sSQL
  If Bloquear_Control Then
     TBarCliente.buttons("Grabar").Enabled = False
     TBarCliente.buttons("CxC").Enabled = False
     TBarCliente.buttons("CxP").Enabled = False
     TBarCliente.buttons("Ahorros").Enabled = False
     TBarCliente.buttons("RolPago").Enabled = False
     TBarCliente.buttons("Facturacion").Enabled = False
  End If
  RatonNormal
  TxtCI_RUC.SetFocus
End Sub

Private Sub Form_Load()
    CentrarForm FClientesFlash
    LblSRI.Visible = False
    FClientesFlash.Height = Label34.Top + Label34.Height + 900
    ConectarAdodc AdoAux
    ConectarAdodc AdoListCtas
End Sub

Private Sub TBarCliente_ButtonClick(ByVal Button As ComctlLib.Button)
 CodigoCliente = LblCodigo.Caption
 NombreCliente = TxtApellidosS
 Select Case Button.key
   Case "Salir"
        RatonNormal
        Unload FClientesFlash
   Case "Grabar"
        GrabarClienteFlash
        RatonNormal
   Case "CxC"
        If Modulo = "CONTABILIDAD" Then
           SubCta = "C"
           Mensajes = "Asignar CxC a " & TxtApellidosS & "."
           Titulo = "Pregunta de CxC"
           If BoxMensaje = vbYes Then FCxCxP.Show 1
        Else
           MsgBox "Modulo sin permiso"
        End If
   Case "CxP"
        If Modulo = "CONTABILIDAD" Then
           SubCta = "P"
           Mensajes = "Asignar CxP a " & TxtApellidosS & "."
           Titulo = "Pregunta de CxP"
           If BoxMensaje = vbYes Then FCxCxP.Show 1
        Else
           MsgBox "Modulo sin permiso"
        End If
   Case "Ahorros"
        If LblCodigo.Caption = "Ninguno" Then
           MsgBox "No ha grabado el cliente, no se puede asignar Cuenta de Ahorro."
        Else
           Mensajes = "Asignar Cuenta de Ahorros a " & TxtApellidosS & "."
           Titulo = "Pregunta de Creación"
           If BoxMensaje = vbYes Then FCtaAhorro.Show 1
        End If
   Case "RolPago"
        If TipoBenef <> "C" Then
           MsgBox "Este tipo de Beneficiario" & vbCrLf _
                  & "no es valido en Nomina," & vbCrLf _
                  & "pero se asignará para procesos del Rol"
        End If
        If LblCodigo.Caption = "Ninguno" Then
           MsgBox "No ha grabado el cliente, no se puede asignar a Rol de Pagos"
        Else
           Mensajes = "Asignar Rol de Pagos a " & TxtApellidosS & "."
           Titulo = "Pregunta de Creación"
           If BoxMensaje = vbYes Then
              'FechaValida MBFecha
              If Modulo = "ROL PAGOS" Then
                 FRolPago.Show 1
              Else
                 sSQL = "SELECT * " _
                      & "FROM Catalogo_Rol_Pagos " _
                      & "WHERE Codigo = '" & CodigoCli & "' " _
                      & "AND Periodo = '" & Periodo_Contable & "' " _
                      & "AND Item = '" & NumEmpresa & "' "
                 Select_Adodc AdoAux, sSQL
                 If AdoAux.Recordset.RecordCount <= 0 Then SetAddNew AdoAux
                 SetFields AdoAux, "Fecha", FechaSistema
                 SetFields AdoAux, "Item", NumEmpresa
                 SetFields AdoAux, "Codigo", CodigoCliente
                 SetFields AdoAux, "CodigoU", CodigoUsuario
                 SetFields AdoAux, "T", Normal
                 SetFields AdoAux, "SN", "2"
                 SetUpdate AdoAux
              End If
           End If
        End If
   Case "Facturacion"
        Mensajes = "Asignar a Facturacion " & TxtApellidosS & "."
        Titulo = "Pregunta de Facturacion"
        If BoxMensaje = vbYes Then
           sSQL = "UPDATE Clientes " _
                & "SET FA = " & Val(adTrue) & " " _
                & "WHERE Codigo = '" & CodigoCliente & "' "
           Ejecutar_SQL_SP sSQL
        End If
 End Select
End Sub

Private Sub TxtApellidosS_GotFocus()
  MarcarTexto TxtApellidosS
End Sub

Private Sub TxtApellidosS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtApellidosS_LostFocus()
    TextoValido TxtApellidosS, , True
    If Leer_Campo_Cliente("Cliente", TxtApellidosS) <> "" Then
       MsgBox "Este Beneficiario ya está asignado"
       TxtCI_RUC.SetFocus
    End If
End Sub

Private Sub TxtCI_RUC_GotFocus()
  TxtGrupo = NivelNo
  MarcarTexto TxtCI_RUC
End Sub

Private Sub TxtCI_RUC_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If AltDown And KeyCode = vbKeyF2 Then TxtCI_RUC.Text = Leer_Codigo_Automatico
  If CtrlDown And KeyCode = vbKeyV Then TxtCI_RUC.Text = Clipboard.GetText()
End Sub

''Private Sub TxtCI_RUC_KeyPress(KeyAscii As Integer)
''   KeyAscii = Solo_Letras_Numeros(KeyAscii)
''End Sub

Private Sub TxtCI_RUC_LostFocus()
  TextoValido TxtCI_RUC, , True
  DigVerif = Digito_Verificador(TxtCI_RUC)
  If Tipo_RUC_CI.Tipo_Beneficiario = "P" Then
     Mensajes = "Este código es un Pasaporte"
     Titulo = "CONFIRMACION DE PASAPORTE"
     If BoxMensaje <> vbYes Then
        Tipo_RUC_CI.Tipo_Beneficiario = "O"
     Else
        DigVerif = Ninguno
     End If
  End If
  If DigVerif = "-" Then
     MsgBox "RUC/CEDULA INCORRECTA"
     TxtCI_RUC.SetFocus
  Else
     If Tipo_RUC_CI.Tipo_Beneficiario = "R" Then
        LblSRI.Visible = True
        FClientesFlash.Height = LblSRI.Top + LblSRI.Height + 600
        Label4.Caption = "* R A Z O N    S O C I A L"
        TipoSRI = consulta_RUC_SRI(TxtCI_RUC)
        If Len(TipoSRI.RazonSocial) > 1 Then
           TipoSRI.RazonSocial = Replace(TipoSRI.RazonSocial, "&", "Y")
           TxtApellidosS = TipoSRI.RazonSocial
           With TipoSRI
              If Len(.RUC_SRI) > 1 Then
                 Mensajes = Mensajes & "R.U.C.: " & .RUC_SRI
                 If Len(.Estado) > 1 Then
                    Mensajes = Mensajes & vbTab & vbTab & "ESTADO DEL CONTRIBUYENTE: """ & UCaseStrg(.Estado) & """ "
                 Else
                    Mensajes = Mensajes & vbCrLf
                 End If
              End If
              If Len(.RazonSocial) > 1 Then Mensajes = Mensajes & "RAZON SOCIAL: " & .RazonSocial & vbCrLf
              If Len(.NombreComercial) > 1 Then Mensajes = Mensajes & "NOMBRE COMERCIAL: " & .NombreComercial & vbCrLf
              If Len(.TipoRUC) > 1 Then Mensajes = Mensajes & UCaseStrg(.TipoRUC) & ", "
              If Len(.Obligado) > 1 Then Mensajes = Mensajes & .Obligado & " OBLIGADO A LLEVAR CONTABILIDAD" & vbCrLf
              If Len(.ActividadEconomica) > 1 Then Mensajes = Mensajes & "ACTIVIDAD ECONOMICA: " & .ActividadEconomica & vbCrLf
              If Len(.FechaInicio) > 1 Then Mensajes = Mensajes & "INICIO SU ACTIVIDAD EL " & .FechaInicio & vbCrLf
              If Len(.FechaActualización) > 1 Then Mensajes = Mensajes & "R.U.C. ACTUALIZADO EL " & .FechaActualización & vbCrLf
              If Len(.FechaReinicio) > 1 Then Mensajes = Mensajes & "REINICIO DE ACTIVIDADES: " & .FechaReinicio & vbCrLf
              If Len(.Categoria) > 1 And Len(.ClaseRUC) > 1 Then Mensajes = Mensajes & "CATEGORIA: " & .Categoria & ", CLASE: " & .ClaseRUC & vbCrLf
              If Len(.FechaCese) > 1 Then Mensajes = Mensajes & "CESE DE ACTIVIDADES: " & .FechaCese & vbCrLf
              If Len(.MicroEmpresa) > 1 Then Mensajes = Mensajes & "TIPO DE CONTRIBUYENTE: """ & UCaseStrg(.MicroEmpresa) & """ " & vbCrLf
              If Len(.AgenteRetencion) > 1 Then Mensajes = Mensajes & "AGENTE DE RETENCION: """ & UCaseStrg(.AgenteRetencion) & """ " & vbCrLf
           End With
           LblSRI.Caption = Mensajes
        End If
     Else
        LblSRI.Visible = False
        FClientesFlash.Height = Label34.Top + Label34.Height + 900
        Label4.Caption = "* APELLIDOS Y NOMBRES"
     End If
     Label6.Caption = "* C.I./R.U.C.  [" & Tipo_RUC_CI.Tipo_Beneficiario & "]"
     Cadena = Leer_Campo_Cliente("CI_RUC", TxtCI_RUC)
     If Cadena <> "" Then
        MsgBox "ESTE CODIGO YA ESTA ASIGNADO HA" & vbCrLf & vbCrLf & Cadena
        TxtCI_RUC.SetFocus
     End If
     TipoBenef = Tipo_RUC_CI.Tipo_Beneficiario
  End If
End Sub

Private Sub TxtDirS_GotFocus()
  MarcarTexto TxtDirS
End Sub

Private Sub TxtDirS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDirS_LostFocus()
  TextoValido TxtDirS, , True
  If TxtDirS.Text = Ninguno Then TxtDirS.Text = "SD"
End Sub

Private Sub TxtEmail_GotFocus()
  MarcarTexto TxtEmail
End Sub

Private Sub TxtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

'Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
'  KeyAscii = Solo_Letras_Numeros(KeyAscii)
'End Sub

Private Sub TxtEmail_LostFocus()
  TxtEmail = LCase(TxtEmail)
End Sub

Private Sub TxtGrupo_GotFocus()
  MarcarTexto TxtGrupo
End Sub

Private Sub TxtGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtGrupo_LostFocus()
    TextoValido TxtGrupo, , True
    sSQL = "SELECT Direccion " _
         & "FROM Clientes " _
         & "WHERE Grupo = '" & TxtGrupo.Text & "' " _
         & "AND FA <> " & Val(adFalse) & " "
    If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
    Select_Adodc AdoAux, sSQL
    If AdoAux.Recordset.RecordCount > 0 Then
       If Len(AdoAux.Recordset.Fields("Direccion")) > 1 Then TxtDirS.Text = AdoAux.Recordset.Fields("Direccion") Else TxtDirS.Text = "SD"
    Else
       TxtDirS.Text = "SD"
    End If
End Sub

Private Sub TxtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumero_LostFocus()
  TextoValido TxtNumero, , True
  If TxtNumero.Text = "" Then TxtNumero.Text = "SN"
  
End Sub

Private Sub TxtTelefonoS_GotFocus()
  MarcarTexto TxtTelefonoS
End Sub

Private Sub TxtTelefonoS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtTelefonoS_LostFocus()
  TextoValido TxtTelefonoS, , True
  TxtTelefonoS.Text = Format$(Val(TxtTelefonoS.Text), "000000000")
End Sub

Private Sub CProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CCiudadS_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

