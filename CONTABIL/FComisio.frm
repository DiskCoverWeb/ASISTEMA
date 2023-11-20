VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FComisiones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Comisiones y el I.E.S.S."
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3570
      Picture         =   "FComisio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1365
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3570
      Picture         =   "FComisio.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   525
      Width           =   855
   End
   Begin VB.TextBox TxtComision 
      Alignment       =   1  'Right Justify
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
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "FComisio.frx":1194
      Top             =   945
      Width           =   1695
   End
   Begin VB.TextBox TxtIESS 
      Alignment       =   1  'Right Justify
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
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "FComisio.frx":1198
      Top             =   1365
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc AdoCxCxP 
      Height          =   330
      Left            =   105
      Top             =   1785
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
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
      Caption         =   "CxCxP"
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
   Begin VB.Label LblCodigo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   1785
      TabIndex        =   2
      Top             =   525
      Width           =   1695
   End
   Begin VB.Label LblCliente 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ASIGNACION DE ROL DE PAGOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4320
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CODIGO"
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
      TabIndex        =   1
      Top             =   525
      Width           =   1695
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Comision"
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
      TabIndex        =   3
      Top             =   945
      Width           =   1695
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Aporte al I.E.S.S."
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
      TabIndex        =   5
      Top             =   1365
      Width           =   1695
   End
End
Attribute VB_Name = "FComisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Total_Comision = Round(Val(TxtComision.Text), 2)
  MontoAper = Round(Val(TxtIESS.Text), 2)
  If Total_Comision >= 0 Then
  sSQL = "UPDATE Trans_Rol_Pagos " _
       & "SET Comision = " & Total_Comision & "," _
       & "Salario = InLiqui + " & Total_Comision & "-AporteP-Emergencia-Dcts_Certif-Dcts_Aport_Adm-Descuento_CxC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Codigo = '" & CodigoCli & "' " _
       & "AND Fecha = #" & BuscarFecha(Mifecha) & "# "
  'MsgBox sSQL
  Ejecutar_SQL_SP sSQL
  End If
  If MontoAper >= 0 Then
  sSQL = "UPDATE Trans_Rol_Pagos " _
       & "SET AporteP = " & MontoAper & "," _
       & "Salario = InLiqui  + Comision - " & MontoAper & "-Emergencia-Dcts_Certif-Dcts_Aport_Adm-Descuento_CxC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Codigo = '" & CodigoCli & "' " _
       & "AND Fecha = #" & BuscarFecha(Mifecha) & "# "
  Ejecutar_SQL_SP sSQL
  End If
  Unload Me
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
  LblCodigo.Caption = CodigoCli
  LblCliente.Caption = NombreCliente
End Sub

Private Sub Form_Load()
  CentrarForm FRolPago
  ConectarAdodc AdoCxCxP
  FCtaAhorro.Caption = "ASIGNACION A ROL DE PAGOS"
End Sub

Private Sub TxtComision_GotFocus()
  MarcarTexto TxtComision
End Sub

Private Sub TxtComision_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtIESS_GotFocus()
  MarcarTexto TxtIESS
End Sub

Private Sub TxtIESS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub
