VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FComision 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ASIGNACION DE COMISIONES"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
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
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   4305
      MaxLength       =   6
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "FComision.frx":0000
      Top             =   525
      Width           =   960
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Continuar"
      Height          =   330
      Left            =   5355
      TabIndex        =   4
      Top             =   525
      Width           =   1065
   End
   Begin MSDataGridLib.DataGrid DGComision 
      Bindings        =   "FComision.frx":0002
      Height          =   1590
      Left            =   105
      TabIndex        =   3
      Top             =   945
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   2805
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      AllowDelete     =   -1  'True
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
            LCID            =   3082
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
            LCID            =   3082
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
   Begin MSDataListLib.DataCombo DCEjec 
      Bindings        =   "FComision.frx":001C
      DataSource      =   "AdoEjecutivo"
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   8388608
      ForeColor       =   16777215
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoComision 
      Height          =   330
      Left            =   315
      Top             =   1155
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Comision"
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
   Begin MSAdodcLib.Adodc AdoEjecutivo 
      Height          =   330
      Left            =   315
      Top             =   1470
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Ejecutivo"
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Porcentaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   1260
      TabIndex        =   6
      Top             =   525
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Porcentaje"
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
      Top             =   525
      Width           =   1170
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Porcentaje"
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
      Left            =   3150
      TabIndex        =   1
      Top             =   525
      Width           =   1170
   End
End
Attribute VB_Name = "FComision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Idx As Long

Private Sub Command4_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT CR.Codigo,C.Cliente,C.CI_RUC,C.Porc_C " _
       & "FROM Clientes As C,Catalogo_Rol_Pagos As CR " _
       & "WHERE CR.Item = '" & NumEmpresa & "' " _
       & "AND CR.Periodo = '" & Periodo_Contable & "' " _
       & "AND C.Codigo = CR.Codigo " _
       & "ORDER BY C.Cliente "
  SelectDBCombo DCEjec, AdoEjecutivo, sSQL, "Cliente"
  SQL1 = "SELECT MAX(ID) As IDMax " _
       & "FROM Trans_Comision " _
       & "WHERE Item <> '.' "
  SelectAdodc AdoComision, SQL1
  If AdoComision.Recordset.RecordCount > 1 Then
     Idx = AdoComision.Recordset.Fields("IDMax") + 1
  Else
     Idx = 1
  End If
  Text1.Text = DatInv.Utilidad * 100
  Label2.Caption = Format$(Real1, "#,##0.00")
  Listado_de_Comisiones
End Sub

Private Sub Form_Load()
  CentrarForm FComision
  ConectarAdodc AdoComision
  ConectarAdodc AdoEjecutivo
End Sub

Private Sub Text1_GotFocus()
  MarcarTexto Text1
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Text1_LostFocus()
  If Val(Text1.Text) > 0 Then
     SetAdoAddNew "Trans_Comision"
     SetAdoFields "T", Pendiente
     SetAdoFields "Codigo_Inv", Codigos
     SetAdoFields "CodigoC", CodigoEjecutivo
     SetAdoFields "TC", TipoFactura
     SetAdoFields "Factura", 0
     SetAdoFields "Total", Real1
     SetAdoFields "Porc", Redondear(Val(Text1.Text) / 100, 4)
     SetAdoFields "Valor_Pagar", CCur(Real1 * Val(Text1.Text) / 100)
     SetAdoFields "Cta", Cta_Comision
     SetAdoFields "Item", NumEmpresa
     SetAdoFields "CodigoU", CodigoUsuario
     SetAdoFields "ID", Idx
     SetAdoUpdate
     Idx = Idx + 1
  End If
  Listado_de_Comisiones
  DCEjec.SetFocus
End Sub

Private Sub DCEjec_GotFocus()
  CodigoEjecutivo = Ninguno
End Sub

Private Sub DCEjec_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then
     'DCEjec.Visible = False
     'TextComEjec.Text = "0"
  End If
End Sub

Private Sub DCEjec_LostFocus()
  With AdoEjecutivo.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente Like '" & DCEjec.Text & "' ")
       If Not .EOF Then
          CodigoEjecutivo = .Fields("Codigo")
       Else
          MsgBox "Ejecutivo no Asignado"
       End If
   Else
       MsgBox "No existen datos"
   End If
  End With
End Sub

Public Sub Listado_de_Comisiones()
  sSQL = "SELECT Total,Porc,Valor_Pagar,Codigo_Inv,CodigoC,Fecha,ID " _
       & "FROM Trans_Comision " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND Codigo_Inv = '" & Codigos & "' " _
       & "AND Factura = 0 " _
       & "ORDER BY ID "
  SQLDec = "Valor_Pagar 4|."
  SelectDataGrid DGComision, AdoComision, sSQL, SQLDec
End Sub
