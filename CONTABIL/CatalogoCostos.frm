VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FCatalogo_Costos 
   Caption         =   "ASIGNACION DE SUBMODULOS POR PROYECTOS"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19755
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8970
   ScaleWidth      =   19755
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10290
      Top             =   6615
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CatalogoCostos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CatalogoCostos.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CatalogoCostos.frx":0D2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CatalogoCostos.frx":1046
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CatalogoCostos.frx":1920
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGCostos 
      Bindings        =   "CatalogoCostos.frx":2572
      Height          =   4635
      Left            =   105
      TabIndex        =   6
      Top             =   1680
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   8176
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin VB.CommandButton Command1 
      Caption         =   "&S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   11235
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6720
      Width           =   330
   End
   Begin MSAdodcLib.Adodc AdoCostos 
      Height          =   330
      Left            =   105
      Top             =   6615
      Width           =   5790
      _ExtentX        =   10213
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
      Caption         =   "Costos"
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
   Begin MSAdodcLib.Adodc AdoSubModulo 
      Height          =   330
      Left            =   840
      Top             =   2415
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
      Caption         =   "SubModulo"
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
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   840
      Top             =   2730
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
      Caption         =   "Ctas"
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
   Begin MSDataListLib.DataCombo DCProyecto 
      Bindings        =   "CatalogoCostos.frx":258A
      DataSource      =   "AdoCtas"
      Height          =   315
      Left            =   1155
      TabIndex        =   1
      Top             =   840
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Ctas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoCtas1 
      Height          =   330
      Left            =   840
      Top             =   3045
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
      Caption         =   "Ctas1"
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
   Begin MSDataListLib.DataCombo DCCtasProyecto 
      Bindings        =   "CatalogoCostos.frx":25A0
      DataSource      =   "AdoCtas1"
      Height          =   315
      Left            =   10815
      TabIndex        =   3
      Top             =   840
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Ctas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCSubModulos 
      Bindings        =   "CatalogoCostos.frx":25B7
      DataSource      =   "AdoSubModulo"
      Height          =   315
      Left            =   2520
      TabIndex        =   5
      Top             =   1260
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Ctas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   840
      Top             =   3465
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   19755
      _ExtentX        =   34846
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Enviar a Excel"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Todas"
            Object.ToolTipText     =   "Listar todos los Costos"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "SubModulo del Proyecto"
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
      TabIndex        =   4
      Top             =   1260
      Width           =   2325
   End
   Begin VB.Label Label2 
      Caption         =   "Cuentas del Proyecto"
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
      Left            =   8715
      TabIndex        =   2
      Top             =   840
      Width           =   2010
   End
   Begin VB.Label Label1 
      Caption         =   "Proyecto"
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
      TabIndex        =   0
      Top             =   840
      Width           =   960
   End
End
Attribute VB_Name = "FCatalogo_Costos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CodSubCta As String
  
Private Sub Command1_Click()
  Unload FCatalogo_Costos
End Sub

Private Sub DCCtasProyecto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCtasProyecto_LostFocus()
  SubCta = Ninguno
  If AdoCtas1.Recordset.RecordCount > 0 Then
     AdoCtas1.Recordset.MoveFirst
     AdoCtas1.Recordset.Find ("Cuenta = '" & DCCtasProyecto.Text & "' ")
     If Not AdoCtas1.Recordset.EOF Then SubCta = AdoCtas1.Recordset.Fields("Codigo")
  End If
End Sub

Private Sub DCProyecto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCProyecto_LostFocus()
  If AdoCtas.Recordset.RecordCount > 0 Then
     AdoCtas.Recordset.MoveFirst
     AdoCtas.Recordset.Find ("Cuenta = '" & DCProyecto.Text & "' ")
     If Not AdoCtas.Recordset.EOF Then CodSubCta = AdoCtas.Recordset.Fields("Codigo") Else CodSubCta = Ninguno
  End If
  
  sSQL = "SELECT Codigo, Cuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Codigo LIKE '" & CodSubCta & "%' " _
       & "AND DG = 'D' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCCtasProyecto, AdoCtas1, sSQL, "Cuenta"
End Sub

Private Sub DCSubModulos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCSubModulos_LostFocus()
  CodigoBenef = Ninguno
  If AdoSubModulo.Recordset.RecordCount > 0 Then
     AdoSubModulo.Recordset.MoveFirst
     AdoSubModulo.Recordset.Find ("Detalle = '" & DCSubModulos.Text & "' ")
     If Not AdoSubModulo.Recordset.EOF Then CodigoBenef = AdoSubModulo.Recordset.Fields("Codigo")
  End If
  
  If Len(SubCta) > 1 And Len(CodigoBenef) > 1 Then
     Titulo = "PREGUNTA DE GRABACION"
     Mensajes = "Esta seguro de Inserta en la cuenta:" & vbCrLf _
              & DCCtasProyecto.Text & vbCrLf & vbCrLf _
              & "El centro de costo: " & vbCrLf & DCSubModulos.Text
     If BoxMensaje = vbYes Then
        sSQL = "SELECT Cta, Codigo " _
             & "FROM Trans_Presupuestos " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Cta = '" & SubCta & "' " _
             & "AND Codigo = '" & CodigoBenef & "' "
        Select_Adodc AdoAux, sSQL
        If AdoAux.Recordset.RecordCount <= 0 Then
           SetAdoAddNew "Trans_Presupuestos"
           SetAdoFields "Cta", SubCta
           SetAdoFields "Codigo", CodigoBenef
           SetAdoFields "Item", NumEmpresa
           SetAdoFields "Periodo", Periodo_Contable
           SetAdoUpdate
        End If
     End If
     Detalle_de_SubCtas
  End If
End Sub

Private Sub DGCostos_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ID_Temp As Long
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyD Then
     ID_Temp = DGCostos.Columns(4)
     Titulo = "PREGUNTA DE ELIMINACION"
     Mensajes = "Esta seguro de eliminar este registro"
     If BoxMensaje = vbYes Then
        sSQL = "DELETE * " _
             & "FROM Trans_Presupuestos " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND ID = " & ID_Temp & " "
        Ejecutar_SQL_SP sSQL
        Detalle_de_SubCtas
     End If
  End If
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Codigo, Cuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND TC = 'CC' " _
       & "AND DG = 'G' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCProyecto, AdoCtas, sSQL, "Cuenta"
 
  sSQL = "SELECT Codigo, Detalle " _
       & "FROM Catalogo_SubCtas " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND TC IN ('G','CC') " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCSubModulos, AdoSubModulo, sSQL, "Detalle"
 
  If Bloquear_Control Then
     Toolbar1.buttons("Procesar").Enabled = False
     Toolbar1.buttons("Listar").Enabled = False
     Toolbar1.buttons("Imprimir").Enabled = False
  End If
  RatonNormal
End Sub

Private Sub Form_Load()
 'CentrarForm BalanceSubCtas
  ConectarAdodc AdoAux
  ConectarAdodc AdoCtas
  ConectarAdodc AdoCtas1
  ConectarAdodc AdoCostos
  ConectarAdodc AdoSubModulo
  
  DGCostos.Height = MDI_Y_Max - DGCostos.Top - 500
  DGCostos.width = MDI_X_Max - DGCostos.Left
  AdoCostos.Top = DGCostos.Top + DGCostos.Height + 30
  AdoCostos.width = MDI_X_Max - AdoCostos.Left
  
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   'MsgBox Button.key
    Select Case Button.key
      Case "Salir": Unload FCatalogo_Costos
      Case "Excel": DGCostos.Visible = False
                    GenerarDataTexto FCatalogo_Costos, AdoCostos
                    DGCostos.Visible = True
      Case "Todas": Detalle_de_SubCtas True
    End Select
End Sub

Public Sub Detalle_de_SubCtas(Optional TodasCtas As Boolean)
  sSQL = "SELECT TP.Cta, CC.Cuenta, CS.Detalle, CS.Codigo, TP.ID " _
       & "FROM Catalogo_Cuentas As CC, Catalogo_SubCtas As CS, Trans_Presupuestos As TP " _
       & "WHERE CC.Periodo = '" & Periodo_Contable & "' " _
       & "AND CC.Item = '" & NumEmpresa & "' "
  If TodasCtas Then sSQL = sSQL & "AND TP.Cta LIKE '" & CodSubCta & "%' " Else sSQL = sSQL & "AND TP.Cta = '" & SubCta & "' "
  sSQL = sSQL _
       & "AND TP.MesNo = 0 " _
       & "AND CC.Periodo = TP.Periodo " _
       & "AND CC.Item = TP.Item " _
       & "AND CC.Periodo = CS.Periodo " _
       & "AND CC.Item = CS.Item " _
       & "AND CC.Codigo = TP.Cta " _
       & "AND CS.Codigo = TP.Codigo " _
       & "ORDER BY TP.Cta, CS.Codigo "
  Select_Adodc_Grid DGCostos, AdoCostos, sSQL
  DCCtasProyecto.SetFocus
  RatonNormal
End Sub
