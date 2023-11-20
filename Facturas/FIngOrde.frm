VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FIngOrden 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INGRESO DE DETALLES DE LA ORDEN DE PRODUCCION"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrmOrden 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3900
      Left            =   105
      TabIndex        =   3
      Top             =   525
      Width           =   9255
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000C0C0&
         Caption         =   "&Aceptar"
         Height          =   330
         Left            =   7140
         TabIndex        =   16
         Top             =   3465
         Width           =   960
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0000C0C0&
         Caption         =   "&Cancelar"
         Height          =   330
         Left            =   8190
         TabIndex        =   17
         Top             =   3465
         Width           =   960
      End
      Begin VB.CheckBox CheqSeco 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Seco"
         Height          =   330
         Left            =   5775
         TabIndex        =   8
         Top             =   525
         Width           =   750
      End
      Begin VB.CheckBox CheqAgua 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Agua"
         Height          =   330
         Left            =   4620
         TabIndex        =   7
         Top             =   525
         Width           =   750
      End
      Begin VB.TextBox TxtCantO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   6825
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "FIngOrde.frx":0000
         Top             =   525
         Width           =   960
      End
      Begin VB.TextBox TxtPVPO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   7770
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "FIngOrde.frx":0002
         Top             =   525
         Width           =   1380
      End
      Begin MSDataGridLib.DataGrid DGDetOrden 
         Bindings        =   "FIngOrde.frx":0007
         Height          =   2430
         Left            =   105
         TabIndex        =   13
         Top             =   945
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   4286
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
      Begin MSAdodcLib.Adodc AdoDetOrden 
         Height          =   330
         Left            =   105
         Top             =   3045
         Visible         =   0   'False
         Width           =   3090
         _ExtentX        =   5450
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
         Caption         =   "DetOrden"
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
      Begin MSAdodcLib.Adodc AdoDetArt 
         Height          =   330
         Left            =   3255
         Top             =   3045
         Visible         =   0   'False
         Width           =   3300
         _ExtentX        =   5821
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
         Caption         =   "DetArt"
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
      Begin MSDataListLib.DataCombo DCDetArt 
         Bindings        =   "FIngOrde.frx":0021
         DataSource      =   "AdoDetArt"
         Height          =   315
         Left            =   105
         TabIndex        =   5
         ToolTipText     =   "<F10> Insertar Orden de Pedidos"
         Top             =   525
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   8454143
         ForeColor       =   8388608
         Text            =   "Producto"
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
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
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
         Left            =   6825
         TabIndex        =   9
         Top             =   210
         Width           =   960
      End
      Begin VB.Label Label32 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PRODUCTO"
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
         Top             =   210
         Width           =   4110
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Left            =   5460
         TabIndex        =   15
         Top             =   3465
         Width           =   1590
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL DE LA ORDEN"
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
         Left            =   3360
         TabIndex        =   14
         Top             =   3465
         Width           =   2115
      End
      Begin VB.Label Label37 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO SERVICIO"
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
         Left            =   4200
         TabIndex        =   6
         Top             =   210
         Width           =   2640
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio Unitario"
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
         Left            =   7770
         TabIndex        =   11
         Top             =   210
         Width           =   1380
      End
   End
   Begin MSDataListLib.DataCombo DCOrden 
      Bindings        =   "FIngOrde.frx":0039
      DataSource      =   "AdoOrden"
      Height          =   315
      Left            =   1155
      TabIndex        =   1
      ToolTipText     =   "<Ctrl+A>: Activa una Orden Procesada"
      Top             =   105
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Clientes"
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
   Begin MSAdodcLib.Adodc AdoOrden 
      Height          =   330
      Left            =   210
      Top             =   735
      Visible         =   0   'False
      Width           =   3300
      _ExtentX        =   5821
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
      Caption         =   "Orden"
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
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2835
      TabIndex        =   2
      Top             =   105
      Width           =   6525
   End
   Begin VB.Label Label25 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Orden No."
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
      Top             =   105
      Width           =   1065
   End
End
Attribute VB_Name = "FIngOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OrdenP As Long
Dim CodArtOrden As String
Dim Ln_No_O As Long

Private Sub Command4_Click()
    Mensajes = "SEGURO DE GRABAR LA ORDEN"
    If BoxMensaje = vbYes Then
        Ln_No_O = 0
        sSQL = "SELECT * " _
             & "FROM Asiento_TK " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' "
        SelectAdodc AdoDetOrden, sSQL
        With AdoDetOrden.Recordset
         If .RecordCount > 0 Then
             Do While Not .EOF
                SetAdoAddNew "Trans_Ticket"
                SetAdoFields "TC", "OP"
                SetAdoFields "Codigo_Inv", .Fields("CODIGO")
                SetAdoFields "Producto", .Fields("PRODUCTO")
                SetAdoFields "Cantidad", .Fields("CANT")
                SetAdoFields "Precio", .Fields("PRECIO")
                SetAdoFields "Total", .Fields("TOTAL")
                SetAdoFields "CodigoC", CodigoCliente
                SetAdoFields "Ticket", OrdenP
                SetAdoFields "Fecha", Mifecha
                SetAdoFields "A", .Fields("A")
                SetAdoFields "L", .Fields("L")
                SetAdoFields "S", .Fields("S")
                SetAdoFields "CodigoU", CodigoUsuario
                SetAdoFields "Item", NumEmpresa
                SetAdoFields "D_No", CByte(Ln_No_O)
                SetAdoUpdate
                Ln_No_O = Ln_No_O + 1
               .MoveNext
             Loop
             sSQL = "DELETE * " _
                  & "FROM Asiento_TK " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND CodigoU = '" & CodigoUsuario & "' "
             ConectarAdoExecute sSQL
             sSQL = "SELECT * " _
                  & "FROM Asiento_TK " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND CodigoU = '" & CodigoUsuario & "' "
             SQLDec = "PRECIO " & CStr(Dec_PVP) & "|TOTAL 4|."
             SelectDataGrid DGDetOrden, AdoDetOrden, sSQL, SQLDec
             DCOrden.SetFocus
         End If
        End With
    End If
End Sub

Private Sub Command5_Click()
  Unload FIngOrden
End Sub

Private Sub DCDetArt_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then Command4.SetFocus
End Sub

Private Sub DCDetArt_LostFocus()
Dim CodArt As String
  CodArt = DCDetArt
  CodArtOrden = Ninguno
  If Len(CodArt) <= 1 Then CodArt = Ninguno
  TxtPVPO = "0"
  TxtCantO = "1"
  With AdoDetArt.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Producto Like '" & CodArt & "' ")
       If Not .EOF Then
          TxtPVPO = .Fields("PVP")
          CodArtOrden = .Fields("Codigo_Inv")
          CheqAgua.value = 0
          CheqSeco.value = 0
       End If
   End If
  End With
End Sub

Private Sub DCOrden_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCOrden_LostFocus()
   OrdenP = Val(DCOrden)
   With AdoOrden.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("Factura = " & OrdenP & " ")
        If Not .EOF Then
           CodigoCliente = .Fields("CodigoC")
           Label1.Caption = .Fields("Cliente")
           Mifecha = .Fields("Fecha")
        End If
    End If
   End With
End Sub

Private Sub Form_Activate()
   sSQL = "SELECT * " _
        & "FROM Catalogo_Productos " _
        & "WHERE TC = 'P' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND INV <> " & Val(adFalse) & " " _
        & "ORDER BY Producto,Codigo_Inv "
   SelectDBCombo DCDetArt, AdoDetArt, sSQL, "Producto"
   
   sSQL = "SELECT OP.*,C.Cliente,C.Grupo,C.CI_RUC,C.TD " _
        & "FROM Facturas AS OP,Clientes As C " _
        & "WHERE OP.Item = '" & NumEmpresa & "' " _
        & "AND OP.Periodo = '" & Periodo_Contable & "' " _
        & "AND OP.TC = 'OP' " _
        & "AND OP.T <> 'A' " _
        & "AND OP.CodigoC = C.Codigo " _
        & "ORDER BY OP.Factura "
   SelectDBCombo DCOrden, AdoOrden, sSQL, "Factura"
   
   sSQL = "DELETE * " _
        & "FROM Asiento_TK " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   ConectarAdoExecute sSQL
   
   sSQL = "SELECT * " _
        & "FROM Asiento_TK " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   SQLDec = "PRECIO " & CStr(Dec_PVP) & "|TOTAL 4|."
   SelectDataGrid DGDetOrden, AdoDetOrden, sSQL, SQLDec
   DCOrden.Text = "0"
   RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm FIngOrden
   
   ConectarAdodc AdoOrden
   ConectarAdodc AdoDetArt
   ConectarAdodc AdoDetOrden
   
End Sub

Private Sub TxtPVPO_GotFocus()
  MarcarTexto TxtPVPO
End Sub

Private Sub TxtPVPO_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPVPO_LostFocus()
Dim Total_OP As Currency
  If Len(CodArtOrden) > 1 And Val(TxtPVPO) > 0 And Val(TxtCantO) > 0 Then
     Label33.Caption = "0"
     SetAdoAddNew "Asiento_TK"
     SetAdoFields "CODIGO", CodArtOrden
     SetAdoFields "PRODUCTO", DCDetArt
     SetAdoFields "CANT", CCur(TxtCantO)
     SetAdoFields "PRECIO", CCur(TxtPVPO)
     SetAdoFields "TOTAL", CCur(TxtCantO) * CCur(TxtPVPO)
     SetAdoFields "Item", NumEmpresa
     SetAdoFields "Numero", OrdenP
     SetAdoFields "CODIGO_C", CodigoCliente
     SetAdoFields "CodigoU", CodigoUsuario
     If CheqAgua.value = 1 Then SetAdoFields "A", adTrue
     If CheqSeco.value = 1 Then SetAdoFields "S", adTrue
     SetAdoFields "A_No", CByte(Ln_No_O)
     SetAdoUpdate
     
     Total_OP = 0
     sSQL = "SELECT * " _
          & "FROM Asiento_TK " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     SQLDec = "PRECIO " & CStr(Dec_PVP) & "|TOTAL 4|."
     SelectDataGrid DGDetOrden, AdoDetOrden, sSQL, SQLDec
     If AdoDetOrden.Recordset.RecordCount > 0 Then
        Do While Not AdoDetOrden.Recordset.EOF
           Total_OP = Total_OP + AdoDetOrden.Recordset.Fields("TOTAL")
           AdoDetOrden.Recordset.MoveNext
        Loop
     End If
     Label33.Caption = Format$(Total_OP, "#,##0.00")
     Ln_No_O = Ln_No_O + 1
  End If
  DCDetArt.SetFocus
End Sub
