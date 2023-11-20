VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form AsientoNC 
   Caption         =   "NOTAS DE CREDITO"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4890
   ScaleWidth      =   11160
   WindowState     =   2  'Maximized
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   945
      TabIndex        =   1
      Top             =   105
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
   End
   Begin MSDataGridLib.DataGrid DGAsiento 
      Bindings        =   "FAsientNC.frx":0000
      Height          =   3270
      Left            =   105
      TabIndex        =   11
      Top             =   1050
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   5768
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
   Begin MSAdodcLib.Adodc AdoAsiento 
      Height          =   330
      Left            =   105
      Top             =   4410
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
      Caption         =   "Asiento"
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
   Begin VB.CommandButton Command7 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9030
      Picture         =   "FAsientNC.frx":0019
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10080
      Picture         =   "FAsientNC.frx":045B
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Consultar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7980
      Picture         =   "FAsientNC.frx":0D25
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   105
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoSubCta 
      Height          =   330
      Left            =   525
      Top             =   2205
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
      Caption         =   "SubCta"
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   525
      Top             =   2520
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   525
      Top             =   2835
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "Banco"
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
   Begin MSAdodcLib.Adodc AdoCaja 
      Height          =   330
      Left            =   525
      Top             =   3150
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "Caja"
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
   Begin MSAdodcLib.Adodc AdoAsientoB 
      Height          =   330
      Left            =   525
      Top             =   3465
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "AsientoB"
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
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   525
      Top             =   3780
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "Cliente"
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
   Begin MSAdodcLib.Adodc AdoTP 
      Height          =   330
      Left            =   525
      Top             =   1890
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "TP"
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
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "FAsientNC.frx":1167
      DataSource      =   "AdoCliente"
      Height          =   315
      Left            =   3675
      TabIndex        =   12
      Top             =   105
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Cliente"
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FACTURA No."
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
      Left            =   2310
      TabIndex        =   13
      Top             =   105
      Width           =   1380
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA"
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
      Width           =   855
   End
   Begin VB.Label LblConcepto 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      TabIndex        =   10
      Top             =   525
      Width           =   7785
   End
   Begin VB.Label LabelHaber 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   8400
      TabIndex        =   5
      Top             =   4410
      Width           =   1800
   End
   Begin VB.Label LabelDebe 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   6615
      TabIndex        =   6
      Top             =   4410
      Width           =   1800
   End
   Begin VB.Label LabelDiferencia 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   3465
      TabIndex        =   7
      Top             =   4410
      Width           =   1695
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Totales"
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
      Left            =   5565
      TabIndex        =   9
      Top             =   4410
      Width           =   1065
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Diferencia"
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
      Left            =   2310
      TabIndex        =   8
      Top             =   4410
      Width           =   1170
   End
End
Attribute VB_Name = "AsientoNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  IniciarAsientosDe DGAsiento, AdoAsiento
  LblConcepto.Caption = ""
  RatonReloj
  FechaValida MBFechaI
  FechaIni = BuscarFecha(MBFechaI.Text)
  SumaDebe = 0: SumaHaber = 0
  LblConcepto.Caption = "Intereses ganados del " & MBFechaI.Text
  SQL2 = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  SelectDataGrid DGAsiento, AdoAsiento, SQL2
  RatonReloj
  If Nuevo = False Then
  sSQL = "SELECT * " _
       & "FROM Ctas_Proceso " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "ORDER BY T_No "
  SelectData AdoAux, sSQL
  Total_Cheque = 0
  sSQL = "SELECT Cta,SUM(Interes) As Tot_Int " _
       & "FROM Trans_Prestamos " _
       & "WHERE Fecha <= #" & FechaIni & "# " _
       & "AND AC = False " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "GROUP BY Cta "
  SelectData AdoSubCta, sSQL
  With AdoSubCta.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Cta = .Fields("Cta")
          Total = Redondear(.Fields("Tot_Int"), 2)
          Codigo = Ninguno
          If AdoAux.Recordset.RecordCount > 0 Then
             TextoBusqueda = "Detalle like '" & Cta & "*' "
             AdoAux.Recordset.MoveFirst
             AdoAux.Recordset.Find (TextoBusqueda)
             If Not AdoAux.Recordset.EOF Then Codigo = AdoAux.Recordset.Fields("Codigo")
          End If
          InsertarAsientos AdoAsiento, Cta, 0, Total, 0
          InsertarAsientos AdoAsiento, Codigo, 0, 0, Total
         .MoveNext
       Loop
   End If
  End With
  Else
     
     NombreCliente = DCCliente.Text
     CodigoCliente = Ninguno
     If AdoCliente.Recordset.RecordCount > 0 Then
        TextoBusqueda = "Cliente like '" & NombreCliente & "*' "
        AdoCliente.Recordset.MoveFirst
        AdoCliente.Recordset.Find (TextoBusqueda)
        If Not AdoCliente.Recordset.EOF Then CodigoCliente = AdoCliente.Recordset.Fields("Codigo")
     End If
     sSQL = "SELECT Fecha,Credito_No,Cuota_No,Pagos " _
          & "FROM Trans_Prestamos " _
          & "WHERE Fecha <= #" & FechaIni & "# " _
          & "AND Cuenta_No = '" & CodigoCliente & "' " _
          & "AND Cta = '" & Cta & "' " _
          & "AND T = 'P' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "ORDER BY Fecha "
     SelectData AdoSubCta, sSQL
     With AdoSubCta.Recordset
      If .RecordCount > 0 Then
          Mifecha = .Fields("Fecha")
          Cuota_No = .Fields("Cuota_No")
          Credito_No = .Fields("Credito_No")
          Total = .Fields("Pagos")
          InsertarAsientos AdoAsiento, Cta, 0, 0, Total
          InsertarAsientos AdoAsiento, Codigo1, 0, Total - Saldo, 0
          InsertarAsientos AdoAsiento, Codigo2, 0, Saldo, 0
          LblConcepto.Caption = "Abono No. " & Cuota_No _
                              & ", del Sr(a). " & NombreCliente _
                              & ", por USD " & Format$(Total, "#,##0.00")
      End If
     End With
  End If
  SumaDebe = 0: SumaHaber = 0
  With AdoAsiento.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          SumaDebe = SumaDebe + .Fields("DEBE")
          SumaHaber = SumaHaber + .Fields("HABER")
         .MoveNext
       Loop
   End If
  End With
  LabelDebe.Caption = Format$(SumaDebe, "#,##0.00")
  LabelHaber.Caption = Format$(SumaHaber, "#,##0.00")
  LabelDiferencia.Caption = Format$(SumaDebe - SumaHaber, "#,##0.00")
  DGAsiento.Visible = True
  AsientoNC.Caption = "INTERESES GANADOS"
  RatonNormal
End Sub

Private Sub Command3_Click()
  Unload Me
End Sub

Private Sub Command7_Click()
  FechaIni = BuscarFecha(MBFechaI.Text)
  SumaDebe = 0: SumaHaber = 0
  DGAsiento.Visible = False
  With AdoAsiento.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          SumaDebe = SumaDebe + .Fields("DEBE")
          SumaHaber = SumaHaber + .Fields("HABER")
         .MoveNext
       Loop
      .MoveFirst
       If Redondear(SumaDebe - SumaHaber, 2) = 0 Then
          RatonReloj
          Co.T = Normal
          If Nuevo Then
             Co.TP = CompIngreso
             Co.Numero = ReadSetDataNum("Ingresos", True, True)
             Co.CodigoB = CodigoCliente
          Else
             Co.TP = CompDiario
             Co.Numero = ReadSetDataNum("Diario", True, True)
             Co.CodigoB = Ninguno
             Co.Efectivo = 0
          End If
          Co.Fecha = MBFechaI.Text
          Co.Monto_Total = SumaDebe
          Co.Item = NumEmpresa
          Co.Usuario = CodigoUsuario
          Co.Concepto = LblConcepto.Caption
          Co.T_No = Trans_No
          GrabarComprobante Co
          If Nuevo Then
             SQL2 = "UPDATE Trans_Prestamos " _
                  & "SET AC = True,T = 'C',Fecha_C = #" & BuscarFecha(Mifecha) & "# " _
                  & "WHERE Credito_No = '" & Credito_No & "' " _
                  & "AND Cuota_No = " & Cuota_No & " " _
                  & "AND Cuenta_No = '" & CodigoCliente & "' " _
                  & "AND Item = '" & NumEmpresa & "' "
          Else
             SQL2 = "UPDATE Trans_Prestamos " _
                  & "SET AC = True " _
                  & "WHERE Fecha <= #" & FechaIni & "# " _
                  & "AND AC = False " _
                  & "AND Item = '" & NumEmpresa & "' "
          End If
          ConectarAdoExecute SQL2
          ImprimirComprobantesDe False, Co
          Unload Me
       Else
         MsgBox "Las Transacciones no cuadran"
         DGAsiento.Visible = True
       End If
   Else
      MsgBox "No Existen Datos"
      DGAsiento.Visible = True
   End If
  End With
End Sub

Private Sub DCCliente_LostFocus()
   FechaValida MBFechaI
   FechaIni = BuscarFecha(MBFechaI.Text)
   IniciarAsientosDe DGAsiento, AdoAsiento
   Factura_No = Val(DCCliente.Text)
   SQL1 = "SELECT DF.*,CP.IVA,CP.Cta_Ventas " _
        & "FROM Detalle_Factura As DF,Catalogo_Productos As CP " _
        & "WHERE DF.Factura = " & Factura_No & " " _
        & "AND DF.Item = '" & NumEmpresa & "' " _
        & "AND DF.TC <> 'C' " _
        & "AND DF.TC <> 'P' " _
        & "ORDER BY Codigo "
   SelectAdodc AdoAux, SQL1
'''   With AdoAux.Recordset
'''    If .RecordCount > 0 Then
'''        MsgBox Factura_No & vbCrLf & .RecordCount
'''    End If
'''   End With
End Sub

Private Sub Form_Activate()
  Trans_No = 68
  FechaValida MBFechaI
  FechaIni = BuscarFecha(MBFechaI.Text)
  IniciarAsientosDe DGAsiento, AdoAsiento
  SQL1 = "SELECT F.TC,F.Factura,F.CodigoC,F.Fecha,F.Fecha_V,F.Saldo_MN,F.Cta_CxP,F.Nota," _
       & "F.Observacion,C.Cliente,C.Direccion,C.CI_RUC,C.Telefono " _
       & "FROM Facturas As F,Clientes As C " _
       & "WHERE F.T = '" & Pendiente & "' " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND F.TC <> 'C' " _
       & "AND F.TC <> 'P' " _
       & "AND F.CodigoC = C.Codigo " _
       & "ORDER BY F.Factura "
  SelectDBCombo DCCliente, AdoCliente, SQL1, "Factura"
  Mifecha = BuscarFecha(FechaTexto)
  RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoTP
  ConectarAdodc AdoAux
  ConectarAdodc AdoCaja
  ConectarAdodc AdoBanco
  ConectarAdodc AdoSubCta
  ConectarAdodc AdoCliente
  ConectarAdodc AdoAsiento
  ConectarAdodc AdoAsientoB
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

