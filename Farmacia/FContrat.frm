VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FContratos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Contratos"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo DCClientes 
      DataSource      =   "AdoClientes"
      Height          =   315
      Left            =   1560
      TabIndex        =   26
      Top             =   435
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DCContrato 
      DataSource      =   "AdoContratos1"
      Height          =   315
      Left            =   120
      TabIndex        =   25
      Top             =   3150
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataList DLContrato 
      DataSource      =   "AdoContratos"
      Height          =   2595
      Left            =   120
      TabIndex        =   24
      Top             =   3480
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4577
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo DCArt 
      DataSource      =   "AdoArt"
      Height          =   315
      Left            =   1800
      TabIndex        =   23
      Top             =   735
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DCGrupoNo 
      DataSource      =   "AdoGrupoNo"
      Height          =   315
      Left            =   120
      TabIndex        =   22
      Top             =   420
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   240
      Top             =   5040
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Caption         =   "Clientes"
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
   Begin MSAdodcLib.Adodc AdoContratos1 
      Height          =   330
      Left            =   240
      Top             =   5640
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Caption         =   "Contratos1"
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
   Begin MSAdodcLib.Adodc AdoGrupoNo 
      Height          =   330
      Left            =   240
      Top             =   4680
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Caption         =   "GrupoNo"
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
   Begin MSAdodcLib.Adodc AdoContratos 
      Height          =   330
      Left            =   240
      Top             =   4320
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Caption         =   "Contratos"
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
   Begin MSAdodcLib.Adodc AdoArt 
      Height          =   330
      Left            =   240
      Top             =   3960
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Art"
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
   Begin MSAdodcLib.Adodc AdoPagos 
      Height          =   330
      Left            =   240
      Top             =   3600
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Caption         =   "Pagos"
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
   Begin VB.TextBox TextMontoME 
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
      Left            =   1680
      MaxLength       =   12
      TabIndex        =   18
      Top             =   2415
      Width           =   1695
   End
   Begin VB.TextBox TextMontoMN 
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
      Left            =   1680
      MaxLength       =   12
      TabIndex        =   14
      Top             =   2100
      Width           =   1695
   End
   Begin VB.TextBox TextInicial 
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
      Left            =   1680
      MaxLength       =   12
      TabIndex        =   10
      Top             =   1785
      Width           =   1695
   End
   Begin VB.TextBox TextIntervalo 
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
      Left            =   5250
      MaxLength       =   2
      TabIndex        =   16
      Top             =   2100
      Width           =   750
   End
   Begin VB.TextBox TextPagos 
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
      Left            =   5250
      MaxLength       =   3
      TabIndex        =   12
      Top             =   1785
      Width           =   750
   End
   Begin VB.TextBox TextConcepto 
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
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   8
      Top             =   1470
      Width           =   4320
   End
   Begin VB.TextBox TextContrato 
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
      Left            =   4305
      MaxLength       =   8
      TabIndex        =   6
      Top             =   1155
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
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
      Left            =   7350
      Picture         =   "FContrat.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1785
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
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
      Left            =   7350
      Picture         =   "FContrat.frx":0282
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   840
      Width           =   1065
   End
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   1680
      TabIndex        =   4
      Top             =   1155
      Width           =   1380
      _ExtentX        =   2434
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
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Intervalo de pagos:"
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
      TabIndex        =   15
      Top             =   2100
      Width           =   1905
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Numero de Pagos:"
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
      TabIndex        =   11
      Top             =   1785
      Width           =   1905
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nombre de Cliente: "
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
      Left            =   1575
      TabIndex        =   2
      Top             =   105
      Width           =   6840
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tipo de Prestamo"
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
      TabIndex        =   21
      Top             =   735
      Width           =   1695
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Contrato No."
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
      Left            =   3045
      TabIndex        =   5
      Top             =   1155
      Width           =   1275
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Montos M/E"
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
      TabIndex        =   17
      Top             =   2415
      Width           =   1590
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Montos M/N"
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
      TabIndex        =   13
      Top             =   2100
      Width           =   1590
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cuota Inicial"
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
      TabIndex        =   9
      Top             =   1785
      Width           =   1590
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " GRUPO No."
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
      Top             =   105
      Width           =   1485
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Concepto:"
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
      TabIndex        =   7
      Top             =   1470
      Width           =   1590
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CONTRATOS EMITIDOS"
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
      Top             =   2835
      Width           =   8310
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha de Inicio:"
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
      Top             =   1155
      Width           =   1590
   End
End
Attribute VB_Name = "FContratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
 Codigo = SinEspaciosDer(DCClientes.Text)
 Codigo1 = SinEspaciosDer(DCArt.Text)
 If Codigo = "" Then Codigo = Ninguno
 If Codigo1 = "" Then Codigo1 = Ninguno
 FechaValida MBoxFechaI, False
 TextoValido TextMontoMN, True
 TextoValido TextInicial, True
 TextoValido TextMontoME, True
 TextoValido TextContrato, True
 TextoValido TextConcepto, False
 If Val(TextPagos.Text) <= 0 Then TextPagos.Text = "1"
 If Val(TextIntervalo.Text) <= 0 Then TextIntervalo.Text = "1"
 If Val(TextPagos.Text) >= 1 Then
    sSQL = "SELECT * FROM Clientes "
    sSQL = sSQL & "WHERE Codigo = '" & Codigo & "' "
    SelectData AdoPagos, sSQL
    If AdoPagos.Recordset.RecordCount > 0 Then
       SelectData AdoPagos, "Contratos_Meses"
       sSQL = "DELETE * FROM Contratos_Meses "
       sSQL = sSQL & "WHERE Contrato_No = '" & TextContrato.Text & "' "
       DeleteData AdoPagos, sSQL
       With AdoPagos.Recordset
            MiFecha = MBoxFechaI.Text
           .AddNew
           .Fields("T") = Pendiente
           .Fields("Codigo_C") = Codigo
           .Fields("Codigo") = Codigo1
           .Fields("Fecha") = MiFecha
           .Fields("Contrato_No") = TextContrato.Text
           .Fields("Pago_No") = "00/" & Format(Val(TextPagos.Text), "00")
           .Fields("Concepto") = TextConcepto.Text
           .Fields("Monto_MN") = TextInicial.Text
           .Fields("Saldo") = TextInicial.Text
           .Fields("Monto_ME") = TextMontoME.Text
           .Fields("Cotizacion") = 0
            If Val(TextInicial.Text) > 0 Then .Update
            For I = 1 To Val(TextPagos.Text)
               .AddNew
               .Fields("T") = Pendiente
               .Fields("Codigo_C") = Codigo
               .Fields("Codigo") = Codigo1
               .Fields("Fecha") = MiFecha
               .Fields("Contrato_No") = TextContrato.Text
               .Fields("Pago_No") = Format(I, "00") & "/" & Format(Val(TextPagos.Text), "00")
               .Fields("Concepto") = TextConcepto.Text
               .Fields("Monto_MN") = TextMontoMN.Text
               .Fields("Saldo") = TextMontoMN.Text
               .Fields("Monto_ME") = TextMontoME.Text
               .Fields("Cotizacion") = 0
               .Update
                MiFecha = CLongFecha(CFechaLong(MiFecha) + Val(TextIntervalo.Text))
            Next I
       End With
       MiFecha = BuscarFecha(FechaSistema)
       sSQL = "SELECT Cliente & Space(90-Len(Cliente)) & Codigo As NomClientes "
       sSQL = sSQL & "FROM Clientes "
       sSQL = sSQL & "WHERE E = 'C' "
       sSQL = sSQL & "AND Grupo = '" & Grupo_No & "' "
       sSQL = sSQL & "ORDER BY Cliente "
       SelectDBCombo DCClientes, AdoClientes, sSQL, "NomClientes"
       
       sSQL = "SELECT Contrato_No & ' de: ' & Cliente As Contratos "
       sSQL = sSQL & "FROM Contratos_Meses As CM,Clientes As Cl "
       sSQL = sSQL & "WHERE CM.Codigo_C = Cl.Codigo "
       sSQL = sSQL & "GROUP BY Contrato_No & ' de: ' & Cliente "
       SelectDBCombo DCContrato, AdoContratos1, sSQL, "Contratos", False
   
       Codigo = SinEspaciosIzq(DCContrato.Text)
       sSQL = "SELECT Fecha & ' , de: ' & Cliente "
       sSQL = sSQL & " & ' => Monto M/N: ' & Format(Monto_MN,'#,##0.00') & ',"
       sSQL = sSQL & "Monto M/E: ' & Format(Monto_ME,'#,##0.00') "
       sSQL = sSQL & "As DetalleContrato "
       sSQL = sSQL & "FROM Contratos_Meses As CM,Clientes As Cl "
       sSQL = sSQL & "WHERE CM.Codigo_C = Cl.Codigo "
       sSQL = sSQL & "AND Contrato_No = '" & Codigo & "' "
       sSQL = sSQL & "AND T = '" & Normal & "' "
       sSQL = sSQL & "AND Cl.Grupo = '" & Grupo_No & "' "
       sSQL = sSQL & "ORDER BY Fecha,Pago_No "
       SelectDBList DLContrato, AdoContratos, sSQL, "DetalleContrato"
       SelectData AdoPagos, "Contratos_Meses"
       RatonNormal
       DCContrato.SetFocus
    End If
 End If
End Sub

Private Sub Command2_Click()
  Unload FContratos
End Sub



Private Sub DCContrato_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCContrato_LostFocus()
   Codigo = SinEspaciosIzq(DCContrato.Text)
   sSQL = "SELECT Fecha & ' , de: ' & Cliente "
   sSQL = sSQL & " & ' => Monto M/N: ' & Format(Monto_MN,'#,##0.00') & ', Monto M/E: ' & Format(Monto_ME,'#,##0.00') "
   sSQL = sSQL & "As DetalleContrato "
   sSQL = sSQL & "FROM Contratos_Meses As CM,Clientes As Cl "
   sSQL = sSQL & "WHERE CM.Codigo_C = Cl.Codigo "
   sSQL = sSQL & "AND Contrato_No = '" & Codigo & "' "
   sSQL = sSQL & "AND T = '" & Procesado & "' "
   sSQL = sSQL & "ORDER BY Fecha,Pago_No "
   SelectDBList DLContrato, AdoContratos, sSQL, "DetalleContrato"
End Sub

Private Sub DCGrupoNo_LostFocus()
   Grupo_No = DCGrupoNo.Text
   sSQL = "SELECT Contrato_No & ' de: ' & Cliente As Contratos "
   sSQL = sSQL & "FROM Contratos_Meses As CM,Clientes As Cl "
   sSQL = sSQL & "WHERE CM.Codigo_C = Cl.Codigo "
   sSQL = sSQL & "AND Cl.Grupo = '" & Grupo_No & "' "
   sSQL = sSQL & "GROUP BY Contrato_No & ' de: ' & Cliente "
   SelectDBCombo DCContrato, AdoContratos1, sSQL, "Contratos", False
   
   sSQL = "SELECT Cliente & Space(90-Len(Cliente)) & Codigo As NomClientes "
   sSQL = sSQL & "FROM Clientes "
   sSQL = sSQL & "WHERE T = 'C' "
   sSQL = sSQL & "AND Grupo = '" & Grupo_No & "' "
   sSQL = sSQL & "ORDER BY Cliente "
   SelectDBCombo DCClientes, AdoClientes, sSQL, "NomClientes"
   
End Sub

Private Sub DLContrato_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
     Codigo = SinEspaciosIzq(DCContrato.Text)
     Codigo1 = SinEspaciosIzq(DLContrato.Text)
     Mensajes = "Esta Seguro que desea eliminar" & Chr(13)
     Mensajes = Mensajes & "El Contrato No. " & Codigo & " de: " & Codigo1
     Titulo = "Formulario de Eliminacion."
     If BoxMensaje = 6 Then
        sSQL = "DELETE * FROM Contratos_Meses "
        sSQL = sSQL & "WHERE Contrato_No = '" & Codigo & "' "
        sSQL = sSQL & "AND Fecha = #" & BuscarFecha(Codigo1) & "# "
        ConectarAdoExecute sSQL
        MiFecha = BuscarFecha(FechaSistema)
        sSQL = "SELECT Contrato_No & ' de: ' & Cliente As Contratos "
        sSQL = sSQL & "FROM Contratos_Meses As CM,Clientes As Cl "
        sSQL = sSQL & "WHERE CM.Codigo_C = Cl.Codigo "
        sSQL = sSQL & "GROUP BY Contrato_No & ' de: ' & Cliente "
        SelectDBCombo DCContrato, AdoContratos1, sSQL, "Contratos", False
        DCContrato.SetFocus
     End If
  End If
End Sub

Private Sub Form_Activate()
   MiFecha = BuscarFecha(FechaSistema)
   sSQL = "SELECT Grupo "
   sSQL = sSQL & "FROM Clientes "
   sSQL = sSQL & "GROUP BY Grupo "
   SelectDBCombo DCGrupoNo, AdoGrupoNo, sSQL, "Grupo"
   Grupo_No = DCGrupoNo.Text
   
   sSQL = "SELECT Producto & '  =>  ' & Codigo_Inv As Art "
   sSQL = sSQL & "FROM Catalogo_Productos "
   sSQL = sSQL & "WHERE Codigo_Inv = '" & Ninguno & "' "
   sSQL = sSQL & "ORDER BY Producto "
   SelectDBCombo DCArt, AdoArt, sSQL, "Art"
   
   sSQL = "SELECT Cliente & Space(90-Len(Cliente)) & Codigo As NomClientes "
   sSQL = sSQL & "FROM Clientes "
   sSQL = sSQL & "WHERE T = 'C' "
   sSQL = sSQL & "AND Grupo = '" & Grupo_No & "' "
   sSQL = sSQL & "ORDER BY Cliente "
   SelectDBCombo DCClientes, AdoClientes, sSQL, "NomClientes"
   
   sSQL = "SELECT Contrato_No & ' de: ' & Cliente As Contratos "
   sSQL = sSQL & "FROM Contratos_Meses As CM,Clientes As Cl "
   sSQL = sSQL & "WHERE CM.Codigo_C = Cl.Codigo "
   sSQL = sSQL & "AND Cl.Grupo = '" & Grupo_No & "' "
   sSQL = sSQL & "GROUP BY Contrato_No & ' de: ' & Cliente "
   SelectDBCombo DCContrato, AdoContratos1, sSQL, "Contratos", False
   
   Codigo = SinEspaciosIzq(DCContrato.Text)
   sSQL = "SELECT Fecha & ' , de: ' & Cliente "
   sSQL = sSQL & " & ' => Monto M/N: ' & Format(Monto_MN,'#,##0.00') & ',"
   sSQL = sSQL & " Monto M/E: ' & Format(Monto_ME,'#,##0.00') "
   sSQL = sSQL & "As DetalleContrato "
   sSQL = sSQL & "FROM Contratos_Meses As CM,Clientes As Cl "
   sSQL = sSQL & "WHERE CM.Codigo_C = Cl.Codigo "
   sSQL = sSQL & "AND Contrato_No = '" & Codigo & "' "
   sSQL = sSQL & "AND T = '" & Normal & "' "
   sSQL = sSQL & "ORDER BY Fecha "
   SelectDBList DLContrato, AdoContratos, sSQL, "DetalleContrato"
   
   SelectData AdoPagos, "Contratos_Meses"
   RatonNormal
   DCContrato.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm FContratos
   ConectarAdodc AdoArt
   ConectarAdodc AdoPagos
   ConectarAdodc AdoGrupoNo
   ConectarAdodc AdoClientes
   ConectarAdodc AdoContratos
   ConectarAdodc AdoContratos1
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI, False
End Sub

Private Sub TextConcepto_GotFocus()
  TextConcepto.Text = ""
End Sub

Private Sub TextConcepto_LostFocus()
  TextoValido TextConcepto, False
End Sub

Private Sub TextContrato_GotFocus()
   TextContrato.Text = ""
End Sub

Private Sub TextContrato_LostFocus()
   TextoValido TextContrato, , True
   TextContrato.Text = String(8 - Len(TextContrato.Text), "0") & TextContrato.Text
End Sub

Private Sub TextInicial_LostFocus()
  TextoValido TextInicial, True
End Sub

Private Sub TextIntervalo_GotFocus()
  TextIntervalo.Text = ""
End Sub

Private Sub TextIntervalo_LostFocus()
   TextoValido TextIntervalo, True
   If Val(TextIntervalo.Text) <= 0 Then TextIntervalo.Text = "1"
End Sub

Private Sub TextMontoME_GotFocus()
   TextMontoME.Text = ""
End Sub

Private Sub TextMontoME_LostFocus()
  TextoValido TextMontoME, True
End Sub

Private Sub TextMontoMN_GotFocus()
  TextMontoMN.Text = ""
End Sub

Private Sub TextMontoMN_LostFocus()
  TextoValido TextMontoMN, True
End Sub

Private Sub TextPagos_GotFocus()
  TextPagos.Text = ""
End Sub

Private Sub TextPagos_LostFocus()
   TextoValido TextPagos, True
   If Val(TextPagos.Text) <= 0 Then TextPagos.Text = "1"
End Sub
