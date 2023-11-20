VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form ChequeEfectivo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flujo de Caja"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DGCheques 
      Bindings        =   "CheqEfec.frx":0000
      Height          =   2640
      Left            =   105
      TabIndex        =   17
      ToolTipText     =   "<Ins> Efectivizar Cheques, <Del/Sup> Protestar Cheque, <F3> Cambio de Días"
      Top             =   1680
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   4657
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
      Caption         =   "CHEQUES EFECTIVIZADOS"
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
   Begin MSDataListLib.DataList DLBanco 
      Bindings        =   "CheqEfec.frx":0019
      DataSource      =   "AdoBanco"
      Height          =   1035
      Left            =   105
      TabIndex        =   16
      Top             =   525
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   1826
      _Version        =   393216
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
   Begin VB.TextBox TextCosto 
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
      Left            =   7665
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "CheqEfec.frx":0030
      Top             =   630
      Width           =   1590
   End
   Begin VB.TextBox TextComision 
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
      Left            =   7665
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "CheqEfec.frx":0032
      Top             =   1155
      Width           =   1590
   End
   Begin VB.CommandButton Command5 
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
      Height          =   1065
      Left            =   9345
      Picture         =   "CheqEfec.frx":0034
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1785
      Width           =   1065
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   9345
      Picture         =   "CheqEfec.frx":02B6
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   630
      Width           =   1065
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   1050
      TabIndex        =   1
      Top             =   105
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
   Begin MSAdodcLib.Adodc AdoAsientos 
      Height          =   330
      Left            =   210
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
      Caption         =   "Asientos"
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
      Left            =   210
      Top             =   2100
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
   Begin MSAdodcLib.Adodc AdoProtesto 
      Height          =   330
      Left            =   210
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
      Caption         =   "Protesto"
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
      Left            =   210
      Top             =   1785
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
   Begin MSAdodcLib.Adodc AdoCheques 
      Height          =   330
      Left            =   210
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
      Caption         =   "Cheques"
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
   Begin MSDataGridLib.DataGrid DGProtesto 
      Bindings        =   "CheqEfec.frx":0B80
      Height          =   1695
      Left            =   105
      TabIndex        =   18
      Top             =   4620
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   2990
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
      Caption         =   "CHEQUES PROTESTADOS"
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
   Begin VB.Label LabelIngCheqME 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   6720
      TabIndex        =   11
      Top             =   4305
      Width           =   1800
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " INGRESOS CHEQUE M/E"
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
      TabIndex        =   12
      Top             =   4305
      Width           =   2430
   End
   Begin VB.Label LabelIngCheqMN 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   2520
      TabIndex        =   9
      Top             =   4305
      Width           =   1800
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " COMISION COOP."
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
      Left            =   5670
      TabIndex        =   4
      Top             =   1155
      Width           =   2010
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " COMISION BANCO"
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
      Left            =   5670
      TabIndex        =   2
      Top             =   630
      Width           =   2010
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   2520
      TabIndex        =   8
      Top             =   6300
      Width           =   1800
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " INGRESOS CHEQUE M/N"
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
      Top             =   6300
      Width           =   2430
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   6720
      TabIndex        =   15
      Top             =   6300
      Width           =   1800
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " INGRESOS CHEQUE M/E"
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
      TabIndex        =   14
      Top             =   6300
      Width           =   2430
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " INGRESOS CHEQUE M/N"
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
      TabIndex        =   10
      Top             =   4305
      Width           =   2430
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA:"
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
      Width           =   960
   End
End
Attribute VB_Name = "ChequeEfectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub InsertarMontos(MBFecha As MaskEdBox, _
                          DtaCta As Adodc, _
                          CuentaNo As String, _
                          TDebe As Currency, _
                          THaber As Currency)
TiempoTexto = Format(Time, FormatoTimes)
If (TDebe + THaber) <> 0 Then
  SaldoDisp = 0: SaldoCont = 0: NumeroLineas = 0
  sSQL = "SELECT TOP 1 * " _
       & "FROM Trans_Libretas " _
       & "WHERE Cuenta_No = '" & CuentaNo & "' " _
       & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
  SelectAdodc DtaCta, sSQL
  With DtaCta.Recordset
       If .RecordCount > 0 Then
           SaldoDisp = .Fields("Saldo_Disp")
           SaldoCont = .Fields("Saldo_Cont")
           NumeroLineas = .Fields("ID")
           ID_Trans = .Fields("IDT")
       End If
      .AddNew
      .Fields("Fecha") = MBFecha.Text
      .Fields("Cuenta_No") = CuentaNo
      .Fields("TP") = TipoProc
      .Fields("Debitos") = TDebe
      .Fields("Creditos") = THaber
       If TipoProc = "CHPR" Then
         .Fields("Saldo_Cont") = SaldoCont + THaber - TDebe
         .Fields("Saldo_Disp") = SaldoDisp
       Else
         .Fields("Saldo_Cont") = SaldoCont + THaber - TDebe
         .Fields("Saldo_Disp") = SaldoDisp + THaber - TDebe
       End If
      .Fields("T") = Normal
      .Fields("CodigoU") = CodigoUsuario
      .Fields("IDT") = ID_Trans + 1
      .Fields("Hora") = TiempoTexto
      .Fields("Item") = NumEmpresa
      .Fields("ME") = False
      .Fields("Cheque") = Ninguno
       SetUpdate DtaCta
  End With
End If
End Sub

Public Sub InsertarCheques(MBFecha As String, _
                           DtaCta As Adodc, _
                           CuentaNo As String, _
                           TDebe As Currency, _
                           THaber As Currency, _
                           NoCheque As String, _
                           NomBanco As String)
  TiempoTexto = Format(Time, FormatoTimes)
  Saldo = 0: SaldoActual = 0
  sSQL = "SELECT TOP 1 * " _
       & "FROM Trans_Libretas " _
       & "WHERE Cuenta_No = '" & CuentaNo & "' " _
       & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
  SelectAdodc DtaCta, sSQL
  With DtaCta.Recordset
       If .RecordCount > 0 Then
           SaldoDisp = .Fields("Saldo_Disp")
           SaldoCont = .Fields("Saldo_Cont")
       End If
      .AddNew
      .Fields("Fecha") = MBFecha
      .Fields("Cuenta_No") = CuentaNo
      .Fields("TP") = MidStrg(TipoProc, 1, 4)
      .Fields("Debitos") = TDebe
      .Fields("Creditos") = THaber
      .Fields("Saldo_Cont") = SaldoCont
      .Fields("Saldo_Disp") = SaldoDisp + THaber - TDebe
      .Fields("T") = Normal
      .Fields("CodigoU") = CodigoUsuario
      .Fields("Item") = NumEmpresa
      .Fields("Hora") = TiempoTexto
      .Fields("ME") = False
      .Fields("Cheque") = NoCheque
      .Update
  End With
End Sub

Public Sub SumarCheques(DtaCajaCheq As Adodc, _
                        DtaCajaEfec As Adodc)
  RatonReloj
  Saldo = 0: Total = 0
  Debe = 0: Haber = 0
  Debe_ME = 0: Haber_ME = 0
  With DtaCajaCheq.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          If .Fields("ME") Then
              Saldo = Saldo + .Fields("Debitos")
          Else
              Total = Total + .Fields("Debitos")
          End If
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  With DtaCajaEfec.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          If .Fields("ME") Then
              Debe_ME = Debe_ME + .Fields("Debitos")
              Haber_ME = Haber_ME + .Fields("Creditos")
          Else
              Debe = Debe + .Fields("Debitos")
              Haber = Haber + .Fields("Creditos")
          End If
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  LabelIngCheqMN.Caption = Format(Total, "#,##0.00")
  LabelIngCheqME.Caption = Format(Saldo, "#,##0.00")
  RatonNormal
End Sub

Private Sub Command4_Click()
  SQLMsg1 = "REPORTE DE FLUJO DE CAJA"
  'ImprimirFlujoCaja adoFlujoCaja, True, 1, 8, True
End Sub

Private Sub Command5_Click()
  Unload ChequeEfectivo
End Sub

Private Sub DGCheques_KeyDown(KeyCode As Integer, Shift As Integer)
   Trans_No = 52
   If KeyCode = vbKeyF12 Then
      TipoProc = "EFCH"
      Cadena = ChequeEfectivo.Caption
      With AdoCheques.Recordset
           FechaTexto = DGCheques.Columns(2)
           CuentaBanco = DGCheques.Columns(3)
           NombreBanco = DGCheques.Columns(4)
           NumCheque = DGCheques.Columns(5)
           Total = CCur(DGCheques.Columns(6))
           ChequeEfectivo.Caption = FechaTexto & ", " & CuentaBanco & ", " & Format(Total, "#,##0.00")
           Mensajes = "SEGURO DE EFECTIVIZAR EL CHEQUE No.: " & NumCheque & vbCrLf _
                    & "DEL BANCO: " & NombreBanco & vbCrLf _
                    & "CUENTA No. " & CuentaBanco & vbCrLf _
                    & "FECHA " & FechaTexto & vbCrLf _
                    & "POR UN VALOR DE USD " & Total & vbCrLf
           Titulo = "Pregunta de grabación"
           If BoxMensaje = vbYes Then
              sSQL = "SELECT TOP 1 * " _
                   & "FROM Trans_Libretas " _
                   & "WHERE Cuenta_No = '" & CuentaBanco & "' " _
                   & "AND Item = '" & NumEmpresa & "' " _
                   & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
              SelectAdodc AdoCtas, sSQL
              If AdoCtas.Recordset.RecordCount > 0 Then
                 SaldoDisp = AdoCtas.Recordset.Fields("Saldo_Disp")
                 SaldoCont = AdoCtas.Recordset.Fields("Saldo_Cont")
                 
                 AdoCtas.Recordset.Fields("Saldo_Disp") = Round(SaldoDisp + Total, 2)
                 AdoCtas.Recordset.Update
                 MsgBox "Proceso Realizado"
              End If
           End If
      End With
      ChequeEfectivo.Caption = Cadena
      DGCheques.Visible = True
   End If
   If KeyCode = vbKeyInsert Or KeyCode = vbKeyDelete Or KeyCode = vbKeyF3 Then
      NoDias = CInt(Val(DGCheques.Columns(0)))
      FechaTexto = DGCheques.Columns(2)
      CuentaBanco = DGCheques.Columns(3)
      NombreBanco = DGCheques.Columns(4)
      NumCheque = DGCheques.Columns(5)
      Total = CCur(DGCheques.Columns(6))
      Mensajes = "Fecha Cheque: " & FechaTexto & vbCrLf _
               & "Cuenta No.    " & CuentaBanco & vbCrLf _
               & "Nombre Banco: " & NombreBanco & vbCrLf _
               & "Cheque No.    " & NumCheque & vbCrLf _
               & "Valor Cheque: " & Format(Total, "#,##0.00") & vbCrLf
      ValorDH = Total
      If KeyCode = vbKeyInsert Then
         DGCheques.Visible = False
         TipoProc = "EFCH"
         Mensajes = "SEGURO DE EFECTIVIZAR LOS CHEQUES: " & Chr(13)
         Titulo = "Pregunta de grabación"
         If BoxMensaje = vbYes Then
            With AdoCheques.Recordset
             .MoveFirst
              Do While Not .EOF
                 Cadena = ChequeEfectivo.Caption
                 ChequeEfectivo.Caption = FechaTexto & ", " & CuentaBanco & ", " & Format(Total, "#,##0.00")
                 FechaTexto = .Fields("Fecha")
                 NoDias = CFechaLong(FechaSistema) - CFechaLong(FechaTexto)
                 
                 CuentaBanco = .Fields("Cuenta_No")
                 NombreBanco = .Fields("Banco")
                 NumCheque = .Fields("Cheque")
                 Total = .Fields("Creditos")
                 'MsgBox CuentaBanco & vbCrLf & NoDias
                 If NoDias >= 2 Then
                    AdoCheques.Recordset.Fields("T") = Normal
                    AdoCheques.Recordset.Update
                    sSQL = "SELECT TOP 1 * " _
                         & "FROM Trans_Libretas " _
                         & "WHERE Cuenta_No = '" & CuentaBanco & "' " _
                         & "AND Item = '" & NumEmpresa & "' " _
                         & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
                    SelectAdodc AdoCtas, sSQL
                    If AdoCtas.Recordset.RecordCount > 0 Then
                       SaldoDisp = AdoCtas.Recordset.Fields("Saldo_Disp")
                       SaldoCont = AdoCtas.Recordset.Fields("Saldo_Cont")
                       NumeroLineas = AdoCtas.Recordset.Fields("ID")
                       ID_Trans = AdoCtas.Recordset.Fields("IDT")
                       TiempoTexto = Format(Time, FormatoTimes)
                       If NumeroLineas <= 0 Then NumeroLineas = 1
                       AdoCtas.Recordset.AddNew
                       AdoCtas.Recordset.Fields("Fecha") = FechaSistema
                       AdoCtas.Recordset.Fields("Cuenta_No") = CuentaBanco
                       AdoCtas.Recordset.Fields("TP") = TipoProc
                       AdoCtas.Recordset.Fields("Debitos") = 0
                       AdoCtas.Recordset.Fields("Creditos") = Total
                       AdoCtas.Recordset.Fields("T") = Normal
                       AdoCtas.Recordset.Fields("CodigoU") = CodigoUsuario
                       AdoCtas.Recordset.Fields("IDT") = ID_Trans + 1
                       AdoCtas.Recordset.Fields("Hora") = TiempoTexto
                       AdoCtas.Recordset.Fields("Item") = NumEmpresa
                       AdoCtas.Recordset.Fields("ME") = False
                       AdoCtas.Recordset.Fields("Cheque") = Ninguno
                       'MsgBox "....."
                       AdoCtas.Recordset.Fields("Saldo_Cont") = SaldoDisp + Total
                       AdoCtas.Recordset.Fields("Saldo_Disp") = SaldoDisp + Total
                       SetUpdate AdoCtas
                    End If
                 End If
                .MoveNext
              Loop
              Unload ChequeEfectivo
              sSQL = "UPDATE Trans_Dep_Chq " _
                   & "SET T = 'N' " _
                   & "WHERE Dias = 0 " _
                   & "AND T = 'P' "
              'ConectarAdoExecute sSQL
            End With
            ChequeEfectivo.Caption = Cadena
            DGCheques.Visible = True
         End If
         DGCheques.Visible = True
      End If
      If KeyCode = vbKeyDelete Then
         Mensajes = "SEGURO DE PROTESTAR CHEQUE: " & Chr(13) & Chr(13) & Mensajes
         Titulo = "Pregunta de grabación"
         If BoxMensaje = vbYes Then
            Saldo = Round(Val(TextCosto.Text), 2)
            Total_Comision = Round(Val(TextComision.Text), 2)
            sSQL = "UPDATE Trans_Dep_Chq SET T = 'A', " _
                 & "Cta = '" & SinEspaciosIzq(DLBanco.Text) & "' " _
                 & "WHERE Fecha = #" & BuscarFecha(FechaTexto) & "# " _
                 & "AND Cuenta_No = '" & CuentaBanco & "' " _
                 & "AND Cheque = '" & NumCheque & "' " _
                 & "AND Banco = '" & NombreBanco & "' " _
                 & "AND Valor = " & Total & " "
            'ConectarAdoExecute sSQL
            TipoProc = "CHPR"
            InsertarMontos MBoxFecha, AdoCtas, CuentaBanco, Saldo + Total_Comision1, 0
            TipoProc = "NDCH"
            InsertarMontos MBoxFecha, AdoCtas, CuentaBanco, Total, 0
           'Procesar Diario
            NumComp = ReadSetDataNum("Diario", True, True)
            BorrarAsientos True
            sSQL = "SELECT * " _
                 & "FROM Asiento " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND CodigoU = '" & CodigoUsuario & "' " _
                 & "AND T_No = " & Trans_No & " "
            SelectAdodc AdoAsientos, sSQL
            ValorDH = Total + Saldo + Total_Comision
            InsertarAsientos AdoAsientos, Cta_Libretas, 0, ValorDH, 0
            InsertarAsientos AdoAsientos, SinEspaciosIzq(DLBanco.Text), 0, 0, Total + Saldo
            InsertarAsientos AdoAsientos, Cta_ComisionB, 0, 0, Total_Comision
            Co.T = Normal
            Co.TP = CompDiario
            Co.Fecha = MBoxFecha.Text
            Co.Numero = NumComp
            Co.Concepto = "(" & NumEmpresa & ")" _
                        & "ND " & CuentaBanco & " CHPR No. " & NumCheque _
                        & " del Banco " & NombreBanco _
                        & ", Depositado el " & FechaTexto _
                        & " por " & Format(Total, "#,##0.00")
            Co.CodigoB = Ninguno
            Co.Efectivo = Total
            Co.Monto_Total = ValorDH
            Co.T_No = Trans_No
            Co.Item = NumEmpresa
            Co.Usuario = CodigoUsuario
            GrabarComprobante Co
            RatonNormal
            ImprimirComprobantesDe False, Co
         End If
      End If
      If KeyCode = vbKeyF3 Then
         Titulo = "Pregunta de Actualizacion"
         Saldo = 4
         Mensajes = Mensajes & vbCrLf & vbCrLf & "Días de efectivización:"
         JE = Val(InputBox(Mensajes, Titulo))
         If JE <= 0 Then JE = 0
         Titulo = "Pregunta de grabación"
         Mensajes = "SEGURO DE GRABAR CAMBIO DE DIAS CHEQUE: "
         If BoxMensaje = vbYes Then
            If JE >= 4 Then
               Titulo = "Pregunta de grabación"
               Mensajes = "SEGURO DE GRABAR CAMBIO DE DIAS CHEQUE: "
               sSQL = "UPDATE Trans_Dep_Chq " _
                    & "SET Dias = " & JE & " " _
                    & "WHERE Fecha = #" & BuscarFecha(FechaTexto) & "# " _
                    & "AND Cuenta_No = '" & CuentaBanco & "' " _
                    & "AND Cheque = '" & NumCheque & "' " _
                    & "AND Banco = '" & NombreBanco & "' " _
                    & "AND Valor = " & Total & " "
               'ConectarAdoExecute sSQL
            Else
               If ClaveSupervisor Then
                  sSQL = "UPDATE Trans_Dep_Chq " _
                       & "SET Dias = " & JE & " " _
                       & "WHERE Fecha = #" & BuscarFecha(FechaTexto) & "# " _
                       & "AND Cuenta_No = '" & CuentaBanco & "' " _
                       & "AND Cheque = '" & NumCheque & "' " _
                       & "AND Banco = '" & NombreBanco & "' " _
                       & "AND Valor = " & Total & " "
                  'ConectarAdoExecute sSQL
               End If
            End If
         End If
      End If
      Mifecha = BuscarFecha(FechaSistema)
      sSQL = "SELECT Dias,ME,Fecha,Cuenta_No,Banco,Cheque,Valor " _
           & "FROM Trans_Dep_Chq " _
           & "WHERE Fecha <= #" & Mifecha & "# AND T = 'P' " _
           & "AND Valor >= 0 " _
           & "ORDER BY ME,Fecha,Cuenta_No "
      'SelectDataGrid DGCheques, AdoCheques, sSQL
  
      sSQL = "SELECT Dias,ME,Fecha,Cuenta_No,Banco,Cheque,Valor " _
           & "FROM Trans_Dep_Chq " _
           & "WHERE Fecha <= #" & Mifecha & "# AND T = 'A' " _
           & "AND Valor >= 0 " _
           & "ORDER BY ME,Fecha,Cuenta_No "
      'SelectDataGrid DGProtesto, AdoProtesto, sSQL
      sSQL = "SELECT Codigo & ' => ' & Cuenta As NomBanco " _
           & "FROM Catalogo_Cuentas " _
           & "WHERE TC = 'BA' " _
           & "AND Item ='" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "ORDER BY Codigo "
      SelectDBList DLBanco, AdoBanco, sSQL, "NomBanco"
      RatonNormal
   End If
End Sub

Private Sub Form_Activate()
  TipoDoc = CompDiario
  Trans_No = 52
  BorrarAsientos True
  Mifecha = BuscarFecha(FechaSistema)
  sSQL = "SELECT ME,Fecha,Cuenta_No,Banco,Cheque,Creditos,Saldo_Cont,Saldo_Disp,ID,T,TP " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha <= #" & Mifecha & "# " _
       & "AND T = 'P' " _
       & "AND Creditos > 0 " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "ORDER BY ME,Fecha,Cuenta_No "
  SelectDataGrid DGCheques, AdoCheques, sSQL
  
  sSQL = "SELECT Dias,ME,Fecha,Cuenta_No,Banco,Cheque,Valor " _
       & "FROM Trans_Dep_Chq " _
       & "WHERE Fecha <= #" & Mifecha & "# AND T = 'A' " _
       & "AND Valor >= 0 " _
       & "AND Item ='" & NumEmpresa & "' " _
       & "ORDER BY ME,Fecha,Cuenta_No "
  'SelectDataGrid DGProtesto, AdoProtesto, sSQL
  
  sSQL = "SELECT Codigo & ' => ' & Cuenta As NomBanco " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'BA' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDBList DLBanco, AdoBanco, sSQL, "NomBanco"
  RatonNormal
  MBoxFecha.SetFocus
End Sub

Private Sub Form_Load()
CentrarForm ChequeEfectivo
ConectarAdodc AdoCtas
ConectarAdodc AdoBanco
ConectarAdodc AdoCheques
ConectarAdodc AdoProtesto
ConectarAdodc AdoAsientos
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha
End Sub

Private Sub TextComision_GotFocus()
  TextComision.Text = ""
End Sub

Private Sub TextComision_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextComision_LostFocus()
  TextoValido TextComision, True
End Sub

Private Sub TextCosto_GotFocus()
  TextCosto.Text = ""
End Sub

Private Sub TextCosto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCosto_LostFocus()
  TextoValido TextCosto, True
End Sub
