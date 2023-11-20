VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FContratosP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Contratos"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DGAbono 
      Bindings        =   "FContraP.frx":0000
      Height          =   1695
      Left            =   120
      TabIndex        =   20
      Top             =   2400
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2990
      _Version        =   393216
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
            LCID            =   2058
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
            LCID            =   2058
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
   Begin MSDataListLib.DataCombo DCGrupoNo 
      DataSource      =   "AdoGrupoNo"
      Height          =   315
      Left            =   3840
      TabIndex        =   19
      Top             =   105
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DCContrato 
      DataSource      =   "AdoContratos1"
      Height          =   315
      Left            =   1365
      TabIndex        =   18
      Top             =   735
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DCClientes 
      DataSource      =   "AdoClientes"
      Height          =   315
      Left            =   1365
      TabIndex        =   17
      Top             =   420
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc AdoArt 
      Height          =   330
      Left            =   2400
      Top             =   3240
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Left            =   2400
      Top             =   2880
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   240
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Top             =   3240
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Top             =   2880
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Top             =   2520
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
   Begin MSAdodcLib.Adodc AdoAbono 
      Height          =   330
      Left            =   2400
      Top             =   2520
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Abono"
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
   Begin VB.TextBox TextInt 
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
      Left            =   4410
      MaxLength       =   12
      TabIndex        =   10
      Top             =   1890
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
      Left            =   4410
      MaxLength       =   12
      TabIndex        =   8
      Top             =   1575
      Width           =   1695
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
      Left            =   1365
      MaxLength       =   8
      TabIndex        =   6
      Top             =   1575
      Width           =   1800
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
      Left            =   1365
      MaxLength       =   50
      TabIndex        =   12
      Top             =   1260
      Width           =   4740
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
      Left            =   6195
      Picture         =   "FContraP.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1785
      Width           =   960
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
      Left            =   6195
      Picture         =   "FContraP.frx":0297
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   840
      Width           =   960
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   1365
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
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Intereses"
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
      TabIndex        =   9
      Top             =   1890
      Width           =   1275
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Abono M/N"
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
      TabIndex        =   7
      Top             =   1575
      Width           =   1275
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " C&ontratos:"
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
      Top             =   735
      Width           =   1275
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Cliente: "
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
      Top             =   420
      Width           =   1275
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Left            =   4200
      TabIndex        =   15
      Top             =   4410
      Width           =   1695
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " T O T A L"
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
      TabIndex        =   16
      Top             =   4410
      Width           =   1065
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
      Left            =   105
      TabIndex        =   5
      Top             =   1575
      Width           =   1275
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " G&rupo No."
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
      Left            =   2625
      TabIndex        =   2
      Top             =   105
      Width           =   1170
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
      TabIndex        =   11
      Top             =   1260
      Width           =   1275
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Fecha:"
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
      Width           =   1275
   End
End
Attribute VB_Name = "FContratosP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
FechaValida MBoxFecha
 If Val(Label7.Caption) > 0 And DataArt.Recordset.RecordCount > 0 Then
    RatonReloj
    SelectData AdoAbono, "Abono_Meses"
    sSQL = "Abono_Contrato_" & CodigoUsuario
    SelectData AdoPagos, sSQL
    With AdoPagos.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
         Do While Not .EOF
            AdoAbono.Recordset.AddNew
            AdoAbono.Recordset.Fields("Fecha") = MBoxFecha.Text
            AdoAbono.Recordset.Fields("Cta") = .Fields("CTA")
            AdoAbono.Recordset.Fields("Codigo") = .Fields("CODIGO")
            AdoAbono.Recordset.Fields("Contrato_No") = .Fields("CONTRATO")
            AdoAbono.Recordset.Fields("Cuota_No") = .Fields("CUOTA_No")
            AdoAbono.Recordset.Fields("Abono") = .Fields("ABONO")
            AdoAbono.Recordset.Fields("Interes") = .Fields("INTERES")
            AdoAbono.Recordset.Update
           .MoveNext
         Loop
     End If
    End With
    sSQL = "UPDATE Contratos_Meses As CM,Abono_Contrato_" & CodigoUsuario & " As AbC " _
         & "SET CM.Saldo = CM.Saldo - AbC.ABONO," _
         & "CM.Fecha_C = #" & BuscarFecha(MBoxFecha.Text) & "# " _
         & "WHERE CM.Contrato_No = AbC.CONTRATO " _
         & "AND CM.Pago_No = AbC.CUOTA_No "
    UpdateData AdoPagos, sSQL
    
    sSQL = "UPDATE Contratos_Meses " _
         & "SET T = '" & Cancelado & "' " _
         & "WHERE Saldo <= 0 "
    UpdateData AdoPagos, sSQL
    
    sSQL = "DELETE * " _
         & "FROM Abono_Contrato_" & CodigoUsuario
    DeleteData AdoArt, sSQL
    
    sSQL = "SELECT * " _
         & "FROM Abono_Contrato_" & CodigoUsuario
    SelectDataGrid DGAbono, AdoArt, sSQL
    sSQL = "SELECT * " _
         & "FROM Contratos_Meses "
    SelectData AdoPagos, sSQL
    RatonNormal
    MensajeEncabData = DGAbono.Caption
    ImprimirAbonoPagos AdoPagos, MBoxFecha.Text, CodigoCli, Codigo
    DCContrato.SetFocus
 End If
End Sub

Private Sub Command2_Click()
  Unload FContratosP
End Sub


Private Sub DCClientes_LostFocus()
  CodigoCli = SinEspaciosDer(DCClientes.Text)
  If CodigoCli = "" Then CodigoCli = Ninguno
  sSQL = "SELECT Contrato_No " _
       & "FROM Contratos_Meses " _
       & "WHERE Codigo_C = '" & CodigoCli & "' " _
       & "AND T = '" & Procesado & "' " _
       & "GROUP BY Contrato_No "
  SelectDBCombo DCContrato, AdoContratos1, sSQL, "Contrato_No"
End Sub

Private Sub DCContrato_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
     Contrato_No = DCContrato.Text
     MiFecha = BuscarFecha(MBoxFecha.Text)
     sSQL = "SELECT CM.*,A.Cta_Ingreso,A.Articulo " _
          & "FROM Contratos_Meses As CM,Articulo As A " _
          & "WHERE Contrato_No = '" & Contrato_No & "' " _
          & "AND CM.Saldo > 0 " _
          & "AND CM.Codigo = A.Codigo " _
          & "ORDER BY CM.Pago_No "
     SelectData AdoPagos, sSQL
     With AdoPagos.Recordset
      If .RecordCount > 0 Then
          CodigoL = .Fields("Codigo")
          Cta_Ingreso = .Fields("Cta_Ingreso")
          TextConcepto.Text = .Fields("Concepto")
          TextContrato.Text = Contrato_No
          DGAbono.Caption = UCase(.Fields("Articulo"))
          Total = 0
          Do While Not .EOF
             Total = Total + .Fields("Saldo")
            .MoveNext
          Loop
          Total = Round(Total)
          TextMontoMN.Text = Format(Total, "#,##0.00")
          TextInt.Text = Format("0.00")
          If Total > 0 Then
             TextMontoMN.SetFocus
          Else
             DCClientes.SetFocus
          End If
      End If
     End With
  End If
End Sub

Private Sub DCContrato_LostFocus()
  Contrato_No = DCContrato.Text
End Sub

Private Sub DCGrupoNo_LostFocus()
  Grupo_No = DCGrupoNo.Text
  sSQL = "SELECT Cliente & Space(10) & ' Codigo: ' & Codigo As NomClientes " _
       & "FROM Clientes " _
       & "WHERE T = 'C' " _
       & "AND Grupo = '" & Grupo_No & "' " _
       & "ORDER BY Cliente "
  SelectDBCombo DCClientes, AdoClientes, sSQL, "NomClientes"
  CodigoCli = SinEspaciosDer(DCClientes.Text)
  If CodigoCli = "" Then CodigoCli = Ninguno
  sSQL = "SELECT Contrato_No " _
       & "FROM Contratos_Meses " _
       & "WHERE Codigo_C = '" & CodigoCli & "' " _
       & "AND T = '" & Procesado & "' " _
       & "GROUP BY Contrato_No "
  SelectDBCombo DCContrato, AdoContratos1, sSQL, "Contrato_No"
End Sub

Private Sub Form_Activate()
   'CTAbono_Contrato
   MiFecha = BuscarFecha(MBoxFecha.Text)
   sSQL = "SELECT Grupo " _
        & "FROM Clientes " _
        & "GROUP BY Grupo "
   SelectDBCombo DCGrupoNo, AdoGrupoNo, sSQL, "Grupo"
   Grupo_No = DCGrupoNo.Text
      
   sSQL = "SELECT Cliente & Space(10) & ' Codigo: ' & Codigo As NomClientes " _
        & "FROM Clientes " _
        & "WHERE T = 'C' " _
        & "AND Grupo = '" & Grupo_No & "' " _
        & "ORDER BY Cliente "
   SelectDBCombo DCClientes, AdoClientes, sSQL, "NomClientes"
   
   CodigoCli = SinEspaciosDer(DCClientes.Text)
   If CodigoCli = "" Then CodigoCli = Ninguno
   sSQL = "SELECT Contrato_No " _
        & "FROM Contratos_Meses " _
        & "WHERE Codigo_C = '" & CodigoCli & "' " _
        & "AND T = '" & Procesado & "' " _
        & "GROUP BY Contrato_No "
   SelectDBCombo DCContrato, AdoContratos1, sSQL, "Contrato_No"
   
   Contrato_No = DCContrato.Text
   sSQL = "DELETE * FROM Abono_Contrato_" & CodigoUsuario
   ConectarAdoExecute sSQL
   
   sSQL = "SELECT * FROM Abono_Contrato_" & CodigoUsuario
   SelectDataGrid DGAbono, AdoArt, sSQL
   sSQL = "SELECT * FROM Contratos_Meses "
   SelectData AdoPagos, sSQL
   RatonNormal
   DCContrato.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm FContratosP
  ConectarAdodc AdoArt
  ConectarAdodc AdoAbono
  ConectarAdodc AdoPagos
  ConectarAdodc AdoGrupoNo
  ConectarAdodc AdoClientes
  ConectarAdodc AdoContratos
  ConectarAdodc AdoContratos1
End Sub

Private Sub MBoxFecha_GotFocus()
  MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha
End Sub

Private Sub TextInt_GotFocus()
  MarcarTexto TextInt
End Sub

Private Sub TextInt_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextInt_LostFocus()
     sSQL = "DELETE * FROM Abono_Contrato_" & CodigoUsuario
     ConectarAdoExecute sSQL
     With AdoArt.Recordset
       If AdoPagos.Recordset.RecordCount > 0 Then
          AdoPagos.Recordset.MoveFirst
          Saldo = CDbl(TextMontoMN.Text)
          Do While Not AdoPagos.Recordset.EOF And Saldo > 0
             Abono = AdoPagos.Recordset.Fields("Saldo")
             If Saldo <= AdoPagos.Recordset.Fields("Saldo") Then Abono = Saldo
            .AddNew
            .Fields("CTA") = Cta_Ingreso
            .Fields("CODIGO") = CodigoCli
            .Fields("CONTRATO") = Contrato_No
            .Fields("CUOTA_No") = AdoPagos.Recordset.Fields("Pago_No")
            .Fields("ABONO") = Abono
            .Fields("INTERES") = 0
            .Update
             Saldo = Saldo - Abono
             AdoPagos.Recordset.MoveNext
          Loop
       End If
     End With
     sSQL = "SELECT * FROM Abono_Contrato_" & CodigoUsuario
     SelectDGrid DGAbono, AdoArt, sSQL
     With AdoArt.Recordset
         .MoveLast
         .Edit
         .Fields("INTERES") = CSng(TextInt.Text)
         .Update
         .MoveFirst
          Abono = 0
          Do While Not .EOF
             Abono = Abono + .Fields("ABONO") + .Fields("INTERES")
            .MoveNext
          Loop
          Label7.Caption = Format(Abono, "#,##0.00")
     End With
End Sub

Private Sub TextConcepto_GotFocus()
  MarcarTexto TextConcepto
End Sub

Private Sub TextConcepto_LostFocus()
  TextoValido TextConcepto, False
End Sub

Private Sub TextMontoMN_GotFocus()
  MarcarTexto TextMontoMN
End Sub

Private Sub TextMontoMN_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextMontoMN_LostFocus()
  TextoValido TextMontoMN, True
End Sub

