VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form Encaje 
   Caption         =   "Encaje"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   11265
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "&Anular Encaje Promocion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   10080
      Picture         =   "Encaje.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3255
      Width           =   1065
   End
   Begin VB.TextBox TxtDias 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6405
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "Encaje.frx":0442
      Top             =   1260
      Width           =   1485
   End
   Begin VB.TextBox TxtCheque 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7875
      MaxLength       =   8
      TabIndex        =   22
      Top             =   1680
      Width           =   2115
   End
   Begin VB.TextBox TxtBanco 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2205
      MaxLength       =   20
      TabIndex        =   20
      Top             =   1680
      Width           =   3585
   End
   Begin MSMask.MaskEdBox MBoxCuenta 
      Height          =   435
      Left            =   1785
      TabIndex        =   3
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   420
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   767
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "CCCCCCCCC-C"
      Mask            =   "########-#"
      PromptChar      =   "0"
   End
   Begin VB.CheckBox CheqCta 
      Caption         =   "Cuenta No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1785
      TabIndex        =   2
      Top             =   105
      Width           =   1590
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   435
      Left            =   105
      TabIndex        =   1
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   420
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   767
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
   Begin MSDataGridLib.DataGrid DGEncaje 
      Bindings        =   "Encaje.frx":0446
      Height          =   5055
      Left            =   105
      TabIndex        =   29
      Top             =   2100
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   8916
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
   Begin MSAdodcLib.Adodc AdoCtaNo 
      Height          =   330
      Left            =   315
      Top             =   2310
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
      Caption         =   "CtaNo"
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
   Begin VB.CommandButton Command6 
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
      Height          =   960
      Left            =   10080
      Picture         =   "Encaje.frx":045E
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5460
      Width           =   1065
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Consultar Encajes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   10080
      Picture         =   "Encaje.frx":0D28
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   105
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   3465
      TabIndex        =   4
      Top             =   0
      Width           =   1695
      Begin VB.OptionButton OpcA 
         Caption         =   "Anulados"
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
         Left            =   105
         TabIndex        =   6
         Top             =   525
         Width           =   1275
      End
      Begin VB.OptionButton OpcP 
         Caption         =   "Pendientes"
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
         Left            =   105
         TabIndex        =   5
         Top             =   210
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Desbloq. encaje y cheques"
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
      Left            =   10080
      Picture         =   "Encaje.frx":116A
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4305
      Width           =   1065
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
      Height          =   960
      Left            =   10080
      Picture         =   "Encaje.frx":15AC
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6510
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Anular Encaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   10080
      Picture         =   "Encaje.frx":1E76
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2205
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar Encaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   10080
      Picture         =   "Encaje.frx":22B8
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1155
      Width           =   1065
   End
   Begin VB.TextBox TxtMonto 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4305
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "Encaje.frx":26FA
      Top             =   1260
      Width           =   2115
   End
   Begin MSAdodcLib.Adodc AdoEncaje 
      Height          =   330
      Left            =   105
      Top             =   7140
      Width           =   9885
      _ExtentX        =   17436
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
      Caption         =   "Encaje"
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
   Begin VB.Label LabelTotalEncaje 
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
      Height          =   330
      Left            =   7875
      TabIndex        =   18
      Top             =   1260
      Width           =   2115
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total de Encajes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7875
      TabIndex        =   17
      Top             =   945
      Width           =   2115
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dias Encaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6405
      TabIndex        =   16
      Top             =   945
      Width           =   1485
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cheque/Abrevia."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5880
      TabIndex        =   21
      Top             =   1680
      Width           =   2010
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Detalle del Encaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   19
      Top             =   1680
      Width           =   2115
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Width           =   1590
   End
   Begin VB.Label LabelSaldoCont 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2205
      TabIndex        =   12
      Top             =   1260
      Width           =   2115
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Contable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2205
      TabIndex        =   11
      Top             =   945
      Width           =   2115
   End
   Begin VB.Label LabelSaldoDisp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   10
      Top             =   1260
      Width           =   2115
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Disponible"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   9
      Top             =   945
      Width           =   2115
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Monto de Encaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4305
      TabIndex        =   13
      Top             =   945
      Width           =   2115
   End
   Begin VB.Label LabelSocio 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5250
      TabIndex        =   8
      Top             =   420
      Width           =   4740
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nombre y Apellidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5250
      TabIndex        =   7
      Top             =   105
      Width           =   4740
   End
End
Attribute VB_Name = "Encaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  sSQL = "SELECT * " _
       & "FROM Trans_Bloqueos "
  SelectAdodc AdoCtaNo, sSQL
  With AdoCtaNo.Recordset
      .AddNew
      .Fields("T") = Normal
      .Fields("Fecha") = MBoxFecha
      .Fields("Cuenta_No") = MBoxCuenta
      .Fields("Valor") = Redondear(CCur(TxtMonto), 2)
      .Fields("Cheque") = TxtCheque
      .Fields("Banco") = TxtBanco
      .Fields("Dias") = Val(TxtDias.Text)
      .Fields("Item") = NumEmpresa
      .Update
  End With
  Listar_Encajes
End Sub

Private Sub Command2_Click()
   DGEncaje.Col = 1: FechaTexto = DGEncaje.Text
   DGEncaje.Col = 2: CuentaBanco = DGEncaje.Text
   DGEncaje.Col = 4: Valor = CCur(DGEncaje.Text)
   Mensajes = "Fecha:  " & FechaTexto & Chr(13) _
            & "Cuenta: " & CuentaBanco & Chr(13) _
            & "Valor:  " & Format(Valor, "#,##0.00") & Chr(13)
   Titulo = "Pregunta de grabación"
   If BoxMensaje = vbYes Then
      sSQL = "UPDATE Trans_Bloqueos SET T = 'A' " _
           & "WHERE Fecha = #" & BuscarFecha(FechaTexto) & "# " _
           & "AND Cuenta_No = '" & CuentaBanco & "' " _
           & "AND Valor = " & Valor & " "
      ConectarAdoExecute sSQL
   End If
   Listar_Encajes
End Sub

Private Sub Command3_Click()
   Unload Encaje
End Sub

Private Sub Command4_Click()
 If ClaveAdministrador Then
   RatonReloj
   sSQL = "SELECT T,Fecha,TP,Cheque,Debitos,Creditos,Saldo_Cont,Saldo_Disp,Hora " _
        & "FROM Trans_Libretas " _
        & "WHERE Cuenta_No = '" & MBoxCuenta.Text & "' " _
        & "ORDER BY Fecha,IDT,Hora,ID "
   SelectAdodc AdoCtaNo, sSQL
   With AdoCtaNo.Recordset
    If .RecordCount > 0 Then
       .MoveLast
        SaldoDisp = .Fields("Saldo_Disp")
       .Fields("Saldo_Disp") = SaldoDisp + CCur(TxtMonto.Text)
        Mensajes = "Actualizar Saldo Contable"
        Titulo = "Pregunta de grabación de Saldos"
        If BoxMensaje = vbYes Then
          .Fields("Saldo_Cont") = SaldoDisp + CCur(TxtMonto.Text)
        End If
       .Update
        MsgBox "Valor Desbloqueado"
    End If
   End With
   Else
      MsgBox "Usted no puede desbloquear este valor"
 End If
End Sub

Private Sub Command5_Click()
   DGEncaje.Visible = False
   SaldoDisp = 0: SaldoCont = 0: Total = 0
   FechaValida MBoxFecha, False
   FechaIni = BuscarFecha(MBoxFecha.Text)
   Listar_Encajes
   LabelSaldoDisp.Caption = Format(SaldoDisp, "#,##0.00")
   LabelSaldoCont.Caption = Format(SaldoCont, "#,##0.00")
   LabelTotalEncaje.Caption = Format(Total, "#,##0.00")
   DGEncaje.Visible = True
End Sub

Private Sub Command6_Click()
  SQLMsg1 = "REPORTE DE ENCAJES"
  If OpcP.value Then
     SQLMsg2 = "PENDIENTES"
  Else
     SQLMsg2 = "ANULADOS"
  End If
  ImprimirAdodc AdoEncaje, True, 1, 8
End Sub

Private Sub Command7_Click()
   DGEncaje.Col = 1: FechaTexto = DGEncaje.Text
   DGEncaje.Col = 2: CuentaBanco = DGEncaje.Text
   DGEncaje.Col = 4: Valor = CCur(DGEncaje.Text)
   Mensajes = "ANULACION DE ENCAJES POR PROMOCIONES:"
   Titulo = "Pregunta de Grabación"
   If BoxMensaje = vbYes Then
      sSQL = "UPDATE Trans_Bloqueos " _
           & "SET T = 'A' " _
           & "WHERE Fecha <= #" & BuscarFecha(FechaSistema) & "# " _
           & "AND Cheque = 'DEPP' " _
           & "AND T <> 'A' "
      ConectarAdoExecute sSQL
   End If
   Listar_Encajes
End Sub

Private Sub DGEncaje_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
     DGEncaje.Visible = False
     GenerarDataTexto Encaje, AdoEncaje
     DGEncaje.Visible = True
  End If
End Sub

Private Sub Form_Activate()
   RatonNormal
End Sub

Private Sub Form_Load()
   ConectarAdodc AdoCtaNo
   ConectarAdodc AdoEncaje
End Sub

Private Sub MBoxCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxCuenta_LostFocus()
   DGEncaje.Visible = False
   SaldoDisp = 0: SaldoCont = 0
   FechaValida MBoxFecha
   FechaIni = BuscarFecha(MBoxFecha.Text)
   Codigo = Ninguno
   sSQL = "SELECT * " _
        & "FROM Clientes_Datos_Extras " _
        & "WHERE Cuenta_No = '" & MBoxCuenta.Text & "' " _
        & "AND Tipo_Dato = 'LIBRETAS' "
   SelectAdodc AdoEncaje, sSQL
   With AdoEncaje.Recordset
    If .RecordCount > 0 Then Codigo = .Fields("Codigo")
   End With
   sSQL = "SELECT * FROM Clientes " _
        & "WHERE Codigo = '" & Codigo & "' "
   SelectAdodc AdoEncaje, sSQL
   With AdoEncaje.Recordset
    If .RecordCount > 0 Then
        LabelSocio.Caption = " " & .Fields("Cliente")
        sSQL = "SELECT * FROM Trans_Libretas " _
             & "WHERE Cuenta_No = '" & MBoxCuenta.Text & "' " _
             & "ORDER BY Fecha,IDT,Hora,ID "
        SelectData AdoCtaNo, sSQL
        With AdoCtaNo.Recordset
         If .RecordCount > 0 Then
            .MoveLast
             SaldoDisp = .Fields("Saldo_Disp")
             SaldoCont = .Fields("Saldo_Cont")
         End If
        End With
        Listar_Encajes
        TxtMonto.SetFocus
    Else
        LabelSocio.Caption = "No existe"
        MBoxCuenta.Text = "00000000-0"
    End If
   End With
   LabelSaldoDisp.Caption = Format(SaldoDisp, "#,##0.00")
   LabelSaldoCont.Caption = Format(SaldoCont, "#,##0.00")
End Sub

Private Sub MBoxFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha
End Sub

Private Sub TxtBanco_GotFocus()
  MarcarTexto TxtBanco
End Sub

Private Sub TxtBanco_KeyDown(KeyCode As Integer, Shift As Integer)
 PresionoEnter KeyCode
End Sub

Private Sub TxtBanco_LostFocus()
  TextoValido TxtBanco
End Sub

Private Sub TxtCheque_GotFocus()
  MarcarTexto TxtCheque
End Sub

Private Sub TxtCheque_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCheque_LostFocus()
  TextoValido TxtCheque
End Sub

Private Sub TxtDias_GotFocus()
 TxtDias.Text = "365"
 MarcarTexto TxtDias
End Sub

Private Sub TxtDias_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDias_LostFocus()
  TextoValido TxtDias, True, , 0
End Sub

Private Sub TxtMonto_GotFocus()
 TxtMonto.Text = ""
End Sub

Private Sub TxtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyS Then
     Total = Round(CCur(TxtMonto.Text), 2)
     Mensajes = "Seguro de Reversar USD " & Total & " al Saldo" & vbCrLf _
              & "De la Cuenta No. " & MBoxCuenta.Text
     Titulo = "Pregunta de Reversión"
     If BoxMensaje = vbYes Then
        sSQL = "SELECT TOP 1 * " _
             & "FROM Trans_Libretas " _
             & "WHERE Cuenta_No = '" & MBoxCuenta.Text & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
        SelectAdodc AdoEncaje, sSQL
        If AdoEncaje.Recordset.RecordCount > 0 Then
           SaldoDisp = AdoEncaje.Recordset.Fields("Saldo_Disp")
           SaldoCont = AdoEncaje.Recordset.Fields("Saldo_Cont")
           AdoEncaje.Recordset.Fields("Saldo_Disp") = Round(SaldoDisp - Total, 2)
           AdoEncaje.Recordset.Update
        End If
     End If
  End If
End Sub

Private Sub TxtMonto_LostFocus()
  TxtMonto.Text = Format(Val(TxtMonto.Text), "#,##0.00")
  
End Sub

Public Sub Listar_Encajes()
   DGEncaje.Visible = False
   sSQL = "SELECT B.T,B.Fecha,B.Cuenta_No,Cliente,Valor,B.Banco,B.Cheque,B.Dias " _
        & "FROM Trans_Bloqueos As B, Clientes_Datos_Extras As C,Clientes Cl "
   If OpcP.value Then
      sSQL = sSQL & "WHERE B.T = 'N' "
   Else
      sSQL = sSQL & "WHERE B.T = 'A' "
   End If
   sSQL = sSQL & "AND B.Cuenta_No = C.Cuenta_No " _
        & "AND C.Tipo_Dato = 'LIBRETAS' " _
        & "AND C.Codigo = Cl.Codigo "
   If CheqCta.value = 1 Then sSQL = sSQL & "AND C.Cuenta_No = '" & MBoxCuenta.Text & "' "
   sSQL = sSQL & "ORDER BY Cliente,B.Cuenta_No,B.Fecha "
   SelectDataGrid DGEncaje, AdoEncaje, sSQL
   Cadena = ""
   With AdoEncaje.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           Total = Total + .Fields("Valor")
           If Len(.Fields("Banco")) > 1 And Len(.Fields("Cheque")) > 1 And .Fields("Dias") = 0 Then
               Cadena = Cadena & "Cuenta No. " & .Fields("Cuenta_No") _
                      & " - Cheque No. " & .Fields("Cheque") & " - Por USD " & Format(.Fields("Valor"), "#,##0.00") _
                      & vbCrLf
           End If
          .MoveNext
        Loop
    End If
   End With
   If Cadena <> "" Then MsgBox "DESACTIVE LOS SIGUIENTES DEPOSITOS:" & vbCrLf & vbCrLf & Cadena
   LabelSaldoDisp.Caption = Format(SaldoDisp, "#,##0.00")
   LabelSaldoCont.Caption = Format(SaldoCont, "#,##0.00")
   LabelTotalEncaje.Caption = Format(Total, "#,##0.00")
   DGEncaje.Visible = True
End Sub
