VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FReportesUAF 
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13650
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   13650
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   330
      Left            =   10500
      TabIndex        =   7
      Top             =   105
      Width           =   1800
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6630
      Left            =   105
      TabIndex        =   8
      Top             =   525
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   11695
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "CLI"
      TabPicture(0)   =   "FRepoUAF.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DGCLI"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "AdoCLI"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "PRO"
      TabPicture(1)   =   "FRepoUAF.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "AdoPRO"
      Tab(1).Control(1)=   "DGPRO"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "TRA"
      TabPicture(2)   =   "FRepoUAF.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DGTRA"
      Tab(2).Control(1)=   "AdoTRA"
      Tab(2).ControlCount=   2
      Begin MSAdodcLib.Adodc AdoCLI 
         Height          =   330
         Left            =   105
         Top             =   6195
         Width           =   7155
         _ExtentX        =   12621
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
         Caption         =   "Adodc1"
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
      Begin MSDataGridLib.DataGrid DGCLI 
         Bindings        =   "FRepoUAF.frx":0054
         Height          =   5790
         Left            =   105
         TabIndex        =   9
         Top             =   420
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   10213
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
      Begin MSDataGridLib.DataGrid DGPRO 
         Bindings        =   "FRepoUAF.frx":0069
         Height          =   4215
         Left            =   -74895
         TabIndex        =   10
         Top             =   420
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7435
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
      Begin MSAdodcLib.Adodc AdoPRO 
         Height          =   330
         Left            =   -74895
         Top             =   4725
         Width           =   7155
         _ExtentX        =   12621
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
         Caption         =   "Adodc1"
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
      Begin MSDataGridLib.DataGrid DGTRA 
         Bindings        =   "FRepoUAF.frx":007E
         Height          =   5790
         Left            =   -74895
         TabIndex        =   11
         Top             =   420
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   10213
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
      Begin MSAdodcLib.Adodc AdoTRA 
         Height          =   330
         Left            =   -74895
         Top             =   6195
         Width           =   7155
         _ExtentX        =   12621
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
         Caption         =   "Adodc1"
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
   End
   Begin VB.TextBox TextValor 
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
      Left            =   6195
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   105
      Width           =   2325
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Consultar Reportes"
      Height          =   330
      Left            =   8610
      TabIndex        =   6
      Top             =   105
      Width           =   1800
   End
   Begin VB.ComboBox CMes 
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
      Left            =   3255
      TabIndex        =   3
      Text            =   "Enero"
      Top             =   105
      Width           =   1416
   End
   Begin VB.ComboBox CAño 
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
      ItemData        =   "FRepoUAF.frx":0093
      Left            =   945
      List            =   "FRepoUAF.frx":0095
      TabIndex        =   1
      Text            =   "2000"
      Top             =   105
      Width           =   1410
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   630
      Top             =   7770
      Visible         =   0   'False
      Width           =   2955
      _ExtentX        =   5212
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
      Caption         =   "Adodc1"
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
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Monto Ma&ximo:"
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
      Left            =   4725
      TabIndex        =   4
      Top             =   105
      Width           =   1485
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Mes:"
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
      Left            =   2415
      TabIndex        =   2
      Top             =   105
      Width           =   855
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Año:"
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
End
Attribute VB_Name = "FReportesUAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   Total = Val(TextValor)
   FechaInicial = "01/" & CMes & "/" & CAño
   FechaFinal = UltimoDiaMes(FechaInicial)
   FechaIni = BuscarFecha(FechaInicial)
   FechaFin = BuscarFecha(FechaFinal)
   sSQL = "UPDATE Clientes_Datos_Extras " _
        & "SET X = '.' " _
        & "WHERE Tipo_Dato = 'LIBRETAS' "
   Ejecutar_SQL_SP sSQL
   
   sSQL = "SELECT C.Cliente,TL.Cuenta_No,SUM(TL.Debitos) As TDebitos,SUM(TL.Creditos) As TCreditos " _
        & "FROM Trans_Libretas As TL, Clientes_Datos_Extras As CL, Clientes As C " _
        & "WHERE TL.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND CL.Tipo_Dato = 'LIBRETAS' " _
        & "AND C.Codigo = CL.Codigo " _
        & "AND CL.Cuenta_No = TL.Cuenta_No " _
        & "GROUP BY C.Cliente,TL.Cuenta_No " _
        & "HAVING SUM(TL.Creditos) >= " & Total & " "
   Select_Adodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           sSQL = "UPDATE Clientes_Datos_Extras " _
                & "SET X = 'T' " _
                & "WHERE Cuenta_No = '" & .Fields("Cuenta_No") & "' " _
                & "AND Tipo_Dato = 'LIBRETAS' "
           Ejecutar_SQL_SP sSQL
          .MoveNext
        Loop
    End If
   End With
   sSQL = "SELECT C.TD As [TIPO ID]," _
        & "C.CI_RUC As ID," _
        & "TL.Fecha As [FECHA DE TRANSACCION]," _
        & "(ID & IDT) As [NUMERO DE TRANSACCION]," _
        & "TL.Cuenta_No As [NUMERO  DE CUENTA/OPERACIÓN]," _
        & "TL.Debitos As [VALOR DEBITO]," _
        & "TL.Creditos As [VALOR CREDITO]," _
        & "'USD' As MONEDA," _
        & "TL.TP As [TIPO DE TRANACCION]," _
        & "C.Cliente As [NOMBRE/ RAZON SOCIAL  ORDENANTE/ BENEFICIARIO]," _
        & "'EC' As [PAIS DESTINO/ORIGEN] " _
        & "FROM Trans_Libretas As TL, Clientes_Datos_Extras As CL, Clientes As C " _
        & "WHERE TL.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND CL.Tipo_Dato = 'LIBRETAS' " _
        & "AND CL.X = 'T' " _
        & "AND C.Codigo = CL.Codigo " _
        & "AND CL.Cuenta_No = TL.Cuenta_No " _
        & "ORDER BY C.Cliente,TL.Cuenta_No "
   Select_Adodc_Grid DGTRA, AdoTRA, sSQL
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub DGTRA_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then GenerarDataTexto FReportesUAF, AdoTRA, True
End Sub

Private Sub Form_Activate()
   CMes.Clear
   CAño.Clear
   sSQL = Listar_Meses
   Select_Adodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           CMes.AddItem .Fields("Dia_Mes")
           CMes.Tag = .Fields("No_D_M")
          .MoveNext
        Loop
    End If
   End With
   For I = Year(FechaSistema) To 2000 Step -1
       CAño.AddItem Format(I, "0000")
   Next I
   CAño.Text = CAño.List(0)
   CMes.Text = MesesLetras(Month(FechaSistema))
   RatonNormal
   CAño.SetFocus
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoCLI
  ConectarAdodc AdoPRO
  ConectarAdodc AdoTRA
End Sub
