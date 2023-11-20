VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSChrt20.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ResumenProduccion 
   Caption         =   "RESUMEN DE PRODUCCION MENSUAL Y DIARIA"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   105
      TabIndex        =   7
      Top             =   525
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "RESUMEN DIARIO Y MENSUAL"
      TabPicture(0)   =   "ResProdM.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblHoras"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "AdoQuery1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DGQuery1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "AdoQuery"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DGQuery"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "GRAFICO DEL RESULTADO"
      TabPicture(1)   =   "ResProdM.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSChart1"
      Tab(1).Control(1)=   "MSChart2"
      Tab(1).ControlCount=   2
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   2850
         Left            =   -74895
         OleObjectBlob   =   "ResProdM.frx":0038
         TabIndex        =   12
         Top             =   420
         Width           =   10410
      End
      Begin MSDataGridLib.DataGrid DGQuery 
         Bindings        =   "ResProdM.frx":24F0
         Height          =   2745
         Left            =   105
         TabIndex        =   8
         Top             =   420
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   4842
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
      Begin MSAdodcLib.Adodc AdoQuery 
         Height          =   330
         Left            =   105
         Top             =   3150
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   30
         CommandTimeout  =   30
         CursorType      =   2
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
         Caption         =   "Query"
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
      Begin MSDataGridLib.DataGrid DGQuery1 
         Bindings        =   "ResProdM.frx":2507
         Height          =   2745
         Left            =   105
         TabIndex        =   9
         Top             =   3570
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   4842
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
      Begin MSAdodcLib.Adodc AdoQuery1 
         Height          =   330
         Left            =   105
         Top             =   6300
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   30
         CommandTimeout  =   30
         CursorType      =   2
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
         Caption         =   "Query"
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
      Begin MSChart20Lib.MSChart MSChart2 
         Height          =   3165
         Left            =   -74895
         OleObjectBlob   =   "ResProdM.frx":251F
         TabIndex        =   13
         Top             =   3465
         Width           =   10410
      End
      Begin VB.Label LblHoras 
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
         Left            =   8715
         TabIndex        =   11
         Top             =   6300
         Width           =   1695
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Horas"
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
         Left            =   7560
         TabIndex        =   10
         Top             =   6300
         Width           =   1170
      End
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
      Height          =   750
      Left            =   10815
      Picture         =   "ResProdM.frx":4875
      Style           =   1  'Graphical
      TabIndex        =   4
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
      Height          =   750
      Left            =   10815
      Picture         =   "ResProdM.frx":4CB7
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1785
      Width           =   960
   End
   Begin VB.CommandButton Command5 
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
      Height          =   750
      Left            =   10815
      Picture         =   "ResProdM.frx":5581
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   945
      Width           =   960
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   2940
      TabIndex        =   3
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   840
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   210
      Top             =   1470
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin MSAdodcLib.Adodc AdoAux1 
      Height          =   330
      Left            =   210
      Top             =   1785
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Aux1"
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
      Left            =   210
      Top             =   2205
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Desde"
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
      Width           =   750
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Hasta"
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
      Left            =   2205
      TabIndex        =   2
      Top             =   105
      Width           =   750
   End
End
Attribute VB_Name = "ResumenProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim TValorDia(32) As Double
Dim TValorProd(32) As Double
  DGQuery.Visible = False
  DGQuery1.Visible = False
  Limpiar_Reporte_Produccion
  RatonReloj
  For i = 1 To 31
      TValorDia(i) = 0
      TValorProd(i) = 0
  Next i
 'Resumen de Produccion del mes
  sSQL = "SELECT TC,Codigo,Fecha,SUM(Total) As V_Total " _
       & "FROM Detalle_Factura " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND T <> 'A' " _
       & "GROUP BY TC,Codigo,Fecha " _
       & "ORDER BY TC,Codigo,Fecha "
  Select_Adodc AdoAux1, sSQL
  With AdoAux1.Recordset
   If .RecordCount > 0 Then
       CodigoP = .Fields("Codigo")
       TipoProc = .Fields("TC")
       NoMeses = Month(.Fields("Fecha"))
       Mifecha = .Fields("Fecha")
       NoAnio = Year(.Fields("Fecha"))
       For i = 1 To 31
           TValorDia(i) = 0
       Next i
       Contador = 0
       Do While Not .EOF
          Contador = Contador + 1
          ResumenProduccion.Caption = Format$(Contador / .RecordCount, "00%")
          If TipoProc <> .Fields("TC") Or Codigo <> .Fields("Codigo") Or NoMeses <> Month(.Fields("Fecha")) Then
             Total = 0: J = 0
             SQL2 = "UPDATE Saldo_Diarios " _
                  & "SET "
             For i = 1 To 31
                 SQL2 = SQL2 & "D_" & Format$(i, "00") & " = " & TValorDia(i) & ","
                 Total = Total + TValorDia(i)
             Next i
             SQL2 = SQL2 & "Total = " & Total & " " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND CodigoU = '" & CodigoUsuario & "' " _
                  & "AND Codigo_Aux = '" & CodigoP & "' " _
                  & "AND No_Mes = " & NoMeses & " " _
                  & "AND TC = '" & TipoProc & "' " _
                  & "AND TP = 'RPM' "
             'MsgBox SQL2
             Ejecutar_SQL_SP SQL2
             For i = 1 To 31
                 TValorDia(i) = 0
             Next i
             CodigoP = .Fields("Codigo")
             TipoProc = .Fields("TC")
             NoMeses = Month(.Fields("Fecha"))
             Mifecha = .Fields("Fecha")
             NoAnio = Year(.Fields("Fecha"))
          End If
          i = Day(.Fields("Fecha"))
          TValorDia(i) = TValorDia(i) + .Fields("V_Total")
          TValorProd(i) = TValorProd(i) + .Fields("V_Total")
          
         .MoveNext
       Loop
       Total = 0: J = 0
       SQL2 = "UPDATE Saldo_Diarios " _
            & "SET "
       For i = 1 To 31
           SQL2 = SQL2 & "D_" & Format$(i, "00") & " = " & TValorDia(i) & ","
           Total = Total + TValorDia(i)
       Next i
       SQL2 = SQL2 & "Total = " & Total & " " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND CodigoU = '" & CodigoUsuario & "' " _
            & "AND Codigo_Aux = '" & CodigoP & "' " _
            & "AND No_Mes = " & NoMeses & " " _
            & "AND TC = '" & TipoProc & "' " _
            & "AND TP = 'RPM' "
       Ejecutar_SQL_SP SQL2
   End If
  End With
  Total = 0: J = 0
  SQL2 = "UPDATE Saldo_Diarios " _
       & "SET "
  For i = 1 To 31
      SQL2 = SQL2 & "D_" & Format$(i, "00") & " = " & TValorProd(i) & ","
      Total = Total + TValorProd(i)
  Next i
  SQL2 = SQL2 & "Total = " & Total & " " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND Codigo_Aux = 'Total' " _
       & "AND No_Mes = 0 " _
       & "AND TC = 'TT' " _
       & "AND TB = 'RPM' "
  Ejecutar_SQL_SP SQL2
  
 'Resumen de Produccion del año
  For i = 1 To 31
      TValorProd(i) = 0
  Next i
  NoAnio = Year(MBFechaI)
  FechaIni = BuscarFecha("01/01/" & Format$(NoAnio, "0000"))
  FechaFin = BuscarFecha("31/12/" & Format$(NoAnio, "0000"))
  sSQL = "SELECT TC,Codigo,Fecha,SUM(Total) As V_Total " _
       & "FROM Detalle_Factura " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND T <> 'A' " _
       & "GROUP BY TC,Codigo,Fecha " _
       & "ORDER BY TC,Codigo,Fecha "
  Select_Adodc AdoAux1, sSQL
  With AdoAux1.Recordset
   If .RecordCount > 0 Then
       CodigoP = .Fields("Codigo")
       TipoProc = .Fields("TC")
       NoMeses = Month(.Fields("Fecha"))
       Mifecha = .Fields("Fecha")
       NoAnio = Year(.Fields("Fecha"))
       For i = 1 To 31
           TValorDia(i) = 0
       Next i
       Contador = 0
       Do While Not .EOF
          Contador = Contador + 1
          ResumenProduccion.Caption = Format$(Contador / .RecordCount, "00%")
          If TipoProc <> .Fields("TC") Or Codigo <> .Fields("Codigo") Then
             Total = 0: J = 0
             SQL2 = "UPDATE Saldo_Diarios " _
                  & "SET "
             For i = 1 To 12
                 Mes1 = MesesLetras(CInt(i))
                 SQL2 = SQL2 & Mes1 & " = " & TValorDia(i) & ","
                 Total = Total + TValorDia(i)
             Next i
             SQL2 = SQL2 & "Total = " & Total & " " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND CodigoU = '" & CodigoUsuario & "' " _
                  & "AND Cta = '" & CodigoP & "' " _
                  & "AND TC = '" & TipoProc & "' " _
                  & "AND TP = 'RPM' "
             'MsgBox SQL2
             Ejecutar_SQL_SP SQL2
             For i = 1 To 31
                 TValorDia(i) = 0
             Next i
             CodigoP = .Fields("Codigo")
             TipoProc = .Fields("TC")
             NoMeses = Month(.Fields("Fecha"))
             Mifecha = .Fields("Fecha")
             NoAnio = Year(.Fields("Fecha"))
          End If
          i = Month(.Fields("Fecha"))
          TValorDia(i) = TValorDia(i) + .Fields("V_Total")
          TValorProd(i) = TValorProd(i) + .Fields("V_Total")
         .MoveNext
       Loop
       Total = 0: J = 0
       SQL2 = "UPDATE Saldo_Diarios " _
            & "SET "
       For i = 1 To 12
           Mes1 = MesesLetras(CInt(i))
           SQL2 = SQL2 & Mes1 & " = " & TValorDia(i) & ","
           Total = Total + TValorDia(i)
       Next i
       SQL2 = SQL2 & "Total = " & Total & " " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND CodigoU = '" & CodigoUsuario & "' " _
            & "AND Cta = '" & CodigoP & "' " _
            & "AND TC = '" & TipoProc & "' " _
            & "AND TP = 'RPM' "
       Ejecutar_SQL_SP SQL2
    End If
  End With
  Total = 0: J = 0
  SQL2 = "UPDATE Saldo_Diarios " _
       & "SET "
  For i = 1 To 12
      Mes1 = MesesLetras(CInt(i))
      SQL2 = SQL2 & Mes1 & " = " & TValorProd(i) & ","
      Total = Total + TValorProd(i)
  Next i
  SQL2 = SQL2 & "Total = " & Total & " " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND Cta = 'Total' " _
       & "AND TC = 'TT' " _
       & "AND TP = 'RPM' "
  Ejecutar_SQL_SP SQL2

  SQL2 = "DELETE * " _
       & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND Total = 0 " _
       & "AND TP = 'RPM' "
  Ejecutar_SQL_SP SQL2
  
  DGQuery.Caption = "RESUMEN DE PRODUCCION DEL " & MBFechaI & " AL " & MBFechaF
  DGQuery1.Caption = "RESUMEN DE PRODUCCION DEL AÑO " & Year(MBFechaI)
  'If SQL_Server = False Then MsgBox "Proceso Terminado, Puede Consultar"
  sSQL = "SELECT BDM.Cuenta As Producto,TM.Dia_Mes As Mes,"
  For i = 1 To 31
      sSQL = sSQL & "BDM.D_" & Format$(i, "00") & ","
  Next i
  sSQL = sSQL & "Total " _
       & "FROM Balances_Mes As BDM,Tabla_Dias_Meses AS TM " _
       & "WHERE BDM.TP = 'RPM' " _
       & "AND BDM.CodigoU = '" & CodigoUsuario & "' " _
       & "AND BDM.Item = '" & NumEmpresa & "' " _
       & "AND TM.Tipo = 'M' " _
       & "AND TM.No_D_M = BDM.No_Mes " _
       & "ORDER BY BDM.TC,BDM.Codigo,TM.No_D_M "
  Select_Adodc_Grid DGQuery, AdoQuery, sSQL
  
  sSQL = "SELECT Comprobante As Producto,"
  For i = 1 To 12
      Mes1 = MesesLetras(CInt(i))
      sSQL = sSQL & Mes1 & ","
  Next i
  sSQL = sSQL & "Total " _
       & "FROM Saldo_Diarios " _
       & "WHERE TP = 'RPM' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "ORDER BY TC,Cta "
  Select_Adodc_Grid DGQuery1, AdoQuery1, sSQL
  Total = 0
  With AdoQuery1.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Total = Total + .Fields("Total")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  LblHoras.Caption = Format$(Total, "#,##0.00")
  ResumenProduccion.Caption = "RESUMEN DE PRODUCCION MENSUAL Y DIARIA"
  DGQuery.Visible = True
  DGQuery1.Visible = True
  RatonNormal
  Opcion = 1
End Sub

Private Sub Command3_Click()
  'Salir
  Unload ResumenProduccion
End Sub
'Imprimir Datos
Private Sub Command5_Click()
  MensajeEncabData = SSTab1.Caption
  Select Case SSTab1.Tab
    Case 0: Imprimir_Entradas_Salidas AdoQuery, 2, 6, SSTab1.Tab, True
    Case 1: Imprimir_Entradas_Salidas AdoQuery, 2, 6, SSTab1.Tab, True
    Case 2: Imprimir_Entradas_Salidas AdoQuery, 1, 9, SSTab1.Tab
  End Select
End Sub

Private Sub Form_Activate()
  RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoAux1
  ConectarAdodc AdoQuery
  ConectarAdodc AdoQuery1
  ConectarAdodc AdoClientes
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
  MBFechaF = UltimoDiaMes(MBFechaI)
End Sub

Public Sub Limpiar_Reporte_Produccion()
  RatonReloj
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
    
  SQL2 = "DELETE * " _
       & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TP = 'RPM' "
  Ejecutar_SQL_SP SQL2
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "ORDER BY Codigo_Inv "
  Select_Adodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount Then
       i = Month(MBFechaI)
       J = Month(MBFechaF)
       Contador = 0
       Do While Not .EOF
          NoMeses = Month(MBFechaI)
          Contador = Contador + 1
          ResumenProduccion.Caption = Format$(Contador / .RecordCount, "00%") & " - " & .Fields("Producto")
          For K = 1 To J - i + 1
             'FA
              SetAdoAddNew "Saldo_Diarios"
              SetAdoFields "Codigo_Aux", .Fields("Codigo_Inv")
              SetAdoFields "Comprobante", .Fields("Producto")
              SetAdoFields "TC", "FA"
              SetAdoFields "TP", "RPM"
              SetAdoFields "No_Mes", CByte(NoMeses)
              SetAdoFields "Item", NumEmpresa
              SetAdoFields "CodigoU", CodigoUsuario
              SetAdoUpdate
             'NV
              SetAdoAddNew "Saldo_Diarios"
              SetAdoFields "Codigo_Aux", .Fields("Codigo_Inv")
              SetAdoFields "Comprobante", .Fields("Producto")
              SetAdoFields "TC", "NV"
              SetAdoFields "TP", "RPM"
              SetAdoFields "No_Mes", CByte(NoMeses)
              SetAdoFields "Item", NumEmpresa
              SetAdoFields "CodigoU", CodigoUsuario
              SetAdoUpdate
             'PV
              SetAdoAddNew "Saldo_Diarios"
              SetAdoFields "Codigo_Aux", .Fields("Codigo_Inv")
              SetAdoFields "Comprobante", .Fields("Producto")
              SetAdoFields "TC", "PV"
              SetAdoFields "TP", "RPM"
              SetAdoFields "No_Mes", CByte(NoMeses)
              SetAdoFields "Item", NumEmpresa
              SetAdoFields "CodigoU", CodigoUsuario
              SetAdoUpdate
              NoMeses = NoMeses + 1
              If NoMeses > 12 Then NoMeses = 1
          Next K
          
          SetAdoAddNew "Saldo_Diarios"
          SetAdoFields "Cta", .Fields("Codigo_Inv")
          SetAdoFields "Comprobante", .Fields("Producto")
          SetAdoFields "TC", "FA"
          SetAdoFields "TP", "RPM"
          SetAdoFields "Item", NumEmpresa
          SetAdoFields "CodigoU", CodigoUsuario
          SetAdoUpdate
          
          SetAdoAddNew "Saldo_Diarios"
          SetAdoFields "Cta", .Fields("Codigo_Inv")
          SetAdoFields "Comprobante", .Fields("Producto")
          SetAdoFields "TC", "NV"
          SetAdoFields "TP", "RPM"
          SetAdoFields "Item", NumEmpresa
          SetAdoFields "CodigoU", CodigoUsuario
          SetAdoUpdate
          
          SetAdoAddNew "Saldo_Diarios"
          SetAdoFields "Cta", .Fields("Codigo_Inv")
          SetAdoFields "Comprobante", .Fields("Producto")
          SetAdoFields "TC", "PV"
          SetAdoFields "TP", "RPM"
          SetAdoFields "Item", NumEmpresa
          SetAdoFields "CodigoU", CodigoUsuario
          SetAdoUpdate
         .MoveNext
       Loop
   End If
  End With
          
  SetAdoAddNew "Saldo_Diarios"
  SetAdoFields "Cta", "Total"
  SetAdoFields "Comprobante", "PRODUCCION TOTAL"
  SetAdoFields "TC", "TT"
  SetAdoFields "TP", "RPM"
  SetAdoFields "Item", NumEmpresa
  SetAdoFields "CodigoU", CodigoUsuario
  SetAdoUpdate
  RatonNormal
End Sub

