VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Begin VB.Form FLibroRetenciones 
   Caption         =   "LIBRO DE RETENCIONES"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13560
   Icon            =   "FLibroRet.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   13560
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   1402
      ButtonWidth     =   1455
      ButtonHeight    =   1244
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Módulo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Retenciones"
            Object.ToolTipText     =   "Resumen de Retenciones"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Por_Codigo"
            Object.ToolTipText     =   "Resumen por Código de Retencion"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Por_Detalle"
            Object.ToolTipText     =   "Resumen Detalle por Código de Retención"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print_Retenciones"
            Object.ToolTipText     =   "Imprimir Resumen de Retenciones"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   4200
         TabIndex        =   1
         Top             =   -105
         Width           =   7050
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
            ItemData        =   "FLibroRet.frx":0442
            Left            =   735
            List            =   "FLibroRet.frx":0444
            TabIndex        =   3
            Text            =   "2000"
            Top             =   315
            Width           =   1410
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
            Left            =   2835
            TabIndex        =   5
            Text            =   "Enero"
            Top             =   315
            Width           =   1416
         End
         Begin VB.ListBox LstAT 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   4305
            TabIndex        =   6
            ToolTipText     =   "Seleccione el tipo de Anexo"
            Top             =   210
            Width           =   2640
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
            TabIndex        =   2
            Top             =   315
            Width           =   645
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
            Left            =   2205
            TabIndex        =   4
            Top             =   315
            Width           =   645
         End
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&S"
      Height          =   330
      Left            =   105
      Picture         =   "FLibroRet.frx":0446
      TabIndex        =   8
      ToolTipText     =   "Salir"
      Top             =   3360
      Width           =   225
   End
   Begin MSDataGridLib.DataGrid DGLibroRet 
      Bindings        =   "FLibroRet.frx":0D10
      Height          =   3165
      Left            =   105
      TabIndex        =   7
      Top             =   840
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   5583
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   20
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Datos Procesados"
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   315
      Top             =   1995
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   315
      Top             =   2310
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
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
   Begin MSAdodcLib.Adodc AdoAir 
      Height          =   330
      Left            =   315
      Top             =   2625
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
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
      Caption         =   "Air"
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
   Begin MSAdodcLib.Adodc AdoLibroRet 
      Height          =   330
      Left            =   315
      Top             =   3360
      Width           =   4050
      _ExtentX        =   7144
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
      Caption         =   "LibroRet"
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   420
      Top             =   3885
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibroRet.frx":0D2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibroRet.frx":1044
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibroRet.frx":135E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibroRet.frx":1678
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibroRet.frx":1992
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FLibroRetenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SumatoriaB, SumatoriaS, SumatoriaBS As Double
'Dim SumatoriaB As Double
'Belenynrol@hotmail.com

Private Sub CAño_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CMes_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CMes_LostFocus()
  Fecha_Del_AT CMes, CAño
End Sub

Private Sub DGLibroRet_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGLibroRet.Visible = False
     GenerarDataTexto FLibroRetenciones, AdoLibroRet
     DGLibroRet.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyP Then
     MensajeEncabData = "RESUMEN DE ANEXOS TRANSACCIONALES EN " & UCaseStrg(SinEspaciosDer(LstAT))
     SQLMsg1 = "Mes de: " & CMes & " del año " & CAño
     Cuadricula = True
     DGLibroRet.Visible = False
     ImprimirAdoAT AdoLibroRet, True
     DGLibroRet.Visible = True
  End If
End Sub

Private Sub Form_Activate()
Dim AltoTab As Single
Dim AnchoTab As Single
Dim InicioTab As Single

   AnchoTab = Screen.width - 250
   AltoTab = Screen.Height - 3100
   InicioTab = DGLibroRet.Top
   Opcion = 0
   
   DGLibroRet.width = AnchoTab
   DGLibroRet.Height = AltoTab
   
   AdoLibroRet.Top = DGLibroRet.Height + DGLibroRet.Top
   AdoLibroRet.width = AnchoTab - AdoLibroRet.Left ' 250
   CmdSalir.Top = DGLibroRet.Height + DGLibroRet.Top
      
   LstAT.Clear
   LstAT.AddItem "1.- COMPRAS"
   LstAT.AddItem "2.- VENTAS"
   LstAT.AddItem "3.- IMPORTACIONES"
   LstAT.AddItem "4.- EXPORTACIONES"
   LstAT.AddItem "9.- RELACION DE DEPENDENCIA"
   LstAT.Text = "1.- COMPRAS"
   CMes.Clear
   CAño.Clear
   sSQL = Listar_Meses
   Select_Adodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        CMes.AddItem "Todos"
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
   
   sSQL = "SELECT Codigo,Count(Codigo) " _
        & "FROM Tipo_Concepto_Retencion " _
        & "WHERE Codigo <> '.' " _
        & "GROUP BY Codigo " _
        & "ORDER BY Codigo "
   Select_Adodc AdoAir, sSQL
   With AdoAir.Recordset
    If .RecordCount > 0 Then
        If Existe_Tabla("Asiento_Codigo_Air") Then
           NombreCampo = ""
           sSQL = "SELECT * " _
                & "FROM Asiento_Codigo_Air " _
                & "WHERE Tipo_Trans = '.' "
           Select_Adodc AdoAux, sSQL
           Do While Not .EOF
              Evaluar = False
              For J = 0 To AdoAux.Recordset.Fields.Count - 1
                 If "Cod_" & .Fields("Codigo") = AdoAux.Recordset.Fields(J).Name Then Evaluar = True
              Next J
              If Evaluar = False Then
                 SQL1 = "ALTER TABLE Asiento_Codigo_Air "
                 If SQL_Server Then
                    SQL1 = SQL1 & "ADD [Cod_" & .Fields("Codigo") & "] FLOAT NULL; "
                 Else
                    SQL1 = SQL1 & "ADD [Cod_" & .Fields("Codigo") & "] DOUBLE NULL; "
                 End If
                 Ejecutar_SQL_SP SQL1
              End If
             .MoveNext
           Loop
        Else
           SQL1 = "CREATE TABLE Asiento_Codigo_Air ("
           If SQL_Server Then
              SQL1 = SQL1 _
                   & "[Ln_No] SMALLINT NULL," _
                   & "[Tipo_Trans] NVARCHAR (1) NULL," _
                   & "[Codigo] NVARCHAR (10) NULL," _
                   & "[Fecha] SMALLDATETIME NULL,"
           Else
              SQL1 = SQL1 _
                   & "[Ln_No] SHORT NULL," _
                   & "[Tipo_Trans] TEXT (1) NULL," _
                   & "[Codigo] TEXT (10) NULL," _
                   & "[Fecha] DATETIME NULL,"
           End If
           Do While Not .EOF
              If SQL_Server Then
                 SQL1 = SQL1 & "Cod_" & .Fields("Codigo") & " FLOAT NULL,"
              Else
                 SQL1 = SQL1 & "Cod_" & .Fields("Codigo") & " DOUBLE NULL,"
              End If
             .MoveNext
           Loop
           If SQL_Server Then
              SQL1 = SQL1 _
                   & "Total_Ret MONEY NULL," _
                   & "[CodigoU] NVARCHAR (10) NULL," _
                   & "[Item] NVARCHAR (3) NULL); "
           Else
              SQL1 = SQL1 _
                   & "Total_Ret CURRENCY NULL," _
                   & "[CodigoU] TEXT (10) NULL," _
                   & "[Item] TEXT (3) NULL); "
           End If
           Ejecutar_SQL_SP SQL1
        End If
    End If
   End With
   CAño.SetFocus
   RatonNormal
End Sub

Private Sub Form_Load()
   ConectarAdodc AdoAux
   ConectarAdodc AdoAir
   ConectarAdodc AdoClientes
   ConectarAdodc AdoLibroRet
End Sub

Public Sub Llenar_AT()
Dim TipoAT As String
    Contador = 0
    sSQL = "DELETE * " _
         & "FROM Saldo_Diarios " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    TipoAT = MidStrg(LstAT.Text, 1, 3)
    Select Case TipoAT
      Case "1.-": sSQL = sSQL & "AND TP = 'ATCO' "
      Case "2.-": sSQL = sSQL & "AND TP = 'ATVE' "
      Case "3.-": sSQL = sSQL & "AND TP = 'ATIM' "
      Case "4.-": sSQL = sSQL & "AND TP = 'ATEX' "
    End Select
    Ejecutar_SQL_SP sSQL
   'Llenamos los datos de las Tablas Trans_Compras y Trans_Air dependiendo del Tipo de consulta
    Select Case TipoAT
      Case "1.-"
            'COMPRAS
            sSQL = "SELECT C.Cliente, C.Codigo, C.CI_RUC, C.TD,TC.* " _
                 & "FROM Trans_Compras As TC, Clientes As C " _
                 & "WHERE TC.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
                 & "AND TC.Periodo = '" & Periodo_Contable & "' "
            If ConSucursal = False Then sSQL = sSQL & "AND TC.Item = '" & NumEmpresa & "' "
            sSQL = sSQL _
                 & "AND TC.IdProv = C.Codigo  " _
                 & "ORDER BY TC.Linea_SRI, C.Cliente, C.CI_RUC, C.TD "
            Select_Adodc AdoAux, sSQL
            With AdoAux.Recordset
               If .RecordCount > 0 Then
                   Do While Not .EOF
                     'Encero las variables de las sumatorias de Bienes,Servicios y Totales
                     'Asigno a variables los campos
                      CodigoCliente = .Fields("IdProv")
                      Factura_No = .Fields("Secuencial")
                      Codigo1 = .Fields("Establecimiento")
                      Codigo2 = .Fields("PuntoEmision")
                      TipoCta = .Fields("TipoComprobante")
                      
                      SetAdoAddNew "Saldo_Diarios"
                      SetAdoFields "TC", TipoCta
                      SetAdoFields "CodigoC", CodigoCliente
                      SetAdoFields "Numero", Factura_No
                      SetAdoFields "Fecha", .Fields("Fecha")
                      SetAdoFields "Ingresos", .Fields("BaseImpGrav")
                      SetAdoFields "Egresos", .Fields("BaseImponible")
                      SetAdoFields "PEN", .Fields("MontoIva")
                      SetAdoFields "Ln", .Fields("Linea_SRI")
                      
                      Real1 = .Fields("ValorRetBienes")
                      Real2 = .Fields("ValorRetServicios")
                      Real3 = .Fields("MontoIce")
                                           
                     'Bienes: R_30
                      If .Fields("PorRetBienes") = "1" Then SetAdoFields "Enero", Real1
                     'Servicios: R_70
                      If .Fields("PorRetServicios") = "2" Then SetAdoFields "Febrero", Real2
                     'Bienes y Servicios: R_100
                      If .Fields("PorRetBienes") = "3" Or .Fields("PorRetServicios") = "3" Then
                          SetAdoFields "Marzo", (Real1 + Real2)
                      End If
                     'Valor retenido por ICE
                      SetAdoFields "Total", (Real1 + Real2)
                      SetAdoFields "Total_Mora", .Fields("BaseImpIce")
                      SetAdoFields "Saldo_Anterior", Real3
                      ValorUnit = 0
                      Real1 = 0
                      Real2 = 0
                      Real3 = 0
                      Real4 = 0
                      Real5 = 0
                      Real6 = 0
                      Real7 = 0
                      Real8 = 0
                      Real9 = 0
                     'Genero el AIR
                      sSQL = "SELECT * " _
                           & "FROM Trans_Air " _
                           & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
                           & "AND Tipo_Trans = 'C' "
                      If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
                      sSQL = sSQL _
                           & "AND Periodo = '" & Periodo_Contable & "' " _
                           & "AND IdProv = '" & CodigoCliente & "' " _
                           & "AND Factura_No = " & Factura_No & " " _
                           & "AND EstabFactura = '" & Codigo1 & "' " _
                           & "AND PuntoEmiFactura = '" & Codigo2 & "' " _
                           & "ORDER BY Linea_SRI,Fecha,CodRet "
                      Select_Adodc AdoAir, sSQL
                      If AdoAir.Recordset.RecordCount > 0 Then
                         Cta = AdoAir.Recordset.Fields("EstabRetencion") & AdoAir.Recordset.Fields("PtoEmiRetencion") & "-" _
                             & Format(AdoAir.Recordset.Fields("SecRetencion"), "00000000")
                         Do While Not AdoAir.Recordset.EOF
                           'Encero las variables
                           'Asigno a cada campo según el porcentaje que corresponda
                            Interes = Round(AdoAir.Recordset.Fields("Porcentaje") * 100, 2)
                            ValorUnit = ValorUnit + AdoAir.Recordset.Fields("ValRet")
                            Select Case Interes
                              Case 1
                                   Real1 = Real1 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Abril", Real1
                              Case 1.75
                                   Real2 = Real2 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Mayo", Real2
                              Case 2
                                   Real3 = Real3 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Junio", Real3
                              Case 2.75
                                   Real4 = Real4 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Julio", Real4
                              Case 5
                                   Real5 = Real5 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Agosto", Real5
                              Case 8
                                   Real6 = Real6 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Septiembre", Real6
                              Case 10
                                   Real7 = Real7 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Octubre", Real7
                              Case 15
                                   Real8 = Real8 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Noviembre", Real8
                              Case 25
                                   Real9 = Real9 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Diciembre", Real9
                            End Select
                            AdoAir.Recordset.MoveNext
                         Loop
                      End If
                      If ValorUnit <= 0 Then Cta = Ninguno
                      SetAdoFields "Cta", Cta
                      SetAdoFields "Saldo_Actual", ValorUnit
                      SetAdoFields "TP", "ATCO"
                      SetAdoUpdate
                     .MoveNext
                   Loop
                End If
            End With
    
     'Llenamos los datos de las Tablas Trans_Ventas y Trans_Air dependiendo del Tipo de consulta
      Case "2.-"
           'VENTAS
            sSQL = "SELECT * " _
                 & "FROM Trans_Ventas " _
                 & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
                 & "AND Periodo = '" & Periodo_Contable & "' "
            If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
            sSQL = sSQL _
                 & "ORDER BY Linea_SRI, Razon_Social "
            Select_Adodc AdoAux, sSQL
            With AdoAux.Recordset
               If .RecordCount > 0 Then
                   Do While Not .EOF
                     'Encero las variables de las sumatorias de Bienes,Servicios y Totales
                     'Asigno a variables los campos
                      CodigoCliente = .Fields("IdProv")
                      If .Fields("Secuencial") > 0 Then
                         Factura_No = .Fields("Secuencial")
                      Else
                         Factura_No = .Fields("NumeroComprobantes")
                      End If
                      
                      Codigo1 = .Fields("Establecimiento")
                      Codigo2 = .Fields("PuntoEmision")
                      TipoCta = .Fields("TipoComprobante")
                      
                      SetAdoAddNew "Saldo_Diarios"
                      SetAdoFields "TC", TipoCta
                      SetAdoFields "CodigoC", CodigoCliente
                      SetAdoFields "Comprobante", .Fields("Razon_Social")
                      SetAdoFields "Cta_Aux", .Fields("RUC_CI")
                      SetAdoFields "TB", .Fields("TB")
                      SetAdoFields "Numero", Factura_No
                      SetAdoFields "Fecha", .Fields("Fecha")
                      SetAdoFields "Ingresos", .Fields("BaseImpGrav")
                      SetAdoFields "Egresos", .Fields("BaseImponible")
                      SetAdoFields "PEN", .Fields("MontoIva")
                      SetAdoFields "Ln", .Fields("Linea_SRI")
                      
                      Real1 = .Fields("ValorRetBienes")
                      Real2 = .Fields("ValorRetServicios")
                      Real3 = .Fields("MontoIce")
                                           
                     'Bienes: R_30
                      If .Fields("PorRetBienes") = "1" Then SetAdoFields "Enero", Real1
                     'Servicios: R_70
                      If .Fields("PorRetServicios") = "2" Then SetAdoFields "Febrero", Real2
                     'Bienes y Servicios: R_100
                      If .Fields("PorRetBienes") = "3" Or .Fields("PorRetServicios") = "3" Then
                          SetAdoFields "Marzo", (Real1 + Real2)
                      End If
                     'Valor retenido por ICE
                      SetAdoFields "Total", (Real1 + Real2)
                      SetAdoFields "Total_Mora", .Fields("BaseImpIce")
                      SetAdoFields "Saldo_Anterior", Real3
                      ValorUnit = 0
                      Real1 = 0
                      Real2 = 0
                      Real3 = 0
                      Real4 = 0
                      Real5 = 0
                      Real6 = 0
                      Real7 = 0
                      Real8 = 0
                      Real9 = 0
                     'Genero el AIR
                      If Val(TipoCta) <> 4 Then
                         sSQL = "SELECT * " _
                              & "FROM Trans_Air " _
                              & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  " _
                              & "AND Tipo_Trans = 'V' "
                         If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
                         sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' " _
                              & "AND IdProv = '" & CodigoCliente & "' " _
                              & "AND Factura_No = " & Factura_No & " " _
                              & "AND EstabFactura = '" & Codigo1 & "' " _
                              & "AND PuntoEmiFactura = '" & Codigo2 & "' " _
                              & "ORDER BY Linea_SRI,Fecha,CodRet "
                         Select_Adodc AdoAir, sSQL
                         If AdoAir.Recordset.RecordCount > 0 Then
                            Cta = AdoAir.Recordset.Fields("EstabRetencion") _
                                & AdoAir.Recordset.Fields("PtoEmiRetencion") & "-" _
                                & Format(AdoAir.Recordset.Fields("SecRetencion"), "0000000")
                            Do While Not AdoAir.Recordset.EOF
                              'Encero las variables
                              'Asigno a cada campo según el porcentaje que corresponda
                               Interes = AdoAir.Recordset.Fields("Porcentaje")
                               ValorUnit = ValorUnit + AdoAir.Recordset.Fields("ValRet")
                            Select Case Interes
                              Case 1
                                   Real1 = Real1 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Abril", Real1
                              Case 1.75
                                   Real2 = Real2 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Mayo", Real2
                              Case 2
                                   Real3 = Real3 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Junio", Real3
                              Case 2.75
                                   Real4 = Real4 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Julio", Real4
                              Case 5
                                   Real5 = Real5 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Agosto", Real5
                              Case 8
                                   Real6 = Real6 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Septiembre", Real6
                              Case 10
                                   Real7 = Real7 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Octubre", Real7
                              Case 15
                                   Real8 = Real8 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Noviembre", Real8
                              Case 25
                                   Real9 = Real9 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Diciembre", Real9
                            End Select
                               AdoAir.Recordset.MoveNext
                            Loop
                         End If
                      End If
                      If ValorUnit <= 0 Then Cta = Ninguno
                      SetAdoFields "Cta", Cta
                      SetAdoFields "Saldo_Actual", ValorUnit
                      SetAdoFields "TP", "ATVE"
                      SetAdoUpdate
                     .MoveNext
                   Loop
                End If
            End With
    
      'Llenamos los datos de las Tablas Trans_Importaciones y Trans_Air dependiendo del Tipo de consulta
       Case "3.-"
           'IMPORTACIONES
            sSQL = "SELECT C.Cliente, C.Codigo, C.CI_RUC, C.TD,TI.* " _
                 & "FROM Trans_Importaciones As TI, Clientes As C " _
                 & "WHERE TI.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
                 & "AND TI.Periodo = '" & Periodo_Contable & "' "
            If ConSucursal = False Then sSQL = sSQL & "AND TI.Item = '" & NumEmpresa & "' "
            sSQL = sSQL & "AND TI.IdFiscalProv = C.Codigo  " _
                 & "ORDER BY TI.Linea_SRI,C.Cliente, C.CI_RUC, C.TD "
            Select_Adodc AdoAux, sSQL
            With AdoAux.Recordset
               If .RecordCount > 0 Then
                   Do While Not .EOF
                     'Encero las variables de las sumatorias de Bienes,Servicios y Totales
                     'Asigno a variables los campos
                      CodigoCliente = .Fields("IdFiscalProv")
                      Factura_No = .Fields("Correlativo")
                      'Codigo1 = .Fields("Establecimiento")
                      'Codigo2 = .Fields("PuntoEmision")
                      SetAdoAddNew "Saldo_Diarios"
                      SetAdoFields "CodigoC", CodigoCliente
                      SetAdoFields "Numero", Factura_No
                      SetAdoFields "Fecha", .Fields("Fecha")
                      SetAdoFields "Ingresos", .Fields("BaseImpGrav")
                      SetAdoFields "Egresos", .Fields("BaseImponible")
                      SetAdoFields "PEN", .Fields("MontoIva")
                      SetAdoFields "Ln", .Fields("Linea_SRI")
                     
                     'Valor CIF
                      SetAdoFields "Enero", .Fields("ValorCIF")
                      'Valor ICE
                      SetAdoFields "Febrero", .Fields("BaseImpIce")
                      'Monto ICE
                      SetAdoFields "Marzo", .Fields("MontoIce")
                      ValorUnit = 0
                      Real1 = 0
                      Real2 = 0
                      Real3 = 0
                      Real4 = 0
                      Real5 = 0
                      Real6 = 0
                      Real7 = 0
                      Real8 = 0
                      Real9 = 0
                     'Genero el AIR
                      sSQL = "SELECT * " _
                           & "FROM Trans_Air " _
                           & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
                           & "AND Tipo_Trans = 'I' "
                      If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
                      sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' " _
                           & "AND IdProv = '" & CodigoCliente & "' " _
                           & "AND Factura_No = " & Factura_No & " " _
                           & "ORDER BY Linea_SRI,Fecha,CodRet "
                      Select_Adodc AdoAir, sSQL
                      If AdoAir.Recordset.RecordCount > 0 Then
                         Cta = AdoAir.Recordset.Fields("EstabRetencion") _
                             & AdoAir.Recordset.Fields("PtoEmiRetencion") & "-" _
                             & Format(AdoAir.Recordset.Fields("SecRetencion"), "0000000")
                         Do While Not AdoAir.Recordset.EOF
                           'Encero las variables
                           'Asigno a cada campo según el porcentaje que corresponda
                            Interes = Round(AdoAir.Recordset.Fields("Porcentaje") * 100, 2)
                            ValorUnit = ValorUnit + AdoAir.Recordset.Fields("ValRet")
                            Select Case Interes
                              Case 1
                                   Real1 = Real1 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Abril", Real1
                              Case 1.75
                                   Real2 = Real2 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Mayo", Real2
                              Case 2
                                   Real3 = Real3 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Junio", Real3
                              Case 2.75
                                   Real4 = Real4 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Julio", Real4
                              Case 5
                                   Real5 = Real5 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Agosto", Real5
                              Case 8
                                   Real6 = Real6 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Septiembre", Real6
                              Case 10
                                   Real7 = Real7 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Octubre", Real7
                              Case 15
                                   Real8 = Real8 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Noviembre", Real8
                              Case 25
                                   Real9 = Real9 + AdoAir.Recordset.Fields("ValRet")
                                   SetAdoFields "Diciembre", Real9
                            End Select
                            AdoAir.Recordset.MoveNext
                         Loop
                      End If
                      If ValorUnit <= 0 Then Cta = Ninguno
                      SetAdoFields "Cta", Cta
                      SetAdoFields "Saldo_Actual", ValorUnit
                      SetAdoFields "TP", "ATIM"
                      SetAdoUpdate
                     .MoveNext
                   Loop
                End If
            End With
            
       'Llenamos los datos de las Tablas Trans_Exportaciones y Trans_Air dependiendo del Tipo de consulta
        Case "4.-"
            'EXPORTACIONES
            sSQL = "SELECT C.Cliente, C.Codigo, C.CI_RUC, C.TD,TE.* " _
                 & "FROM Trans_Exportaciones As TE, Clientes As C " _
                 & "WHERE TE.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
                 & "AND TE.Periodo = '" & Periodo_Contable & "' "
            If ConSucursal = False Then sSQL = sSQL & "AND TE.Item = '" & NumEmpresa & "' "
            sSQL = sSQL & "AND TE.IdFiscalProv = C.Codigo  " _
                 & "ORDER BY TE.Linea_SRI,C.Cliente, C.CI_RUC, C.TD "
            Select_Adodc AdoAux, sSQL
            With AdoAux.Recordset
               If .RecordCount > 0 Then
                   Do While Not .EOF
                     'Encero las variables de las sumatorias de Bienes,Servicios y Totales
                     'Asigno a variables los campos
                      CodigoCliente = .Fields("IdFiscalProv")
                      Factura_No = .Fields("Correlativo")
                      Codigo1 = .Fields("Establecimiento")
                      Codigo2 = .Fields("PuntoEmision")
                      SetAdoAddNew "Saldo_Diarios"
                      SetAdoFields "CodigoC", CodigoCliente
                      SetAdoFields "Numero", Factura_No
                      SetAdoFields "Fecha", .Fields("Fecha")
                      SetAdoFields "Ingresos", .Fields("ValorFOB")
                      SetAdoFields "Egresos", .Fields("ValorFOBComprobante")
                      SetAdoFields "Ln", .Fields("Linea_SRI")
                      SetAdoFields "TP", "ATEX"
                      SetAdoUpdate
                     .MoveNext
                   Loop
                End If
            End With
    End Select
End Sub

Public Sub Despliega_Compras()
    SQL1 = "SD.Numero,"
    sSQL = "SELECT SUM(Ingresos) AS Con_IVA, SUM(Egresos) As Sin_IVA, SUM(PEN) As IVA, SUM(Enero) As R_30, SUM(Febrero) As R_70, SUM(Marzo) As R_100, " _
         & "SUM(Total) As RET_IVA, SUM(Total_Mora) As ICE, SUM(Saldo_Anterior) As RET_ICE, SUM(Abril) As R_1, SUM(Mayo) As R_1_75, SUM(Junio) As R_2, " _
         & "SUM(Julio) As R_2_75, SUM(Agosto) As R_5, SUM(Septiembre) As R_8, SUM(Octubre) As R_10, SUM(Noviembre) As R_15, SUM(Diciembre) As R_25, " _
         & "SUM(Saldo_Actual) As Total_RET " _
         & "FROM Saldo_Diarios " _
         & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND TP = 'ATCO' " _
         & "GROUP BY TP "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
      If .RecordCount > 0 Then
         SetAdoAddNew "Saldo_Diarios"
         SetAdoFields "Ln", 9999
         SetAdoFields "CodigoC", ".."
         SetAdoFields "Numero", 0
         SetAdoFields "Fecha", FechaFinal
         SetAdoFields "Ingresos", .Fields("Con_IVA")
         SetAdoFields "Egresos", .Fields("Sin_IVA")
         SetAdoFields "PEN", .Fields("IVA")
         SetAdoFields "Enero", .Fields("R_30")
         SetAdoFields "Febrero", .Fields("R_70")
         SetAdoFields "Marzo", .Fields("R_100")
         SetAdoFields "Total", .Fields("RET_IVA")
         SetAdoFields "Total_Mora", .Fields("ICE")
         SetAdoFields "Saldo_Anterior", .Fields("RET_ICE")
         SetAdoFields "Abril", .Fields("R_1")
         SetAdoFields "Mayo", .Fields("R_1_75")
         SetAdoFields "Junio", .Fields("R_2")
         SetAdoFields "Julio", .Fields("R_2_75")
         SetAdoFields "Agosto", .Fields("R_5")
         SetAdoFields "Septiembre", .Fields("R_8")
         SetAdoFields "Octubre", .Fields("R_10")
         SetAdoFields "Noviembre", .Fields("R_15")
         SetAdoFields "Diciembre", .Fields("R_25")
         SetAdoFields "Saldo_Actual", .Fields("Total_RET")
         SetAdoFields "TP", "ATCO"
         SetAdoUpdate
         
        'Trama de los subtotales de las Retenciones en Ventas
         If .Fields("Con_IVA") > 0 Then SQL1 = SQL1 & "SD.Ingresos As Con_IVA,"
         If .Fields("Sin_IVA") > 0 Then SQL1 = SQL1 & "SD.Egresos As Sin_IVA,"
         If .Fields("IVA") > 0 Then SQL1 = SQL1 & "SD.PEN As IVA,"
         If .Fields("R_30") > 0 Then SQL1 = SQL1 & "SD.Enero As R_30,"
         If .Fields("R_70") > 0 Then SQL1 = SQL1 & "SD.Febrero As R_70,"
         If .Fields("R_100") > 0 Then SQL1 = SQL1 & "SD.Marzo As R_100,"
         If .Fields("RET_IVA") > 0 Then SQL1 = SQL1 & "SD.Total As RET_IVA,"
         If .Fields("ICE") > 0 Then SQL1 = SQL1 & "SD.Total_Mora As ICE,"
         If .Fields("RET_ICE") > 0 Then SQL1 = SQL1 & "SD.Saldo_Anterior As RET_ICE,"
         If .Fields("R_1") > 0 Then SQL1 = SQL1 & "SD.Abril As R_1,"
         If .Fields("R_1_75") > 0 Then SQL1 = SQL1 & "SD.Mayo As R_1_75,"
         If .Fields("R_2") > 0 Then SQL1 = SQL1 & "SD.Junio As R_2,"
         If .Fields("R_2_75") > 0 Then SQL1 = SQL1 & "SD.Julio As R_2_75,"
         If .Fields("R_5") > 0 Then SQL1 = SQL1 & "SD.Agosto As R_5,"
         If .Fields("R_8") > 0 Then SQL1 = SQL1 & "SD.Septiembre As R_8,"
         If .Fields("R_10") > 0 Then SQL1 = SQL1 & "SD.Octubre As R_10,"
         If .Fields("R_15") > 0 Then SQL1 = SQL1 & "SD.Noviembre As R_15,"
         If .Fields("R_25") > 0 Then SQL1 = SQL1 & "SD.Diciembre As R_25,"
         If .Fields("Total_RET") > 0 Then SQL1 = SQL1 & "SD.Saldo_Actual As Total_RET,"
      End If
    End With
    sSQL = "SELECT SD.Ln,C.Especial As CE,C.RISE,C.Cliente As Razon_Social,C.CI_RUC,C.TD,SD.TC,SD.Fecha," _
         & SQL1 & "SD.Cta As Retencion_No " _
         & "FROM Saldo_Diarios As SD, Clientes As C " _
         & "WHERE SD.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  " _
         & "AND SD.Item = '" & NumEmpresa & "' " _
         & "AND SD.CodigoU = '" & CodigoUsuario & "' " _
         & "AND SD.TP = 'ATCO' " _
         & "AND SD.CodigoC = C.Codigo " _
         & "ORDER BY SD.Ln,SD.PEN,SD.Fecha "
    Select_Adodc_Grid DGLibroRet, AdoLibroRet, sSQL
End Sub

Public Sub Despliega_Ventas()
    SQL1 = "Numero,"
    sSQL = "SELECT SUM(Ingresos) AS Con_IVA, SUM(Egresos) As Sin_IVA, SUM(PEN) As IVA, SUM(Enero) As R_30, SUM(Febrero) As R_70, SUM(Marzo) As R_100, " _
         & "SUM(Total) As RET_IVA, SUM(Total_Mora) As ICE, SUM(Saldo_Anterior) As RET_ICE, SUM(Abril) As R_1, SUM(Mayo) As R_1_75, SUM(Junio) As R_2, " _
         & "SUM(Julio) As R_2_75, SUM(Agosto) As R_5, SUM(Septiembre) As R_8, SUM(Octubre) As R_10, SUM(Noviembre) As R_15, SUM(Diciembre) As R_25, " _
         & "SUM(Saldo_Actual) As Total_RET " _
         & "FROM Saldo_Diarios " _
         & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND TP = 'ATVE' " _
         & "GROUP BY TP "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
      If .RecordCount > 0 Then
         SetAdoAddNew "Saldo_Diarios"
         SetAdoFields "Ln", 9999
         SetAdoFields "CodigoC", ".."
         SetAdoFields "Numero", 0
         SetAdoFields "Fecha", FechaFinal
         SetAdoFields "Ingresos", .Fields("Con_IVA")
         SetAdoFields "Egresos", .Fields("Sin_IVA")
         SetAdoFields "PEN", .Fields("IVA")
         SetAdoFields "Enero", .Fields("R_30")
         SetAdoFields "Febrero", .Fields("R_70")
         SetAdoFields "Marzo", .Fields("R_100")
         SetAdoFields "Total", .Fields("RET_IVA")
         SetAdoFields "Total_Mora", .Fields("ICE")
         SetAdoFields "Saldo_Anterior", .Fields("RET_ICE")
         SetAdoFields "Abril", .Fields("R_1")
         SetAdoFields "Mayo", .Fields("R_1_75")
         SetAdoFields "Junio", .Fields("R_2")
         SetAdoFields "Julio", .Fields("R_2_75")
         SetAdoFields "Agosto", .Fields("R_5")
         SetAdoFields "Septiembre", .Fields("R_8")
         SetAdoFields "Octubre", .Fields("R_10")
         SetAdoFields "Noviembre", .Fields("R_15")
         SetAdoFields "Diciembre", .Fields("R_25")
         SetAdoFields "Saldo_Actual", .Fields("Total_RET")
         SetAdoFields "TP", "ATVE"
         SetAdoUpdate
         
        'Trama de los subtotales de las Retenciones en Ventas
         If .Fields("Con_IVA") > 0 Then SQL1 = SQL1 & "Ingresos As Con_IVA,"
         If .Fields("Sin_IVA") > 0 Then SQL1 = SQL1 & "Egresos As Sin_IVA,"
         If .Fields("IVA") > 0 Then SQL1 = SQL1 & "PEN As IVA,"
         If .Fields("R_30") > 0 Then SQL1 = SQL1 & "Enero As R_30,"
         If .Fields("R_70") > 0 Then SQL1 = SQL1 & "Febrero As R_70,"
         If .Fields("R_100") > 0 Then SQL1 = SQL1 & "Marzo As R_100,"
         If .Fields("RET_IVA") > 0 Then SQL1 = SQL1 & "Total As RET_IVA,"
         If .Fields("ICE") > 0 Then SQL1 = SQL1 & "Total_Mora As ICE,"
         If .Fields("RET_ICE") > 0 Then SQL1 = SQL1 & "Saldo_Anterior As RET_ICE,"
         If .Fields("R_1") > 0 Then SQL1 = SQL1 & "Abril As R_1,"
         If .Fields("R_1_75") > 0 Then SQL1 = SQL1 & "Mayo As R_1_75,"
         If .Fields("R_2") > 0 Then SQL1 = SQL1 & "Junio As R_2,"
         If .Fields("R_2_75") > 0 Then SQL1 = SQL1 & "Julio As R_2_75,"
         If .Fields("R_5") > 0 Then SQL1 = SQL1 & "Agosto As R_5,"
         If .Fields("R_8") > 0 Then SQL1 = SQL1 & "Septiembre As R_8,"
         If .Fields("R_10") > 0 Then SQL1 = SQL1 & "Octubre As R_10,"
         If .Fields("R_15") > 0 Then SQL1 = SQL1 & "Noviembre As R_15,"
         If .Fields("R_25") > 0 Then SQL1 = SQL1 & "Diciembre As R_25,"
         If .Fields("Total_RET") > 0 Then SQL1 = SQL1 & "Saldo_Actual As Total_RET,"
      End If
    End With
    sSQL = "SELECT Ln,Comprobante As Razon_Social,Cta_Aux As CI_RUC,TB As TD,Fecha,TC," _
         & SQL1 & "Cta As Retencion_No " _
         & "FROM Saldo_Diarios " _
         & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND TP = 'ATVE' " _
         & "ORDER BY Ln,PEN,Fecha "
    Select_Adodc_Grid DGLibroRet, AdoLibroRet, sSQL
    
'''& "Ingresos As Con_IVA,Egresos As Sin_IVA,PEN As Valor_IVA," _
'''         & "ENE As R_30,FEB As R_70,MAR As R_100,Total As RET_IVA,Total_Mora As ICE,NOV As RET_ICE," _
'''         & "ABR As R_1,MAY As R_2,JUN As R_5,JUL As R_8,AGO As R_10,SEP As R_15,OCT As R_25,
End Sub

Public Sub Despliega_Importaciones()
    sSQL = "SELECT SUM(Ingresos) AS Con_IVA,SUM(Egresos) As Sin_IVA,SUM(PEN) As Valor_IVA,SUM(Enero) As Valor_CIF,SUM(Febrero) As ICE,SUM(Marzo) As RET_ICE," _
         & "SUM(Abril) As R_1,SUM(Mayo) As R_2,SUM(Junio) As R_5,SUM(Julio) As R_8,SUM(Agosto) As R_10, SUM(Septiembre) As R_15, SUM(Octubre) As R_25,SUM(Saldo_Actual) As Total_RET " _
         & "FROM Saldo_Diarios " _
         & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND TP = 'ATIM' " _
         & "GROUP BY TP "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
      If .RecordCount > 0 Then
         SetAdoAddNew "Saldo_Diarios"
         SetAdoFields "CodigoC", ".."
         SetAdoFields "Numero", 0
         SetAdoFields "Fecha", FechaFinal
         SetAdoFields "Ingresos", .Fields("Con_IVA")
         SetAdoFields "Egresos", .Fields("Sin_IVA")
         SetAdoFields "PEN", .Fields("Valor_IVA")
         SetAdoFields "Ln", 9999
         SetAdoFields "Enero", .Fields("Valor_CIF")
         SetAdoFields "Febrero", .Fields("ICE")
         SetAdoFields "Marzo", .Fields("RET_ICE")
         SetAdoFields "Abril", .Fields("R_1")
         SetAdoFields "Mayo", .Fields("R_2")
         SetAdoFields "Junio", .Fields("R_5")
         SetAdoFields "Julio", .Fields("R_8")
         SetAdoFields "Agosto", .Fields("R_15")
         SetAdoFields "Septiembre", .Fields("R_25")
         SetAdoFields "Saldo_Actual", .Fields("Total_RET")
         SetAdoFields "TP", "ATIM"
         SetAdoUpdate
      End If
    End With
    sSQL = "SELECT SD.Ln,C.Especial As C_E,C.RISE,Cliente,SD.Fecha,Numero,Ingresos As Con_IVA,Egresos As Sin_IVA,PEN As Valor_IVA," _
         & "Enero As Valor_CIF,Febrero As ICE,Marzo As RET_ICE,Cta As Retencion_No," _
         & "Abril As R_1,Mayo As R_2,Junio As R_5,Julio As R_8,Agosto As R_10, Septiembre As R_15, Octubre As R_25,Saldo_Actual As Total_RET " _
         & "FROM Saldo_Diarios As SD, Clientes As C " _
         & "WHERE SD.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  " _
         & "AND SD.Item = '" & NumEmpresa & "' " _
         & "AND SD.CodigoU = '" & CodigoUsuario & "' " _
         & "AND SD.TP = 'ATIM' " _
         & "AND SD.CodigoC = C.Codigo " _
         & "ORDER BY SD.Ln,SD.PEN,SD.Fecha "
    Select_Adodc_Grid DGLibroRet, AdoLibroRet, sSQL
End Sub

Public Sub Despliega_Exportaciones()
    sSQL = "SELECT SUM(Ingresos) As Valor_FOB,SUM(Egresos) As Valor_FOBComprob " _
         & "FROM Saldo_Diarios " _
         & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND TP = 'ATEX' " _
         & "GROUP BY TP "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
      If .RecordCount > 0 Then
         SetAdoAddNew "Saldo_Diarios"
         SetAdoFields "CodigoC", ".."
         SetAdoFields "Numero", 0
         SetAdoFields "Fecha", FechaFinal
         SetAdoFields "Ingresos", .Fields("Valor_FOB")
         SetAdoFields "Egresos", .Fields("Valor_FOBComprob")
         SetAdoFields "Ln", 9999
         SetAdoFields "TP", "ATEX"
         SetAdoUpdate
      End If
    End With
    sSQL = "SELECT SD.Ln,C.Especial As C_E,C.RISE,Cliente,SD.Fecha,Numero,Ingresos As Valor_FOB,Egresos As Valor_FOBComprob " _
         & "FROM Saldo_Diarios As SD, Clientes As C " _
         & "WHERE SD.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND SD.Item = '" & NumEmpresa & "' " _
         & "AND SD.CodigoU = '" & CodigoUsuario & "' " _
         & "AND SD.TP = 'ATEX' " _
         & "AND SD.CodigoC = C.Codigo " _
         & "ORDER BY SD.Ln,SD.Fecha "
    Select_Adodc_Grid DGLibroRet, AdoLibroRet, sSQL
End Sub

Public Sub Llenar_Air_AT(TipoTrans As String)
Dim TipoAT As String
Dim TotCodRet() As Campos_Tabla
Dim TotCod As Integer
  Contador = 0
  sSQL = "DELETE * " _
       & "FROM Asiento_Codigo_Air " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND Tipo_Trans = '" & TipoTrans & "' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "SELECT * " _
       & "FROM Asiento_Codigo_Air " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND Tipo_Trans = '" & TipoTrans & "' "
  Select_Adodc AdoAux, sSQL
  TotCod = AdoAux.Recordset.Fields.Count
  ReDim TotCodRet(TotCod) As Campos_Tabla
  For J = 0 To TotCod - 1
      TotCodRet(J).Campo = AdoAux.Recordset.Fields(J).Name
      TotCodRet(J).Valor = 0
  Next J

  sSQL = "SELECT TA.*,C.Cliente " _
       & "FROM Trans_Air As TA,Clientes As C " _
       & "WHERE TA.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND TA.Tipo_Trans = '" & TipoTrans & "' "
  If ConSucursal = False Then sSQL = sSQL & "AND TA.Item = '" & NumEmpresa & "' "
  sSQL = sSQL & "AND TA.IdProv = C.Codigo " _
       & "ORDER BY C.Cliente,TA.CodRet "
  Select_Adodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       CodigoCliente = .Fields("IdProv")
       SubTotal = 0
       Do While Not .EOF
          If CodigoCliente <> .Fields("IdProv") Then
             SetAdoAddNew "Asiento_Codigo_Air"
             Total = 0
             For J = 4 To TotCod - 1
                 SetAdoFields TotCodRet(J).Campo, TotCodRet(J).Valor
                 Total = Total + TotCodRet(J).Valor
                 TotCodRet(J).Valor = 0
             Next J
             SetAdoFields "Base_Imponible", SubTotal
             SetAdoFields "Total_Ret", Total
             SetAdoFields "Codigo", CodigoCliente
             SetAdoFields "Fecha", FechaMitad
             SetAdoFields "Tipo_Trans", TipoTrans
             SetAdoFields "CodigoU", CodigoUsuario
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "Ln_No", Contador
             SetAdoUpdate
             Contador = Contador + 1
             CodigoCliente = .Fields("IdProv")
             SubTotal = 0
          End If
          For J = 0 To TotCod
             If "Cod_" & .Fields("CodRet") = TotCodRet(J).Campo Then
                TotCodRet(J).Valor = TotCodRet(J).Valor + .Fields("ValRet")
             End If
          Next J
          SubTotal = SubTotal + .Fields("BaseImp")
         .MoveNext
       Loop
       'MsgBox CodigoCliente
       SetAdoAddNew "Asiento_Codigo_Air"
       Total = 0
       For J = 4 To TotCod - 1
           SetAdoFields TotCodRet(J).Campo, TotCodRet(J).Valor
           Total = Total + TotCodRet(J).Valor
       Next J
       SetAdoFields "Base_Imponible", SubTotal
       SetAdoFields "Total_Ret", Total
       SetAdoFields "Codigo", CodigoCliente
       SetAdoFields "Fecha", FechaMitad
       SetAdoFields "Tipo_Trans", TipoTrans
       SetAdoFields "CodigoU", CodigoUsuario
       SetAdoFields "Item", NumEmpresa
       SetAdoFields "Ln_No", Contador
       SetAdoUpdate
       Contador = Contador + 1
   End If
  End With
 'Sacamos Totales
  sSQL = "SELECT Tipo_Trans,"
  For J = 0 To TotCod - 1
      If MidStrg(TotCodRet(J).Campo, 1, 4) = "Cod_" Then sSQL = sSQL & "SUM(" & TotCodRet(J).Campo & ") As T" & TotCodRet(J).Campo & ", "
  Next J
  sSQL = sSQL _
       & "SUM(Total_Ret) As TTtoal_Ret " _
       & "FROM Asiento_Codigo_Air " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND Tipo_Trans = '" & TipoTrans & "' " _
       & "GROUP BY Tipo_Trans " _
       & "ORDER BY Tipo_Trans "
  Select_Adodc AdoAux, sSQL
 'Insertamos el gran total
  Total = 0
  SetAdoAddNew "Asiento_Codigo_Air"
  sSQL = "SELECT C.Cliente,TA.Base_Imponible,"
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       For J = 0 To .Fields.Count - 1
           If ((MidStrg(.Fields(J).Name, 1, 5) = "TCod_") And (.Fields(J) > 0)) Then
               sSQL = sSQL & "TA." & MidStrg(.Fields(J).Name, 2, Len(.Fields(J).Name)) & ","
               SetAdoFields MidStrg(.Fields(J).Name, 2, Len(.Fields(J).Name)), .Fields(J)
               Total = Total + .Fields(J)
           End If
       Next J
   End If
  End With
  SetAdoFields "Base_Imponible", 0
  SetAdoFields "Total_Ret", Total
  SetAdoFields "Codigo", ".."
  SetAdoFields "Fecha", FechaMitad
  SetAdoFields "Tipo_Trans", TipoTrans
  SetAdoFields "CodigoU", CodigoUsuario
  SetAdoFields "Item", NumEmpresa
  SetAdoFields "Ln_No", 9999
  SetAdoUpdate
  sSQL = sSQL _
       & "TA.Total_Ret " _
       & "FROM Asiento_Codigo_Air As TA, Clientes As C " _
       & "WHERE TA.Item = '" & NumEmpresa & "' " _
       & "AND TA.CodigoU = '" & CodigoUsuario & "' " _
       & "AND TA.Tipo_Trans = '" & TipoTrans & "' " _
       & "AND TA.Codigo = C.Codigo " _
       & "ORDER BY TA.Ln_No,C.Cliente "
  Select_Adodc_Grid DGLibroRet, AdoLibroRet, sSQL
End Sub

Public Sub Llenar_TotAir_AT(TipoTrans As String)
 sSQL = "SELECT TA.CodRet As Codigo,CCR.Concepto As Concepto_Retencion,COUNT(Concepto) As Cant_Reg," _
      & "SUM(BaseImp) As Base_Imponible,SUM(ValRet) As Valor_Retenido " _
      & "FROM Trans_Air As TA,Tipo_Concepto_Retencion As CCR " _
      & "WHERE TA.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
 If ConSucursal = False Then sSQL = sSQL & "AND TA.Item = '" & NumEmpresa & "' "
 sSQL = sSQL & "AND TA.Periodo = '" & Periodo_Contable & "' " _
      & "AND TA.Tipo_Trans = '" & TipoTrans & "' " _
      & "AND CCR.Fecha_Inicio <= #" & FechaMid & "# " _
      & "AND CCR.Fecha_Final >= #" & FechaMid & "# " _
      & "AND TA.CodRet = CCR.Codigo  " _
      & "GROUP BY TA.CodRet,CCR.Concepto " _
      & "UNION " _
      & "SELECT '999' As Codigo,'TOTALES DE CONCEPTOS DE RETENCIONES' As Concepto_Retencion,COUNT(*) As Cant_Reg," _
      & "SUM(BaseImp) As Base_Imponible,SUM(ValRet) As Valor_Retenido " _
      & "FROM Trans_Air As TA,Tipo_Concepto_Retencion As CCR " _
      & "WHERE TA.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
 If ConSucursal = False Then sSQL = sSQL & "AND TA.Item = '" & NumEmpresa & "' "
 sSQL = sSQL & "AND TA.Periodo = '" & Periodo_Contable & "' " _
      & "AND TA.Tipo_Trans = '" & TipoTrans & "' " _
      & "AND CCR.Fecha_Inicio <= #" & FechaMid & "# " _
      & "AND CCR.Fecha_Final >= #" & FechaMid & "# " _
      & "AND TA.CodRet = CCR.Codigo  " _
      & "GROUP BY TA.Tipo_Trans " _
      & "ORDER BY Codigo "
 Select_Adodc_Grid DGLibroRet, AdoLibroRet, sSQL
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim No_Mes As Byte
    Fecha_Del_AT CMes, CAño
    Select Case Button.key
      Case "Salir"
           Unload Me
      Case "Retenciones"
           Opcion = 1
           MensajeEncabData = "RESUMEN DE ANEXOS TRANSACCIONALES EN " & UCaseStrg(SinEspaciosDer(LstAT))
           DGLibroRet.Caption = MensajeEncabData
           Select Case MidStrg(LstAT.Text, 1, 3)
             Case "1.-"
                  Llenar_AT
                  Despliega_Compras
             Case "2.-"
                  Llenar_AT
                  Despliega_Ventas
             Case "3.-"
                  Llenar_AT
                  Despliega_Importaciones
             Case "4.-"
                  Llenar_AT
                  Despliega_Exportaciones
           End Select
      Case "Por_Codigo"
           Opcion = 2
           MensajeEncabData = "RESUMEN DE ANEXOS TRANSACCIONALES POR CODIGOS EN " & UCaseStrg(SinEspaciosDer(LstAT))
           DGLibroRet.Caption = MensajeEncabData
           Select Case MidStrg(LstAT.Text, 1, 3)
             Case "1.-"
                  Llenar_Air_AT "C"
             Case "2.-"
                  Llenar_Air_AT "V"
             Case "3.-"
                  Llenar_Air_AT "I"
             Case "4.-"
                  Llenar_Air_AT "E"
           End Select
      Case "Por_Detalle"
           Opcion = 3
           MensajeEncabData = "RESUMEN DE ANEXOS TRANSACCIONALES POR DETALLES EN " & UCaseStrg(SinEspaciosDer(LstAT))
           DGLibroRet.Caption = MensajeEncabData
           Select Case MidStrg(LstAT.Text, 1, 3)
             Case "1.-"
                  Llenar_TotAir_AT "C"
             Case "2.-"
                  Llenar_TotAir_AT "V"
             Case "3.-"
                  Llenar_TotAir_AT "I"
             Case "4.-"
                  Llenar_TotAir_AT "E"
           End Select
      Case "Print_Retenciones"
           If Opcion = 1 Then
              SQLMsg1 = "Mes de: " & CMes & " del año " & CAño
              Cuadricula = True
              DGLibroRet.Visible = False
              ImprimirAdoAT AdoLibroRet, True
              DGLibroRet.Visible = True
           ElseIf Opcion = 2 Then
              SQLMsg1 = "CODIGOS DE RETENCION DE " & UCaseStrg(SinEspaciosDer(LstAT))
              SQLMsg2 = "Mes de: " & CMes & " del año " & CAño
              Cuadricula = True
              DGLibroRet.Visible = False
              ImprimirAdodc AdoLibroRet, 1, 7
              DGLibroRet.Visible = True
           ElseIf Opcion = 3 Then
              SQLMsg1 = "CODIGOS DE RETENCION DE " & UCaseStrg(SinEspaciosDer(LstAT))
              SQLMsg2 = "Mes de: " & CMes & " del año " & CAño
              Cuadricula = True
              DGLibroRet.Visible = False
              ImprimirAdodc AdoLibroRet, 1, 7
              DGLibroRet.Visible = True
           Else
              MsgBox "No ha seleccionado ningún proceso"
           End If
    End Select
End Sub
