VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FCentroCostos 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingresar Subcuentas de Proceso"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "&Cancelar"
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
      Left            =   6300
      Picture         =   "CentroCosto.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1155
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "&Aceptar"
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
      Left            =   6300
      Picture         =   "CentroCosto.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
      Width           =   1170
   End
   Begin MSDataGridLib.DataGrid DGSubCta 
      Bindings        =   "CentroCosto.frx":0D0C
      Height          =   6525
      Left            =   105
      TabIndex        =   0
      Top             =   945
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   11509
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16777152
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   6
      AllowDelete     =   -1  'True
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
   Begin MSAdodcLib.Adodc AdoSubCtaDet1 
      Height          =   330
      Left            =   105
      Top             =   7665
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "SubCtaDet1"
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
   Begin MSDataListLib.DataCombo DCSubModulos 
      Bindings        =   "CentroCosto.frx":0D28
      DataSource      =   "AdoSubModulo"
      Height          =   315
      Left            =   105
      TabIndex        =   5
      Top             =   525
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12648384
      Text            =   "Ctas"
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
   Begin VB.Label LabelTotalSCMN 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   4095
      TabIndex        =   2
      Top             =   7665
      Width           =   1800
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
      TabIndex        =   6
      Top             =   105
      Width           =   6105
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3045
      TabIndex        =   1
      Top             =   7665
      Width           =   1065
   End
End
Attribute VB_Name = "FCentroCostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    sSQL = "DELETE * " _
         & "FROM Asiento_SC " _
         & "WHERE TC = '" & SubCta & "' " _
         & "AND Cta = '" & SubCtaGen & "' " _
         & "AND DH = '" & OpcDH & "' " _
         & "AND TM = '" & OpcTM & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND Valor = 0 "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Asiento_SC " _
         & "SET Valor = ROUND(Valor,2,0) " _
         & "WHERE TC = '" & SubCta & "' " _
         & "AND Cta = '" & SubCtaGen & "' " _
         & "AND DH = '" & OpcDH & "' " _
         & "AND TM = '" & OpcTM & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Ejecutar_SQL_SP sSQL
        
    SumatoriaSC = Sumatoria_CC
    
    Unload FCentroCostos
End Sub

Private Sub Command2_Click()
    sSQL = "DELETE * " _
         & "FROM Asiento_SC " _
         & "WHERE TC = '" & SubCta & "' " _
         & "AND Cta = '" & SubCtaGen & "' " _
         & "AND DH = '" & OpcDH & "' " _
         & "AND TM = '" & OpcTM & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Ejecutar_SQL_SP sSQL

    CodigoCC = Ninguno
    Unload FCentroCostos
End Sub

Private Sub DGSubCta_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
     SumatoriaSC = Sumatoria_CC
     LabelTotalSCMN.Caption = Format(SumatoriaSC, "#,##0.00")
     Command1.SetFocus
  End If
  
  If KeyCode = vbKeyReturn Then
     If AdoSubCtaDet1.Recordset.RecordCount > 0 Then
        AdoSubCtaDet1.Recordset.MoveNext
        If AdoSubCtaDet1.Recordset.EOF Then AdoSubCtaDet1.Recordset.MoveFirst
     End If
  End If
End Sub

Private Sub Form_Activate()
    RatonReloj
    CodigoCC = Ninguno
    LnSC_No = 1
    FCentroCostos.Caption = "CENTRO DE COSTOS PARA: " & SubCtaGen & " - " & Cuenta
    sSQL = "UPDATE Catalogo_SubCtas " _
         & "SET X = '.' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Catalogo_SubCtas " _
         & "SET X = 'I' " _
         & "FROM Catalogo_SubCtas As CS, Asiento_SC As A " _
         & "WHERE CS.Item = '" & NumEmpresa & "' " _
         & "AND CS.Periodo = '" & Periodo_Contable & "' " _
         & "AND CS.TC = 'CC' " _
         & "AND A.Cta = '" & SubCtaGen & "' " _
         & "AND A.DH = '" & OpcDH & "' " _
         & "AND A.TM = '" & OpcTM & "' " _
         & "AND A.T_No = " & Trans_No & " " _
         & "AND A.CodigoU = '" & CodigoUsuario & "' " _
         & "AND CS.Item = A.Item " _
         & "AND CS.Codigo = A.Codigo " _
         & "AND CS.TC = A.TC "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "DELETE * " _
         & "FROM Asiento_SC " _
         & "WHERE TC = '" & SubCta & "' " _
         & "AND Cta = '" & SubCtaGen & "' " _
         & "AND DH = '" & OpcDH & "' " _
         & "AND TM = '" & OpcTM & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    'Ejecutar_SQL_SP sSQL
    
    sSQL = "INSERT INTO Asiento_SC (Codigo, Beneficiario, FECHA_V, TC, TM, DH, Cta, T_No, Item, CodigoU, Detalle_SubCta, Factura, Prima, Valor, Valor_ME, SC_No, Fecha_D, Fecha_H, Bloquear, Serie) " _
         & "SELECT Codigo, Detalle, #" & BuscarFecha(Co.Fecha) & "#, 'CC' , '" & OpcTM & "', '" & OpcDH & "', '" _
         & SubCtaGen & "', " & Trans_No & ", '" & NumEmpresa & "', '" & CodigoUsuario & "', '.', 0, 0, 0, 0, 0, Fecha_D, Fecha_H, Bloquear, '001001' " _
         & "FROM Catalogo_SubCtas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'CC' " _
         & "AND X = '.' " _
         & "AND Agrupacion = 0 " _
         & "ORDER BY Codigo, Detalle "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "DELETE * " _
         & "FROM Asiento_SC " _
         & "WHERE TC = '" & SubCta & "' " _
         & "AND Cta = '" & SubCtaGen & "' " _
         & "AND DH = '" & OpcDH & "' " _
         & "AND TM = '" & OpcTM & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND Bloquear <> 0 " _
         & "AND Fecha_D <= #" & BuscarFecha(Co.Fecha) & "# " _
         & "AND Fecha_H >= #" & BuscarFecha(Co.Fecha) & "# "
    Ejecutar_SQL_SP sSQL
    
    SumatoriaSC = Sumatoria_CC
    
    RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FCentroCostos
  ConectarAdodc AdoSubCtaDet1
End Sub

Public Function Sumatoria_CC() As Currency
Dim SumaCC As Currency
    SumaCC = 0
    sSQL = "SELECT Codigo, Beneficiario, Valor, DH, TC, Cta, TM, T_No, SC_No, Item, CodigoU " _
         & "FROM Asiento_SC " _
         & "WHERE TC = '" & SubCta & "' " _
         & "AND Cta = '" & SubCtaGen & "' " _
         & "AND DH = '" & OpcDH & "' " _
         & "AND TM = '" & OpcTM & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "ORDER BY Codigo "
    Select_Adodc_Grid DGSubCta, AdoSubCtaDet1, sSQL, , True
    With AdoSubCtaDet1.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            SumaCC = SumaCC + .fields("Valor")
           .MoveNext
         Loop
     End If
    End With
    Sumatoria_CC = SumaCC
End Function

