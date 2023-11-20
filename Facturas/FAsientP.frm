VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form AsientoAuto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catalogo de Rol de Pagos"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   420
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
      Bindings        =   "FAsientP.frx":0000
      Height          =   5160
      Left            =   105
      TabIndex        =   12
      Top             =   1365
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   9102
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
      Top             =   6615
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
      Height          =   750
      Left            =   8610
      Picture         =   "FAsientP.frx":0019
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
      Left            =   9660
      Picture         =   "FAsientP.frx":045B
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Procesar"
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
      Left            =   7560
      Picture         =   "FAsientP.frx":0D25
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoSubCta 
      Height          =   330
      Left            =   315
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
      Left            =   315
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
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   1365
      TabIndex        =   2
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   420
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
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Periodo de Traslado"
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
      Width           =   2535
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
      Height          =   330
      Left            =   105
      TabIndex        =   11
      Top             =   945
      Width           =   10515
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
      Left            =   8505
      TabIndex        =   6
      Top             =   6615
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
      Left            =   6720
      TabIndex        =   7
      Top             =   6615
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
      TabIndex        =   8
      Top             =   6615
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
      Left            =   5670
      TabIndex        =   10
      Top             =   6615
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
      TabIndex        =   9
      Top             =   6615
      Width           =   1170
   End
End
Attribute VB_Name = "AsientoAuto"
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
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  SumaDebe = 0: SumaHaber = 0
  If FechaIni = FechaFin Then
     LblConcepto.Caption = "Traslado de Ventas Anticipadas del " & MBFechaI.Text
  Else
     LblConcepto.Caption = "Traslado de Ventas Anticipadas del " & MBFechaI.Text _
                         & " al " & MBFechaF.Text
  End If
  SQL2 = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Select_Adodc_Grid DGAsiento, AdoAsiento, SQL2
  RatonReloj
  sSQL = "SELECT P.Cta,CP.Cta_Venta_Anticipada,P.Pagos,P.Pagos * COUNT(Contrato_No) As VentaAnt " _
        & "FROM Prestamos As P,Trans_Suscripciones As TS,Catalogo_Productos As CP " _
        & "WHERE TS.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND TS.Item = '" & NumEmpresa & "' " _
        & "AND TS.AC = " & Val(adFalse) & " " _
        & "AND CP.Periodo = '" & Periodo_Contable & "' " _
        & "AND TS.Item = CP.Item " _
        & "AND TS.Item = P.Item " _
        & "AND TS.Contrato_No = P.Credito_No " _
        & "AND P.Cta = CP.Cta_Ventas " _
        & "GROUP BY P.Cta,CP.Cta_Venta_Anticipada,P.Pagos " _
        & "ORDER BY P.Cta,CP.Cta_Venta_Anticipada,P.Pagos "
  Select_Adodc AdoSubCta, sSQL
  With AdoSubCta.Recordset
   If .RecordCount > 0 Then
       Total = 0
       Cta = .Fields("Cta")
       Contra_Cta = .Fields("Cta_Venta_Anticipada")
       Do While Not .EOF
          If Cta <> .Fields("Cta") Then
             Total = Redondear(Total, 2)
             InsertarAsientos AdoAsiento, Contra_Cta, 0, Total, 0
             InsertarAsientos AdoAsiento, Cta, 0, 0, Total
             Total = 0
             Cta = .Fields("Cta")
             Contra_Cta = .Fields("Cta_Venta_Anticipada")
          End If
          Total = Total + .Fields("VentaAnt")
         .MoveNext
       Loop
       Total = Redondear(Total, 2)
       InsertarAsientos AdoAsiento, Contra_Cta, 0, Total, 0
       InsertarAsientos AdoAsiento, Cta, 0, 0, Total
   End If
  End With
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
  AsientoAuto.Caption = "INTERESES GANADOS"
  RatonNormal
End Sub

Private Sub Command3_Click()
  Unload Me
End Sub

Private Sub Command7_Click()
  FechaFinal = BuscarFecha(MBFechaF.Text)
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
          Co.TP = CompDiario
          Co.Numero = ReadSetDataNum("Diario", True, True)
          Co.CodigoB = Ninguno
          Co.Efectivo = 0
          Co.Fecha = MBFechaI.Text
          Co.Monto_Total = SumaDebe
          Co.Item = NumEmpresa
          Co.Usuario = CodigoUsuario
          Co.Concepto = LblConcepto.Caption
          Co.T_No = Trans_No
          GrabarComprobante Co
          SQL2 = "UPDATE Trans_Suscripciones " _
               & "SET AC = -1 " _
               & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND AC = " & Val(adFalse) & " "
          Ejecutar_SQL_SP SQL2
          ImprimirComprobantesDe False, Co
          RatonNormal
          Unload Me
       Else
         RatonNormal
         MsgBox "Las Transacciones no cuadran"
         DGAsiento.Visible = True
       End If
   Else
      RatonNormal
      MsgBox "No Existen Datos"
      DGAsiento.Visible = True
   End If
  End With
End Sub

Private Sub Form_Activate()
  Trans_No = 90
  FechaValida MBFechaI
  FechaValida MBFechaF
  IniciarAsientosDe DGAsiento, AdoAsiento
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm AsientoAuto
  ConectarAdodc AdoAux
  ConectarAdodc AdoSubCta
  ConectarAdodc AdoAsiento
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

