VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FPagoTJ 
   Caption         =   "PAGO DE TARJETAS DE CREIDTO/DEBITO"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17340
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   17340
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CheqRetCom 
      Caption         =   "Calcular Retencion menos Comision"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3045
      TabIndex        =   10
      Top             =   105
      Width           =   2220
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10395
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPagoTJ.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPagoTJ.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPagoTJ.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPagoTJ.frx":1BD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPagoTJ.frx":24B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   17340
      _ExtentX        =   30586
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Modulo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Calcular_Comision"
            Object.ToolTipText     =   "Calcular Comision"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Calcular_IVA"
            Object.ToolTipText     =   "Calcular Comision IVA"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Actualizar"
            Object.ToolTipText     =   "Calcular Retenciones"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Resultados"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGPagoTJ 
      Bindings        =   "FPagoTJ.frx":2904
      Height          =   3795
      Left            =   105
      TabIndex        =   9
      Top             =   1890
      Width           =   15555
      _ExtentX        =   27437
      _ExtentY        =   6694
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
   Begin MSMask.MaskEdBox MBFechaH 
      Height          =   330
      Left            =   1365
      TabIndex        =   4
      Top             =   1470
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
   Begin MSAdodcLib.Adodc AdoLote 
      Height          =   330
      Left            =   525
      Top             =   4620
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
      Caption         =   "Lote"
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
   Begin MSMask.MaskEdBox MBFechaD 
      Height          =   330
      Left            =   105
      TabIndex        =   3
      Top             =   1470
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
   Begin MSDataListLib.DataCombo DCTarjeta 
      Bindings        =   "FPagoTJ.frx":291C
      DataSource      =   "AdoTarjeta"
      Height          =   360
      Left            =   2625
      TabIndex        =   6
      Top             =   1470
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "SELECCIONE LA TARJETA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCLote 
      Bindings        =   "FPagoTJ.frx":2935
      DataSource      =   "AdoLote"
      Height          =   360
      Left            =   8925
      TabIndex        =   8
      Top             =   1470
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "000000"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoTarjeta 
      Height          =   330
      Left            =   525
      Top             =   2940
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
      Caption         =   "Tarjeta"
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
   Begin MSAdodcLib.Adodc AdoPagoTJ 
      Height          =   330
      Left            =   525
      Top             =   3360
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
      Caption         =   "PagoTJ"
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
      Top             =   3780
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
   Begin MSAdodcLib.Adodc AdoProvTJ 
      Height          =   330
      Left            =   525
      Top             =   4200
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
      Caption         =   "ProvTJ"
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
   Begin MSDataListLib.DataCombo DCProvTJ 
      Bindings        =   "FPagoTJ.frx":294B
      DataSource      =   "AdoProvTJ"
      Height          =   360
      Left            =   2625
      TabIndex        =   1
      Top             =   735
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "SELECCIONE LA TARJETA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TARJETA EMISORA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Width           =   2535
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SELECCIONE LA &TARJETA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   2625
      TabIndex        =   5
      Top             =   1155
      Width           =   6315
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &LOTE No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   8925
      TabIndex        =   7
      Top             =   1155
      Width           =   1380
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &FECHA DE LOTES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Top             =   1155
      Width           =   2535
   End
End
Attribute VB_Name = "FPagoTJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Actualizar()
    sSQL = "UPDATE Asiento_TJ " _
         & "SET IVA_BI = Total_TJ - Base_Imponible " _
         & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
         & "AND Lote = '" & DCLote & "' " _
         & "AND Cta = '" & Cta & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Asiento_TJ "
    If CheqRetCom.value <> 0 Then
       sSQL = sSQL & "SET IRF_Ret = ROUND((Base_Imponible - Comision) * (Porc_Ret/100),2,0)  "
    Else
       sSQL = sSQL & "SET IRF_Ret = ROUND(Base_Imponible * (Porc_Ret/100),2,0)  "
    End If
    sSQL = sSQL _
         & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
         & "AND Lote = '" & DCLote & "' " _
         & "AND Cta = '" & Cta & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Asiento_TJ " _
         & "SET IVA_Ret = ROUND(IVA_BI * (Porc_IVA/100),2,0) " _
         & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
         & "AND Lote = '" & DCLote & "' " _
         & "AND Cta = '" & Cta & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Asiento_TJ " _
         & "SET Neto_TJ = Total_TJ - IRF_Ret - IVA_Ret - Comision - Comision_IVA " _
         & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
         & "AND Lote = '" & DCLote & "' " _
         & "AND Cta = '" & Cta & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "SELECT * " _
         & "FROM Asiento_TJ " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Lote = '" & DCLote & "' " _
         & "AND Cta = '" & Cta & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "ORDER BY Fecha, Factura "
    Select_Adodc_Grid DGPagoTJ, AdoPagoTJ, sSQL, 2
End Sub

Public Sub Grabar()
   If SQL_Server Then
      sSQL = "UPDATE Trans_Abonos " _
           & "SET Neto_TJ = ATJ.Neto_TJ, " _
           & "IRF_Ret = ATJ.IRF_Ret, " _
           & "IVA_Ret = ATJ.IVA_Ret, " _
           & "Comision = ATJ.Comision, " _
           & "Comision_IVA = ATJ.Comision_IVA, " _
           & "Autorizacion_R = ATJ.Autorizacion_R, " _
           & "Serie_R = ATJ.Serie_R, " _
           & "Retencion = ATJ.Retencion, " _
           & "Porc_IVA = ATJ.Porc_IVA, " _
           & "Porc_Ret = ATJ.Porc_Ret, " _
           & "Codigo_Prov = '" & CodigoProv & "' " _
           & "FROM Trans_Abonos TA, Asiento_TJ As ATJ "
   Else
      sSQL = "UPDATE Trans_Abonos TA, Asiento_TJ As ATJ " _
           & "SET TA.Neto_TJ = ATJ.Neto_TJ, " _
           & "TA.IRF_Ret = ATJ.IRF_Ret, " _
           & "TA.IVA_Ret = ATJ.IVA_Ret, " _
           & "TA.Comision = ATJ.Comision, " _
           & "TA.Comision_IVA = ATJ.Comision_IVA, " _
           & "TA.Autorizacion_R = ATJ.Autorizacion_R, " _
           & "TA.Serie_R = ATJ.Serie_R, " _
           & "TA.Retencion = ATJ.Retencion, " _
           & "TA.Porc_IVA = ATJ.Porc_IVA, " _
           & "TA.Porc_Ret = ATJ.Porc_Ret, " _
           & "TA.Codigo_Prov = '" & CodigoProv & "' "
   End If
   sSQL = sSQL _
        & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND TA.Item = '" & NumEmpresa & "' " _
        & "AND TA.Periodo = '" & Periodo_Contable & "' " _
        & "AND TA.Tipo_Cta = 'TJ' " _
        & "AND TA.Cheque = ATJ.Lote " _
        & "AND TA.Cta = ATJ.Cta " _
        & "AND TA.Fecha = ATJ.Fecha " _
        & "AND TA.CodigoC = ATJ.CodigoC " _
        & "AND TA.Factura = ATJ.Factura " _
        & "AND TA.Item = ATJ.Item "
   Ejecutar_SQL_SP sSQL
   MsgBox "Proceso realizado con exito"
End Sub

Public Sub Calcular_Comision()
Dim Porc_Com As Single
     Porc_Com = InputBox("Ingrese el porcentaje de la comision:", "PORCENTAJES DE COMISIONES", "0")
     If Porc_Com > 0 Then
        'Base_Imponible
        sSQL = "UPDATE Asiento_TJ " _
             & "SET Comision = ROUND(Total_TJ * " & Val(Porc_Com / 100) & ",2,0) " _
             & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
             & "AND Lote = '" & DCLote & "' " _
             & "AND Cta = '" & Cta & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' "
        Ejecutar_SQL_SP sSQL
        
        sSQL = "SELECT * " _
             & "FROM Asiento_TJ " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Lote = '" & DCLote & "' " _
             & "AND Cta = '" & Cta & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' " _
             & "ORDER BY Fecha, Factura "
        Select_Adodc_Grid DGPagoTJ, AdoPagoTJ, sSQL, 2
     End If
End Sub

Public Sub Calcular_IVA()
Dim Porc_Com As Single
        'Base_Imponible
        sSQL = "UPDATE Asiento_TJ " _
             & "SET Comision_IVA = ROUND(Comision * " & Val(12 / 100) & ",2,0) " _
             & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
             & "AND Lote = '" & DCLote & "' " _
             & "AND Cta = '" & Cta & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' "
        Ejecutar_SQL_SP sSQL
        
        sSQL = "SELECT * " _
             & "FROM Asiento_TJ " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Lote = '" & DCLote & "' " _
             & "AND Cta = '" & Cta & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' " _
             & "ORDER BY Fecha, Factura "
        Select_Adodc_Grid DGPagoTJ, AdoPagoTJ, sSQL, 2
End Sub

Private Sub DCLote_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCLote_LostFocus()
    sSQL = "DELETE * " _
         & "FROM Asiento_TJ " _
         & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Lote = '" & DCLote & "' " _
         & "AND Cta = '" & Cta & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Ejecutar_SQL_SP sSQL

    sSQL = "INSERT INTO Asiento_TJ (Lote, Fecha, Factura, Base_Imponible, IVA_BI, Total_TJ, Comision,Comision_IVA, " _
         & "IVA_Ret, IRF_Ret, Neto_TJ, Autorizacion_R, Serie_R, Retencion, Cta, CodigoC, Item, CodigoU, Porc_IVA, Porc_Ret) " _
         & "SELECT Cheque, Fecha, Factura, ROUND(Abono/1.12,2,1), 0 As IVAs, Abono, Comision, Comision_IVA, IVA_Ret, IRF_Ret, " _
         & "Neto_TJ, Autorizacion_R, Serie_R, Retencion, Cta, CodigoC, Item, '" & CodigoUsuario & "' As CodigoUx,Porc_IVA,Porc_Ret " _
         & "FROM Trans_Abonos " _
         & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
         & "AND Cheque = '" & DCLote & "' " _
         & "AND Cta = '" & Cta & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Asiento_TJ " _
         & "SET IVA_BI = Total_TJ - Base_Imponible " _
         & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
         & "AND Lote = '" & DCLote & "' " _
         & "AND Cta = '" & Cta & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Ejecutar_SQL_SP sSQL

    sSQL = "SELECT * " _
         & "FROM Asiento_TJ " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Lote = '" & DCLote & "' " _
         & "AND Cta = '" & Cta & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "ORDER BY Fecha, Factura "
    Select_Adodc_Grid DGPagoTJ, AdoPagoTJ, sSQL, 2
End Sub

Private Sub DCProvTJ_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCProvTJ_LostFocus()
  With AdoProvTJ.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCProvTJ.Text & "' ")
       If Not .EOF Then
          CodigoProv = .Fields("Codigo")
          TBeneficiario = Leer_Datos_Clientes(CodigoProv)
          NombreCliente = TBeneficiario.Cliente
       End If
   End If
  End With
End Sub

Private Sub DCTarjeta_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCTarjeta_LostFocus()
   FechaValida MBFechaD
   FechaValida MBFechaH
   FechaIni = BuscarFecha(MBFechaD)
   FechaFin = BuscarFecha(MBFechaH)
   Cta = SinEspaciosIzq(DCTarjeta)
   sSQL = "SELECT Cheque " _
        & "FROM Trans_Abonos " _
        & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Tipo_Cta = 'TJ' " _
        & "AND Cta = '" & Cta & "' " _
        & "GROUP BY Cheque " _
        & "ORDER BY Cheque "
   SelectDB_Combo DCLote, AdoLote, sSQL, "Cheque"
   
End Sub

Private Sub DGPagoTJ_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyReturn Then
     If AdoPagoTJ.Recordset.RecordCount > 0 Then
        AdoPagoTJ.Recordset.MoveNext
        If AdoPagoTJ.Recordset.EOF Then AdoPagoTJ.Recordset.MoveFirst
        MsgBox "Proceso Ok"
     End If
  End If
End Sub

Private Sub Form_Activate()
  DGPagoTJ.Height = MDI_Y_Max - DGPagoTJ.Top - 400
  DGPagoTJ.width = MDI_X_Max - 100
  'AdoQuery.Top = DGQuery.Top + DGQuery.Height
  sSQL = "SELECT Cliente,Codigo,CI_RUC " _
       & "FROM Clientes  " _
       & "WHERE Codigo <> '.' " _
       & "ORDER BY Cliente "
  SelectDB_Combo DCProvTJ, AdoProvTJ, sSQL, "Cliente"

  sSQL = "SELECT Codigo & Space(2) & Cuenta As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'TJ' " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCTarjeta, AdoTarjeta, sSQL, "NomCuenta"
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoLote
  ConectarAdodc AdoPagoTJ
  ConectarAdodc AdoProvTJ
  ConectarAdodc AdoTarjeta
End Sub

Private Sub MBFechaD_GotFocus()
  MarcarTexto MBFechaD
End Sub

Private Sub MBFechaD_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaD_LostFocus()
  FechaValida MBFechaD
End Sub

Private Sub MBFechaH_GotFocus()
  MarcarTexto MBFechaH
End Sub

Private Sub MBFechaH_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaH_LostFocus()
  FechaValida MBFechaH
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    Case "Salir"
         Unload FPagoTJ
    Case "Calcular_Comision"
         Calcular_Comision
    Case "Calcular_IVA"
         Calcular_IVA
    Case "Actualizar"
         Actualizar
    Case "Grabar"
         Grabar
  End Select
  RatonNormal
End Sub


