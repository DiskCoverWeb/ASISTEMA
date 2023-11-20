VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FEjercicio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compras"
   ClientHeight    =   6012
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6012
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   3456
      Top             =   3564
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Borrar"
      Height          =   645
      Left            =   5040
      Picture         =   "FEjercicio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3885
      Width           =   1485
   End
   Begin VB.CommandButton CmdInsertar 
      Caption         =   "Insertar"
      Height          =   750
      Left            =   7035
      Picture         =   "FEjercicio.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3780
      Width           =   1800
   End
   Begin VB.TextBox TxtvalRetAir 
      Height          =   330
      Left            =   7455
      TabIndex        =   8
      Top             =   2625
      Width           =   1800
   End
   Begin VB.TextBox TxtporcentajeAir 
      Height          =   330
      Left            =   7455
      TabIndex        =   7
      Top             =   2100
      Width           =   1800
   End
   Begin VB.TextBox TxtbaseImpAir 
      Height          =   330
      Left            =   7455
      TabIndex        =   6
      Top             =   1575
      Width           =   1800
   End
   Begin VB.TextBox TxtcodRetAir 
      Height          =   330
      Left            =   7455
      TabIndex        =   5
      Top             =   1050
      Width           =   1800
   End
   Begin MSDataGridLib.DataGrid DGAir 
      Bindings        =   "FEjercicio.frx":0614
      Height          =   1380
      Left            =   210
      TabIndex        =   0
      Top             =   1155
      Width           =   4845
      _ExtentX        =   8551
      _ExtentY        =   2434
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
   Begin MSAdodcLib.Adodc AdoAir 
      Height          =   330
      Left            =   315
      Top             =   315
      Width           =   1905
      _ExtentX        =   3366
      _ExtentY        =   572
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
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSForms.Label lblTime 
      Height          =   660
      Left            =   540
      TabIndex        =   11
      Top             =   3456
      Width           =   2496
      Size            =   "4403;1164"
      FontHeight      =   156
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   330
      Left            =   5565
      TabIndex        =   4
      Top             =   2100
      Width           =   1590
      Caption         =   "Porcentaje Air"
      Size            =   "2805;582"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   435
      Left            =   5565
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
      Caption         =   "Valor Retención Air"
      Size            =   "2990;767"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   330
      Left            =   5565
      TabIndex        =   2
      Top             =   1680
      Width           =   1485
      Caption         =   "Base Imponible Air"
      Size            =   "2619;582"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   5565
      TabIndex        =   1
      Top             =   1155
      Width           =   1590
      Caption         =   "Codigo Retencion Air"
      Size            =   "2805;582"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FEjercicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdBorrar_Click()
    RatonReloj
     sSQL = "DELETE * " _
         & "FROM Trans_Air " _
         & "WHERE codRetAir = '" & TxtcodRetAir & "' "
         
    ConectarAdoExecute sSQL
    sSQL = "SELECT * " _
        & "FROM Trans_Air " _
        & "WHERE codRetAir <> '.' " _
        & "ORDER BY codRetAir "
   SelectDataGrid DGAir, AdoAir, sSQL
   RatonNormal

End Sub

Private Sub CmdInsertar_Click()
    RatonReloj
     sSQL = "INSERT INTO Trans_Air " _
         & "(codRetAir,baseImpAir,porcentajeAir,valRetAir,Periodo,Item,CodigoU) " _
         & "VALUES " _
         & "('" & TxtcodRetAir & "'," & TxtbaseImpAir & "," & TxtporcentajeAir & "," & TxtvalRetAir & ",'" & Periodo_Contable & "','" & NumEmpresa & "','" & CodigoUsuario & "') "
    ConectarAdoExecute sSQL
    sSQL = "SELECT * " _
        & "FROM Trans_Air " _
        & "WHERE codRetAir <> '.' " _
        & "ORDER BY codRetAir "
   SelectDataGrid DGAir, AdoAir, sSQL
   RatonNormal
   
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
   sSQL = "SELECT * " _
        & "FROM Trans_Air " _
        & "WHERE codRetAir <> '.' " _
        & "ORDER BY codRetAir "
   SelectDataGrid DGAir, AdoAir, sSQL
   RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAir
  
  'Union
  sSQL = "SELECT Cliente,SD.Fecha,Numero,Ingresos AS Subtotal,ENE As Total_IVA,FEB As Total_ICE" _
         & "ABR As R_1,MAY As R_2,JUN As R_5,JUL As R_8,AGO As R_15, SEP As R_25,Saldo_Actual As Total_RET " _
         & "FROM Saldo_Diarios As SD, Clientes As C " _
         & "WHERE SD.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  " _
         & "AND SD.Item = '" & NumEmpresa & "' " _
         & "AND SD.CodigoU = '" & CodigoUsuario & "' " _
         & "AND SD.TP = 'ATIM' " _
         & "AND SD.CodigoC = C.Codigo " _
         & "ORDER BY SD.PEN,SD.Fecha " _
         & "UNION " _
         & "SELECT 'z - TOTAL IMPORTACIONES' As Cliente,'" & FechaSistema & "' As Fecha,0 As Numero,SUM(Ingresos) AS Subtotal,SUM(ENE) As Total_IVA,SUM(FEB) As Total_ICE " _
         & "SUM(ABR) As R_1,SUM(MAY) As R_2,SUM(JUN) As R_5,SUM(JUL) As R_8,SUM(AGO) As R_15, SUM(SEP) As R_25,SUM(Saldo_Actual) As Total_RET " _
         & "FROM Saldo_Diarios As SD " _
         & "WHERE SD.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  " _
         & "AND SD.Item = '" & NumEmpresa & "' " _
         & "AND SD.CodigoU = '" & CodigoUsuario & "' " _
         & "AND SD.TP = 'ATIM' " _
         & "GROUP BY SD.TP "
    SelectDataGrid DGAir, AdoAir, sSQL
End Sub

Private Sub Timer1_Timer()
   If lblTime.Caption <> CStr(Time) Then
      lblTime.Caption = Time
   End If
End Sub
