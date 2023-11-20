VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FCalculos 
   Caption         =   "Cálculos"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FCalculos.frx":0000
      Height          =   1020
      Left            =   210
      TabIndex        =   9
      Top             =   2625
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   1799
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Datos Ingresados"
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
   Begin VB.ComboBox CModulos 
      Height          =   315
      ItemData        =   "FCalculos.frx":001D
      Left            =   210
      List            =   "FCalculos.frx":002D
      TabIndex        =   8
      Top             =   210
      Width           =   2220
   End
   Begin VB.ComboBox CAño 
      Height          =   315
      ItemData        =   "FCalculos.frx":0060
      Left            =   2730
      List            =   "FCalculos.frx":0062
      TabIndex        =   5
      Top             =   735
      Width           =   1380
   End
   Begin MSDataListLib.DataCombo DCMes 
      Bindings        =   "FCalculos.frx":0064
      DataSource      =   "AdoMes"
      Height          =   315
      Left            =   210
      TabIndex        =   3
      Top             =   735
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   750
      Left            =   9030
      TabIndex        =   1
      Top             =   210
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   750
      Left            =   7455
      TabIndex        =   0
      Top             =   210
      Width           =   1380
   End
   Begin MSDataGridLib.DataGrid DGCal 
      Bindings        =   "FCalculos.frx":0079
      Height          =   1020
      Left            =   210
      TabIndex        =   2
      Top             =   1260
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   1799
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Datos Ingresados"
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
   Begin MSAdodcLib.Adodc AdoMes 
      Height          =   330
      Left            =   210
      Top             =   2625
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "Meses"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoTransCompras 
      Height          =   330
      Left            =   210
      Top             =   1995
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "TransCompras"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoTransVentas 
      Height          =   330
      Left            =   210
      Top             =   1680
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "TransVentas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoTransAir 
      Height          =   330
      Left            =   210
      Top             =   2310
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "TransAir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Seleccione el Módulo:"
      Height          =   225
      Left            =   210
      TabIndex        =   7
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Seleccione el Año:"
      Height          =   225
      Left            =   2730
      TabIndex        =   6
      Top             =   525
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccione el Mes:"
      Height          =   225
      Left            =   210
      TabIndex        =   4
      Top             =   525
      Width           =   1695
   End
End
Attribute VB_Name = "FCalculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cod As Byte

Private Sub Command1_Click()
Dim aux As String
Contador = 1
Select Case CModulos
    Case "Compras"
    FechaIni = "01/" & Format(cod, "00") & "/" & Format(CAño, "0000")
    FechaFin = UltimoDiaMes(FechaIni)
         sSQL = "SELECT * " _
              & "FROM Trans_Compras " _
              & "WHERE FechaEmision Between # " & BuscarFecha(FechaIni) & " # AND # " & BuscarFecha(FechaFin) & " # " _
              & "AND IdProv <> '.' " _
              & "ORDER BY IdProv,Fecha,Establecimiento,PuntoEmision,Secuencial "
         SelectAdodc AdoTransCompras, sSQL
         With AdoTransCompras.Recordset
           If .RecordCount > 0 Then
               Do While Not .EOF
                 .Fields("Linea_SRI") = Contador
                  Contador = Contador + 1
                 .MoveNext
               Loop
              .UpdateBatch
           End If
         End With
  
     Case "Ventas"
           sSQL = "SELECT * " _
                 & "FROM Trans_Ventas " _
                 & "WHERE FechaEmision Between # " & BuscarFecha(FechaIni) & " # AND # " & BuscarFecha(FechaFin) & " # " _
                 & "AND IdProv <> '.' " _
                 & "ORDER BY IdProv,Fecha,Establecimiento,PuntoEmision,Secuencial "
            SelectAdodc AdoTransVentas, sSQL
            With AdoTransVentas.Recordset
             If .RecordCount > 0 Then
                 Do While Not .EOF
                   .Fields("Linea_SRI") = Contador
                    Contador = Contador + 1
                   .MoveNext
                   If Contador > 10 Then
                      .MoveLast
                    End If
                 Loop
                 .UpdateBatch
             End If
            End With
               
     
End Select
  
  
  
  'Ventas
  
  
'  sSQL = "SELECT * " _
'        & "FROM Tabla_Por_ICE_IVA " _
'        & "WHERE #" & BuscarFecha(MBFechaRegis) & "# Between Fecha_Inicio AND Fecha_Final " _
'        & "AND IVA <> " & Val(adFalse) & " " _
'        & "ORDER BY Porc "
'   SelectDBCombo DCPorcenIva, AdoPorIva, sSQL, "Porc"
'
  
'  Contador = 1
'  sSQL = "SELECT * " _
'       & "FROM Trans_Ventas " _
'       & "WHERE IdProv <> '.' " _
'       & "ORDER BY IdProv,Fecha,Establecimiento,PuntoEmision,Secuencial "
'  SelectAdodc AdoTransCompras, sSQL
'  With AdoTransCompras.Recordset
'   If .RecordCount > 0 Then
'       Do While Not .EOF
'         .Fields("Linea_SRI") = Contador
'          Contador = Contador + 1
'         .MoveNext
'       Loop
'       .UpdateBatch
'   End If
'
'  End With

End Sub

Private Sub DCMes_LostFocus()
    With AdoMes.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
           .Find ("Mes = '" & DCMes & "' ")
           If Not .EOF Then
              cod = .Fields("NoMes")
           Else
              MsgBox "Vuelva a seleccionar", vbInformation, "Aviso"
           End If
        End If
    End With
End Sub

Private Sub Form_Activate()
  sSQL = Listar_Meses
  SelectDBCombo DCMes, AdoMes, sSQL, "Dia_Mes"
  For I = 2000 To 2007
     CAño.AddItem I
  Next I
End Sub

Private Sub Form_Load()
    ConectarAdodc AdoTransCompras
    ConectarAdodc AdoMes
End Sub
