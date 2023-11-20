VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FAnularMasa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ANULACION VARIOS COMPROBANTES"
   ClientHeight    =   3720
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FAnularMasa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo DCTP 
      Bindings        =   "FAnularMasa.frx":0442
      DataSource      =   "AdoTP"
      Height          =   315
      Left            =   2205
      TabIndex        =   3
      Top             =   840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "CD"
   End
   Begin VB.CommandButton Command1 
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
      Height          =   750
      Left            =   4515
      Picture         =   "FAnularMasa.frx":0456
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   945
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
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
      Height          =   750
      Left            =   4515
      Picture         =   "FAnularMasa.frx":0E4C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   105
      Width           =   1065
   End
   Begin MSDataListLib.DataList DLDesde 
      Bindings        =   "FAnularMasa.frx":1716
      DataSource      =   "AdoListComp"
      Height          =   2010
      Left            =   105
      TabIndex        =   5
      Top             =   1575
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   3545
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoSup 
      Height          =   330
      Left            =   4515
      Top             =   1890
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Sup"
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
   Begin MSAdodcLib.Adodc AdoListComp 
      Height          =   330
      Left            =   4515
      Top             =   2205
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "ListComp"
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
   Begin MSAdodcLib.Adodc AdoTP 
      Height          =   330
      Left            =   4515
      Top             =   2520
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "TP"
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
   Begin MSDataListLib.DataList DLHasta 
      Bindings        =   "FAnularMasa.frx":1730
      DataSource      =   "AdoListComp"
      Height          =   2010
      Left            =   2310
      TabIndex        =   7
      Top             =   1575
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   3545
      _Version        =   393216
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Hasta:"
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
      TabIndex        =   6
      Top             =   1260
      Width           =   2115
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Desde:"
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
      Top             =   1260
      Width           =   2115
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tipo de Comprobante:"
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
      Top             =   840
      Width           =   2115
   End
   Begin VB.Label LblUsuario 
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
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   4320
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Comprobantes del Usuario:"
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
      Width           =   4320
   End
End
Attribute VB_Name = "FAnularMasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Anular_Comprobantes_Tabla(TP As String, Tabla As String)
    sSQL = "UPDATE " & Tabla & " " _
         & "SET T = 'A' "
    Select Case Tabla
      Case "Comprobantes": sSQL = sSQL & ", Concepto = '(ANULAR) ' + Concepto "
      Case "Trans_SubCtas": sSQL = sSQL & ", Debitos = 0, Creditos = 0 "
      Case "Transacciones": sSQL = sSQL & ", Debe = 0, Haber = 0 "
    End Select
    sSQL = sSQL _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TP = '" & TP & "' " _
         & "AND Numero BETWEEN " & Val(DLDesde) & " AND " & Val(DLHasta) & " "
    Ejecutar_SQL_SP sSQL
End Sub

Private Sub Command1_Click()
  Unload FAnularMasa
End Sub

Private Sub Command2_Click()

 If Val(DLDesde) <= Val(DLHasta) And Len(DCTP) = 2 Then
    Titulo = "ANULACION VARIOS COMPROBANTES"
    Mensajes = "Seguro que quiere anular los " & DCTP & vbCrLf _
             & "Desde " & Val(DLDesde) & " hasta " & Val(DLHasta) & "? "
    If BoxMensaje = vbYes Then
        Anular_Comprobantes_Tabla DCTP, "Comprobantes"
        Anular_Comprobantes_Tabla DCTP, "Trans_Air"
        Anular_Comprobantes_Tabla DCTP, "Trans_Compras"
        Anular_Comprobantes_Tabla DCTP, "Trans_Exportaciones"
        Anular_Comprobantes_Tabla DCTP, "Trans_Importaciones"
        Anular_Comprobantes_Tabla DCTP, "Trans_Kardex"
        Anular_Comprobantes_Tabla DCTP, "Trans_Rol_Pagos"
        Anular_Comprobantes_Tabla DCTP, "Trans_SubCtas"
        Anular_Comprobantes_Tabla DCTP, "Trans_Ventas"
        Anular_Comprobantes_Tabla DCTP, "Transacciones"
    
       'Actualizar las Ctas a mayoriazar
        sSQL = "SELECT Cta " _
             & "FROM Transacciones " _
             & "WHERE Periodo = '" & Periodo_Contable & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND TP = '" & DCTP & "' " _
             & "AND Numero BETWEEN " & Val(DLDesde) & " AND " & Val(DLHasta) & " " _
             & "GROUP BY Cta " _
             & "ORDER BY Cta "
        Select_Adodc AdoSup, sSQL
        With AdoSup.Recordset
         If .RecordCount > 0 Then
             Do While Not .EOF
               'Determinamos que la cuenta ya fue mayorizada
                SubCta = .Fields("Cta")
                sSQL = "UPDATE Transacciones " _
                     & "SET Procesado = " & Val(adFalse) & " " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' " _
                     & "AND Cta = '" & SubCta & "' "
                Ejecutar_SQL_SP sSQL
               .MoveNext
             Loop
         End If
        End With
        sSQL = "SELECT Codigo_Inv " _
             & "FROM Trans_Kardex " _
             & "WHERE Periodo = '" & Periodo_Contable & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND TP = '" & DCTP & "' " _
             & "AND Numero BETWEEN " & Val(DLDesde) & " AND " & Val(DLHasta) & " " _
             & "GROUP BY Codigo_Inv " _
             & "ORDER BY Codigo_Inv "
        Select_Adodc AdoSup, sSQL
        With AdoSup.Recordset
         If .RecordCount > 0 Then
             Do While Not .EOF
               'Determinamos que la cuenta ya fue mayorizada
                SubCta = .Fields("Codigo_Inv")
                sSQL = "UPDATE Trans_Kardex " _
                     & "SET Procesado = " & Val(adFalse) & " " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' " _
                     & "AND Codigo_Inv = '" & SubCta & "' "
                Ejecutar_SQL_SP sSQL
               .MoveNext
             Loop
         End If
        End With
        Control_Procesos "A", "Anulo los Comprobante de " & DCTP & " desede " & Val(DLDesde) & " hasta " & Val(DLHasta)
        MsgBox "Proceso Terminado, Vuelva a Mayorizar"
        Unload FAnularMasa
    End If
 Else
    MsgBox "Rango Invalidos, no se procedera hacer nada"
 End If
 Unload IngFechas
End Sub

Private Sub DCTP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTP_LostFocus()
  Listar_Comprobantes_Usuario DCTP
End Sub

Private Sub Form_Activate()
    LblUsuario.Caption = " " & NombreUsuario
    sSQL = "SELECT TP " _
         & "FROM Comprobantes " _
         & "WHERE Periodo = '" & Periodo_Contable & "' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "GROUP BY TP " _
         & "ORDER BY TP "
    SelectDB_Combo DCTP, AdoTP, sSQL, "TP"
    Listar_Comprobantes_Usuario DCTP
End Sub

Private Sub Form_Load()
  CentrarForm FAnularMasa
  ConectarAdodc AdoTP
  ConectarAdodc AdoSup
  ConectarAdodc AdoListComp
  RatonNormal
End Sub

Public Sub Listar_Comprobantes_Usuario(TP As String)
    sSQL = "SELECT Numero " _
         & "FROM Comprobantes " _
         & "WHERE Periodo = '" & Periodo_Contable & "' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND TP = '" & TP & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "GROUP BY Numero " _
         & "ORDER BY Numero "
    SelectDB_List DLDesde, AdoListComp, sSQL, "Numero"
    SelectDB_List DLHasta, AdoListComp, sSQL, "Numero"
    DLDesde.SetFocus
End Sub
