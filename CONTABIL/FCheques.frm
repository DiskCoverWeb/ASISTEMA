VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form PCheques 
   Caption         =   "LISTAR E IMPRIMIR CHEQUES EMITIDOS"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   11565
   WindowState     =   2  'Maximized
   Begin MSDataListLib.DataCombo DCComp 
      Bindings        =   "FCheques.frx":0000
      DataSource      =   "AdoComp"
      Height          =   345
      Left            =   3570
      TabIndex        =   2
      Top             =   105
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "Comp"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataList DLCta 
      Bindings        =   "FCheques.frx":0016
      DataSource      =   "AdoCta"
      Height          =   735
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1296
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtHasta 
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
      Left            =   3570
      MaxLength       =   8
      TabIndex        =   6
      Top             =   735
      Width           =   1695
   End
   Begin VB.TextBox TxtDesde 
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
      Left            =   1890
      MaxLength       =   8
      TabIndex        =   4
      Top             =   735
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir &Bloque de Cheque"
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
      Left            =   6930
      Picture         =   "FCheques.frx":002B
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   105
      Width           =   1485
   End
   Begin VB.PictureBox PictCheques 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   105
      ScaleHeight     =   6.033
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   19.923
      TabIndex        =   12
      Top             =   2205
      Width           =   11355
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Imprimir Cheque"
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
      Left            =   5355
      Picture         =   "FCheques.frx":08AD
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   105
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
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
      Left            =   8505
      Picture         =   "FCheques.frx":112F
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   105
      Width           =   1485
   End
   Begin MSAdodcLib.Adodc AdoComp1 
      Height          =   330
      Left            =   420
      Top             =   3360
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
      Caption         =   "Comp1"
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
   Begin MSAdodcLib.Adodc AdoComp 
      Height          =   330
      Left            =   420
      Top             =   3045
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
      Caption         =   "Comp"
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
   Begin MSAdodcLib.Adodc AdoCta 
      Height          =   330
      Left            =   420
      Top             =   3675
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
      Caption         =   "Cta"
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
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   420
      Top             =   3990
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
      Caption         =   "Cliente"
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cheque No."
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
      Left            =   1890
      TabIndex        =   13
      Top             =   105
      Width           =   1695
   End
   Begin VB.Label LabelEst 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
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
      Height          =   960
      Left            =   6930
      TabIndex        =   11
      Top             =   1155
      Width           =   4635
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cheque &Hasta"
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
      Left            =   3570
      TabIndex        =   5
      Top             =   420
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cheque &Desde"
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
      Left            =   1890
      TabIndex        =   3
      Top             =   420
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
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
      Height          =   960
      Left            =   105
      TabIndex        =   10
      Top             =   1155
      Width           =   6735
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Cuentas"
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
      Width           =   1800
   End
End
Attribute VB_Name = "PCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ListaDeCheques(CtaCheq As String)
  sSQL = "SELECT C.T,C.TP,C.Numero,Cl.Cliente,C.Codigo_B,C.Fecha,C1.Cuenta,Cl.CI_RUC," _
       & "Cl.Direccion,T.Cta,T.Cheq_Dep,T.Haber,T.Codigo_C,T.Fecha_Efec " _
       & "FROM Transacciones As T,Comprobantes As C,Clientes As Cl,Catalogo_Cuentas As C1 " _
       & "WHERE IsNumeric(T.Cheq_Dep) <> " & adFalse & " " _
       & "AND T.Item = '" & NumEmpresa & "' " _
       & "AND T.Periodo = '" & Periodo_Contable & "' " _
       & "AND T.Cta = '" & CtaCheq & "' " _
       & "AND T.Haber > 0 " _
       & "AND T.Item = C.Item " _
       & "AND T.Item = C1.Item " _
       & "AND T.Periodo = C.Periodo " _
       & "AND T.Periodo = C1.Periodo " _
       & "AND T.TP = C.TP " _
       & "AND T.Numero = C.Numero " _
       & "AND C.Codigo_B = Cl.Codigo " _
       & "AND C1.Codigo = T.Cta " _
       & "ORDER BY T.Cheq_Dep DESC "
 'MsgBox sSQL
  SelectDB_Combo DCComp, AdoComp, sSQL, "Cheq_Dep"
  RatonNormal
End Sub

Private Sub Command1_Click()
  Unload PCheques
End Sub

Private Sub Command2_Click()
  Imprimir_Bloque_Cheques NumCheque, NumCheque, DLCta.Text, DCComp.Text, Label1.Caption
End Sub

Private Sub Command3_Click()
  Imprimir_Bloque_Cheques TxtDesde, TxtHasta, DLCta, DCComp, Label1.Caption
End Sub

Private Sub DCComp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCComp_LostFocus()
  ListarCheques
End Sub

Private Sub DLCta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLCta_LostFocus()
  Cta_Aux = DLCta.Text
  Codigo1 = Cta_Aux
  ListaDeCheques Cta_Aux
End Sub

Private Sub Form_Activate()
Dim X As Printer

  PictCheques.width = MDI_X_Max - 200
  PictCheques.Height = MDI_Y_Max - PictCheques.Top - 100
  PictCheques.Cls
  sSQL = "SELECT Codigo,Cliente,CI_RUC " _
       & "FROM Clientes " _
       & "WHERE Codigo <> '.' " _
       & "ORDER BY Cliente "
  Select_Adodc AdoCliente, sSQL
  
  CEConLineas = ProcesarSeteos("Egresos")
  sSQL = "SELECT Codigo " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'BA' " _
       & "ORDER BY Codigo "
  SelectDB_List DLCta, AdoCta, sSQL, "Codigo"
  Cta_Aux = DLCta.Text
  Codigo1 = Cta_Aux
  ListaDeCheques Cta_Aux
  If AdoComp.Recordset.RecordCount > 0 Then ListarCheques
  RatonNormal
End Sub

Private Sub Form_Load()
 'CentrarForm PCheques
  ConectarAdodc AdoCta
  ConectarAdodc AdoComp
  ConectarAdodc AdoComp1
  ConectarAdodc AdoCliente
End Sub

Public Sub ListarCheques()
Dim TipoBank As String
  RatonReloj
  PictCheques.FontSize = 10
  PictCheques.FontName = TipoCourier
  Codigo1 = SinEspaciosIzq(DLCta.Text)    ' Cta
  Codigo2 = DCComp.Text   ' Cheque No
  NumCheque = Val(Codigo2)
  TipoBank = TrimStrg(MidStrg(Codigo1, Len(Codigo1) - 1, 2))
  CCHQConLineas = ProcesarSeteos(TipoBank)
  With AdoComp.Recordset
   If .RecordCount > 0 Then
     ' Presentacion de Comp
      'MsgBox Codigo1 & vbCrLf & Codigo2 & vbCrLf & TipoBank
      .MoveFirst
      .Find ("Cta = '" & Codigo1 & "' ")
       If Not .EOF Then
         .Find ("Cheq_Dep = '" & Codigo2 & "' ")
          If Not .EOF Then
             
            'MsgBox TipoBank
             CCHQConLineas = ProcesarSeteos(TipoBank)
             Cta = .Fields("Cta")
             If .Fields("T") = Anulado Then LabelEst.Caption = "ANULADO" Else LabelEst.Caption = "NORMAL"
             FechaTexto = .Fields("Fecha")
             Beneficiario = .Fields("Cliente")
             CodigoCli = .Fields("Codigo_C")
             If Beneficiario = Ninguno Then
                If AdoCliente.Recordset.RecordCount > 0 Then
                   AdoCliente.Recordset.MoveFirst
                   AdoCliente.Recordset.Find ("Codigo = '" & CodigoCli & "' ")
                   If Not AdoCliente.Recordset.EOF Then Beneficiario = AdoCliente.Recordset.Fields("Cliente")
                End If
             End If
             Valor = .Fields("Haber")
             NumCheque = .Fields("Cheq_Dep")
             Cuenta = .Fields("Cta")
             LabelEst.Caption = LabelEst.Caption & vbCrLf _
                              & "Cuenta No. " & Cuenta & vbCrLf _
                              & "Cheque No. " & Format(NumCheque, "00000000")
             Label1.Caption = "Comprobante No.  " & .Fields("TP") & " - " & .Fields("Numero") _
                            & Space(15) & "POR " & Moneda & " " & Format(Valor, "#,##0.00") & vbCrLf _
                            & "Pagado a: " & Beneficiario & vbCrLf _
                            & .Fields("Cuenta") & vbCrLf
             If CFechaLong(.Fields("Fecha")) < CFechaLong(.Fields("Fecha_Efec")) Then
                Label1.Caption = Label1.Caption & "CHEQUE POSFECHADO"
                FechaTexto = .Fields("Fecha_Efec")
             End If
             PictCheques.AutoRedraw = True
             PictCheques.ScaleMode = vbCentimeters
             PictCheques.Cls
             PictCheques.FontBold = True
             PictCheques.DrawWidth = 1
             PictCheques.FontName = TipoTimes
             
             PictCheques.FontSize = SetD(2).Porte
             PictCheques.CurrentX = SetD(2).PosX
             PictCheques.CurrentY = SetD(2).PosY
             PictCheques.Print Beneficiario
             
             PictCheques.FontSize = SetD(3).Porte
             PictCheques.CurrentX = SetD(3).PosX
             PictCheques.CurrentY = SetD(3).PosY
             PictCheques.Print Format(Valor, "#,###.00")
             
             If SetD(4).PosX > 0 And SetD(4).PosY > 0 Then
                PictCheques.FontSize = SetD(4).Porte
                PictCheques.CurrentX = SetD(4).PosX
                PictCheques.CurrentY = SetD(4).PosY
                PictCheques.Print Cambio_Letras(Valor, 2)
             End If
             If SetD(9).PosX > 0 And SetD(9).PosY > 0 Then
                Cadena = Empresa & " " & Moneda & " " & Format(Valor, "#,##0.00")
                PictCheques.FontSize = SetD(9).Porte
                PictCheques.CurrentX = SetD(9).PosX
                PictCheques.CurrentY = SetD(9).PosY
                PictCheques.Print Cadena
             End If
             PictCheques.FontSize = SetD(10).Porte
             PictCheques.CurrentX = SetD(10).PosX
             PictCheques.CurrentY = SetD(10).PosY
             PictCheques.Print ULCase(NombreCiudad)
             
             PictCheques.FontSize = SetD(6).Porte
             PictCheques.CurrentX = SetD(6).PosX
             PictCheques.CurrentY = SetD(6).PosY
             PictCheques.Print Format(FechaTexto, "yyyy/MM/dd")
             
             RatonNormal
         End If
       End If
   Else
       RatonNormal
       MsgBox "Este Comprobante no Existe."
   End If
  End With
End Sub

Private Sub TxtDesde_GotFocus()
  MarcarTexto TxtDesde
End Sub

Private Sub TxtDesde_LostFocus()
  TextoValido TxtDesde
End Sub

Private Sub TxtHasta_GotFocus()
  MarcarTexto TxtHasta
End Sub

Private Sub TxtHasta_LostFocus()
  TextoValido TxtHasta
End Sub
