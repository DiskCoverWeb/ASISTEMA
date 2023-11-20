VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FModTransSC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSACCIONES CON SUB CUENTAS DE BLOQUE"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   11145
   Begin VB.CommandButton Command2 
      Caption         =   "Consulta Comprobante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   4830
      Picture         =   "Fmtransc.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   105
      Width           =   1380
   End
   Begin VB.TextBox TextNumero 
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
      Left            =   1050
      TabIndex        =   3
      Top             =   1050
      Width           =   1065
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5475
      Left            =   105
      TabIndex        =   6
      Top             =   1470
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   9657
      _Version        =   327680
      TabHeight       =   520
      TabCaption(0)   =   "Comprobante y Transacciones"
      TabPicture(0)   =   "Fmtransc.frx":0442
      Tab(0).ControlCount=   10
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LDebe_MN"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LHaber_MN"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LDebe_ME"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LHaber_ME"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DBGComp"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DataComp"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "DBGTrans"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "DataTrans"
      Tab(0).Control(9).Enabled=   0   'False
      TabCaption(1)   =   "Submódulo de Cuentas"
      TabPicture(1)   =   "Fmtransc.frx":045E
      Tab(1).ControlCount=   8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataTransSC"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DBGTransSC"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "LSHaber_ME"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "LSDebe_ME"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label13"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "LSHaber_MN"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "LSDebe_MN"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label10"
      Tab(1).Control(7).Enabled=   0   'False
      TabCaption(2)   =   "Bancos y Retenciones"
      TabPicture(2)   =   "Fmtransc.frx":047A
      Tab(2).ControlCount=   6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DataRet"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "DBGRet"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "DataBanco"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "DBGBanco"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "LRet_MN"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label16"
      Tab(2).Control(5).Enabled=   0   'False
      Begin VB.Data DataTransSC 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -74895
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4620
         Width           =   10725
      End
      Begin VB.Data DataRet 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -74895
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4620
         Width           =   10725
      End
      Begin VB.Data DataTrans 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   105
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4620
         Width           =   10725
      End
      Begin MSDBGrid.DBGrid DBGRet 
         Bindings        =   "Fmtransc.frx":0496
         Height          =   1905
         Left            =   -74895
         OleObjectBlob   =   "Fmtransc.frx":04A8
         TabIndex        =   10
         Top             =   2730
         Width           =   10725
      End
      Begin MSDBGrid.DBGrid DBGTrans 
         Bindings        =   "Fmtransc.frx":0E61
         Height          =   2640
         Left            =   105
         OleObjectBlob   =   "Fmtransc.frx":0E75
         TabIndex        =   7
         Top             =   1995
         Width           =   10725
      End
      Begin VB.Data DataBanco 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -74895
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2415
         Width           =   10725
      End
      Begin VB.Data DataComp 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   105
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1680
         Width           =   10725
      End
      Begin MSDBGrid.DBGrid DBGTransSC 
         Bindings        =   "Fmtransc.frx":1834
         Height          =   4215
         Left            =   -74895
         OleObjectBlob   =   "Fmtransc.frx":184A
         TabIndex        =   8
         Top             =   420
         Width           =   10725
      End
      Begin MSDBGrid.DBGrid DBGComp 
         Bindings        =   "Fmtransc.frx":2213
         Height          =   1275
         Left            =   105
         OleObjectBlob   =   "Fmtransc.frx":2226
         TabIndex        =   9
         Top             =   420
         Width           =   10725
      End
      Begin MSDBGrid.DBGrid DBGBanco 
         Bindings        =   "Fmtransc.frx":2BE0
         Height          =   2010
         Left            =   -74895
         OleObjectBlob   =   "Fmtransc.frx":2BF4
         TabIndex        =   11
         Top             =   420
         Width           =   10725
      End
      Begin VB.Label LSHaber_ME 
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
         Height          =   330
         Left            =   -66075
         TabIndex        =   19
         Top             =   5040
         Width           =   1905
      End
      Begin VB.Label LSDebe_ME 
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
         Height          =   330
         Left            =   -67965
         TabIndex        =   18
         Top             =   5040
         Width           =   1905
      End
      Begin VB.Label LRet_MN 
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
         Height          =   330
         Left            =   -66075
         TabIndex        =   25
         Top             =   5040
         Width           =   1905
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTALES/MN"
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
         Left            =   -67440
         TabIndex        =   24
         Top             =   5040
         Width           =   1380
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTALES/ME"
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
         Left            =   -69330
         TabIndex        =   23
         Top             =   5040
         Width           =   1380
      End
      Begin VB.Label LSHaber_MN 
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
         Height          =   330
         Left            =   -71325
         TabIndex        =   22
         Top             =   5040
         Width           =   1905
      End
      Begin VB.Label LSDebe_MN 
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
         Height          =   330
         Left            =   -73215
         TabIndex        =   21
         Top             =   5040
         Width           =   1905
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TOTALES/MN"
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
         Left            =   -74895
         TabIndex        =   20
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label LHaber_ME 
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
         Height          =   330
         Left            =   8925
         TabIndex        =   13
         Top             =   5040
         Width           =   1905
      End
      Begin VB.Label LDebe_ME 
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
         Height          =   330
         Left            =   7035
         TabIndex        =   12
         Top             =   5040
         Width           =   1905
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTALES/ME"
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
         TabIndex        =   17
         Top             =   5040
         Width           =   1380
      End
      Begin VB.Label LHaber_MN 
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
         Height          =   330
         Left            =   3675
         TabIndex        =   16
         Top             =   5040
         Width           =   1905
      End
      Begin VB.Label LDebe_MN 
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
         Height          =   330
         Left            =   1785
         TabIndex        =   15
         Top             =   5040
         Width           =   1905
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTALES/MN"
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
         Left            =   420
         TabIndex        =   14
         Top             =   5040
         Width           =   1380
      End
   End
   Begin MSDBCtls.DBCombo DBCTP 
      Bindings        =   "Fmtransc.frx":35AB
      DataSource      =   "DataTP"
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   630
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   556
      _Version        =   327680
      Text            =   "DBCombo1"
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
   Begin MSDBCtls.DBList DBLNumero 
      Bindings        =   "Fmtransc.frx":35BC
      DataSource      =   "DataNumero"
      Height          =   1230
      Left            =   2205
      TabIndex        =   4
      Top             =   105
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2170
      _Version        =   327680
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
   Begin VB.Data DataNumero 
      Caption         =   "Numero"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9030
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   315
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data DataTP 
      Caption         =   "TP"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9030
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   630
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir de la Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   6300
      Picture         =   "Fmtransc.frx":35D1
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   105
      Width           =   1380
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Numero"
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
      Top             =   1050
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TIPO DE COMPROBANTE"
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
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   2010
   End
End
Attribute VB_Name = "FModTransSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload FModTransSC
End Sub

Private Sub Command2_Click()
  RatonReloj
  sSQL = "SELECT * " _
       & "FROM Transacciones " _
       & "WHERE TP = '" & DBCTP.Text & "' " _
       & "AND Numero = " & Val(SinEspaciosIzq(DBLNumero.Text)) & " " _
       & "ORDER BY Cta "
  SelectDBGrid DBGTrans, DataTrans, sSQL
  Debe = 0: Haber = 0: Debe_ME = 0: Haber_ME = 0
  With DataTrans.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .Fields("Debe")
          Haber = Haber + .Fields("Haber")
          Debe_ME = Debe_ME + .Fields("Debe_ME")
          Haber_ME = Haber_ME + .Fields("Haber_ME")
         .MoveNext
       Loop
   End If
  End With
  LDebe_MN.Caption = Format(Debe, "#,##0.00")
  LHaber_MN.Caption = Format(Haber, "#,##0.00")
  LDebe_ME.Caption = Format(Debe_ME, "#,##0.00")
  LHaber_ME.Caption = Format(Haber_ME, "#,##0.00")
  
  sSQL = "SELECT * " _
       & "FROM Comprobantes " _
       & "WHERE TP = '" & DBCTP.Text & "' " _
       & "AND Numero = " & Val(SinEspaciosIzq(DBLNumero.Text)) & " "
  SelectDBGrid DBGComp, DataComp, sSQL
  sSQL = "SELECT * " _
       & "FROM Bancos " _
       & "WHERE TP = '" & DBCTP.Text & "' " _
       & "AND Numero = " & Val(SinEspaciosIzq(DBLNumero.Text)) & " " _
       & "ORDER BY Cta_Banco "
  SelectDBGrid DBGBanco, DataBanco, sSQL
  sSQL = "SELECT * " _
       & "FROM Retenciones " _
       & "WHERE TP = '" & DBCTP.Text & "' " _
       & "AND Numero = " & Val(SinEspaciosIzq(DBLNumero.Text)) & " " _
       & "ORDER BY Cta "
  SelectDBGrid DBGRet, DataRet, sSQL
  Debe = 0: Haber = 0: Debe_ME = 0: Haber_ME = 0
  With DataRet.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Haber = Haber + .Fields("Valor_Retenido")
         .MoveNext
       Loop
   End If
  End With
  LRet_MN.Caption = Format(Haber, "#,##0.00")
  sSQL = "SELECT * " _
       & "FROM TransaccionesSC " _
       & "WHERE TP = '" & DBCTP.Text & "' " _
       & "AND Numero = " & Val(SinEspaciosIzq(DBLNumero.Text)) & " " _
       & "ORDER BY Cta,Codigo "
  SelectDBGrid DBGTransSC, DataTransSC, sSQL
  Debe = 0: Haber = 0: Debe_ME = 0: Haber_ME = 0
  With DataTransSC.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .Fields("Debitos")
          Haber = Haber + .Fields("Creditos")
          Debe_ME = Debe_ME + .Fields("Debitos_ME")
          Haber_ME = Haber_ME + .Fields("Creditos_ME")
         .MoveNext
       Loop
   End If
  End With
  LSDebe_MN.Caption = Format(Debe, "#,##0.00")
  LSHaber_MN.Caption = Format(Haber, "#,##0.00")
  LSDebe_ME.Caption = Format(Debe_ME, "#,##0.00")
  LSHaber_ME.Caption = Format(Haber_ME, "#,##0.00")
  RatonNormal
End Sub

Private Sub DBLNumero_DblClick()
  SiguienteControl
End Sub

Private Sub DBLNumero_GotFocus()
  sSQL = "SELECT (Numero & space(3) & Fecha) AS Numeros " _
       & "FROM Comprobantes " _
       & "WHERE TP = '" & DBCTP.Text & "' " _
       & "ORDER BY Numero "
  SelectDBList DBLNumero, DataNumero, sSQL, "Numeros"
End Sub

Private Sub DBLNumero_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT TP FROM Comprobantes " _
       & "GROUP BY TP "
  SelectDBCombo DBCTP, DataTP, sSQL, "TP", False
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FModTransSC
  DataTP.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataNumero.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataComp.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataBanco.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataRet.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataTrans.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataTransSC.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub

Private Sub TextNumero_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextNumero_LostFocus()
  RatonReloj
  sSQL = "SELECT * " _
       & "FROM Transacciones " _
       & "WHERE TP = '" & DBCTP.Text & "' " _
       & "AND Numero = " & Val(TextNumero.Text) & " " _
       & "ORDER BY Cta "
  SelectDBGrid DBGTrans, DataTrans, sSQL
  Debe = 0: Haber = 0: Debe_ME = 0: Haber_ME = 0
  With DataTrans.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .Fields("Debe")
          Haber = Haber + .Fields("Haber")
          Debe_ME = Debe_ME + .Fields("Debe_ME")
          Haber_ME = Haber_ME + .Fields("Haber_ME")
         .MoveNext
       Loop
   End If
  End With
  LDebe_MN.Caption = Format(Debe, "#,##0.00")
  LHaber_MN.Caption = Format(Haber, "#,##0.00")
  LDebe_ME.Caption = Format(Debe_ME, "#,##0.00")
  LHaber_ME.Caption = Format(Haber_ME, "#,##0.00")
  
  sSQL = "SELECT * " _
       & "FROM Comprobantes " _
       & "WHERE TP = '" & DBCTP.Text & "' " _
       & "AND Numero = " & Val(TextNumero.Text) & " "
  SelectDBGrid DBGComp, DataComp, sSQL
  sSQL = "SELECT * " _
       & "FROM Bancos " _
       & "WHERE TP = '" & DBCTP.Text & "' " _
       & "AND Numero = " & Val(TextNumero.Text) & " " _
       & "ORDER BY Cta_Banco "
  SelectDBGrid DBGBanco, DataBanco, sSQL
  sSQL = "SELECT * " _
       & "FROM Retenciones " _
       & "WHERE TP = '" & DBCTP.Text & "' " _
       & "AND Numero = " & Val(TextNumero.Text) & " " _
       & "ORDER BY Cta "
  SelectDBGrid DBGRet, DataRet, sSQL
  Debe = 0: Haber = 0: Debe_ME = 0: Haber_ME = 0
  With DataRet.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Haber = Haber + .Fields("Valor_Retenido")
         .MoveNext
       Loop
   End If
  End With
  LRet_MN.Caption = Format(Haber, "#,##0.00")
  sSQL = "SELECT * " _
       & "FROM TransaccionesSC " _
       & "WHERE TP = '" & DBCTP.Text & "' " _
       & "AND Numero = " & Val(TextNumero.Text) & " " _
       & "ORDER BY Cta,Codigo "
  SelectDBGrid DBGTransSC, DataTransSC, sSQL
  Debe = 0: Haber = 0: Debe_ME = 0: Haber_ME = 0
  With DataTransSC.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .Fields("Debitos")
          Haber = Haber + .Fields("Creditos")
          Debe_ME = Debe_ME + .Fields("Debitos_ME")
          Haber_ME = Haber_ME + .Fields("Creditos_ME")
         .MoveNext
       Loop
   End If
  End With
  LSDebe_MN.Caption = Format(Debe, "#,##0.00")
  LSHaber_MN.Caption = Format(Haber, "#,##0.00")
  LSDebe_ME.Caption = Format(Debe_ME, "#,##0.00")
  LSHaber_ME.Caption = Format(Haber_ME, "#,##0.00")
  RatonNormal
End Sub
