VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form IngSubCtaK 
   BorderStyle     =   0  'None
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame FrmSubCta 
      Caption         =   "SUB CUENTA"
      Height          =   2385
      Left            =   108
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4440
      Begin VB.ListBox LstSubCta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   108
         TabIndex        =   1
         Top             =   216
         Width           =   4224
      End
   End
   Begin MSAdodcLib.Adodc AdoSubCta 
      Height          =   336
      Left            =   216
      Top             =   432
      Visible         =   0   'False
      Width           =   2016
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
End
Attribute VB_Name = "IngSubCtaK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
  sSQL = "SELECT * " _
       & "FROM Catalogo_SubCtas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = '" & SubCta & "' " _
       & "ORDER BY Detalle "
 SelectAdodc AdoSubCta, sSQL
 LstSubCta.Clear
 LstSubCta.AddItem "Ninguna SubCuenta"
 With AdoSubCta.Recordset
  If .RecordCount > 0 Then
      IngSubCtaK.Visible = True
      FrmSubCta.Visible = True
      Do While Not .EOF
         LstSubCta.AddItem .Fields("Detalle") & " (" & .Fields("TC") & ")"
        .MoveNext
      Loop
      LstSubCta.Text = LstSubCta.List(0)
      LstSubCta.SetFocus
  Else
      SubCtaGen = Ninguno
      Unload Me
  End If
 End With
End Sub

Private Sub Form_Load()
  CentrarForm IngSubCtaK
  ConectarAdodc AdoSubCta
End Sub

Private Sub LstSubCta_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
     SubCtaGen = Ninguno
     Unload Me
  End If
  If KeyCode = vbKeyReturn Then
     SubCtaGen = Ninguno
     With AdoSubCta.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
          Cadena = SinEspaciosDer(LstSubCta.Text)
          Cadena = Trim(Mid(LstSubCta.Text, 1, Len(LstSubCta.Text) - Len(Cadena)))
         .Find ("Detalle = '" & Cadena & "' ")
          If Not .EOF Then SubCtaGen = .Fields("Codigo")
      End If
     End With
     Unload Me
  End If
  
End Sub

