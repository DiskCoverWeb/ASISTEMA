VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FActivar 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ACTIVACION O RENOVACION DE LA CLAVE DEL SISTEMA"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CommandButton3 
      Caption         =   "&Archivo Llave"
      Height          =   435
      Left            =   105
      TabIndex        =   4
      Top             =   105
      Width           =   1905
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   525
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton CommandButton17 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   8505
      TabIndex        =   7
      Top             =   3150
      Width           =   1380
   End
   Begin VB.CommandButton CommandButton1 
      Caption         =   "&Listado Llaves"
      Height          =   435
      Left            =   6930
      TabIndex        =   6
      Top             =   3150
      Width           =   1485
   End
   Begin VB.CommandButton CommandButton16 
      Caption         =   "&Renovar Llave"
      Height          =   435
      Left            =   5040
      TabIndex        =   5
      Top             =   3150
      Width           =   1800
   End
   Begin VB.TextBox TxtResult 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2430
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "FActivar.frx":0000
      Top             =   630
      Width           =   9780
   End
   Begin MSComDlg.CommonDialog CDialogDir 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSMask.MaskEdBox MBPeriodo 
      Height          =   435
      Left            =   3360
      TabIndex        =   3
      Top             =   3150
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   767
      _Version        =   393216
      BackColor       =   16744576
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
   Begin VB.Label LblLogin 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   2100
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   7785
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Fecha de Vencimiento del Contrato:"
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
      Top             =   3255
      Width           =   3270
   End
End
Attribute VB_Name = "FActivar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'carla.jativa@ dhl.com

Option Explicit

Private Sub CommandButton1_Click()
Dim IniX As Single
Dim IniY As Single
Dim Texto As String
Dim LineTexto As String
Dim CampoTexto As String
Dim AnchoDeLinea As Single
On Error GoTo Errorhandler
RatonReloj
SQLMsg1 = ""
SQLMsg2 = ""
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
Escala_Centimetro 1, TipoArialNarrow, 8
'Iniciamos la impresion
AnchoDeLinea = 18: Pagina = 1
IniX = 1: IniY = 1
Empresa = "DISKCOVER SYSTEM"
Direccion = "REGISTRO DE CLAVES ACTIVADAS"
'EncabezadoDocumento IniX, IniY, "DEL: " & FechaStrg(FechaSistema)
Printer.FontSize = 7
Printer.FontName = TipoSystem
LineTexto = TxtResult.Text
Texto = LineTexto
J = Len(Texto)
I = 1: K = 1
LineTexto = ""
Do While I < J
   Caracter = Mid(Texto, I, 1)
   CampoTexto = Mid(Texto, I, 3)
   LineTexto = LineTexto & Caracter
   If Printer.TextWidth(LineTexto) > AnchoDeLinea Or Asc(Caracter) = 13 Or Asc(Caracter) = 10 Then
      If Printer.TextWidth(LineTexto) > AnchoDeLinea Then
         K = Len(LineTexto)
         If K > 0 Then
            Do
               K = K - 1
               I = I - 1
            Loop Until K < 2 Or Mid(LineTexto, K, 1) = " "
            LineTexto = Mid(LineTexto, 1, K)
         End If
      End If
      'MsgBox IniX & vbCrLf & PosLinea & vbCrLf & LineTexto
      Printer.CurrentX = IniX
      Printer.CurrentY = PosLinea
      Printer.Print LineTexto
      PosLinea = PosLinea + Printer.TextHeight("H") + 0.1
      LineTexto = ""
      If Asc(Caracter) = 13 Then I = I + 1
   End If
   If PosLinea >= LimiteAlto Then
      Printer.NewPage
      PosLinea = IniY + 2
      LineTexto = ""
   End If
   I = I + 1
Loop
'Producto = InsertarLinea
Printer.EndDoc
RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub CommandButton16_Click()
  If RutaOrigen <> "" Then
     FechaValida MBPeriodo
     If LineasLogIn(3) <> "." Then
        LineasLogIn(2) = MBPeriodo
        Cadena1 = "Renovación de la Clave Exitosa"
     Else
        LineasLogIn(2) = MBPeriodo
        LineasLogIn(3) = MBPeriodo
        Cadena1 = "Activacion de la Clave Exitosa"
     End If
     Cadena = ""
     For I = 1 To ContadorLogIn
         Cadena = Cadena & LineasLogIn(I) & " ^ "
     Next I
     Escribir_Archivo RutaOrigen, Crear_Encriptado(Cadena)
     MsgBox Cadena1
  End If
  Unload Me
End Sub

Private Sub CommandButton17_Click()
 Unload Me
End Sub

Private Sub CommandButton3_Click()
  'OpenZip
  Cadena = ""
  RutaOrigen = UCase(SelectZipFile(CDialogDir, SelectAll))
  Cadena = Leer_Archivo(RutaOrigen)
  Cadena = Leer_Encriptado(Cadena)
  'MsgBox Cadena
  'FechaIni = Trim(Mid(Cadena, 1, 10))
  LblLogin.Caption = Cadena
  TxtResult.Text = ""
  For I = 1 To ContadorLogIn
      TxtResult.Text = TxtResult.Text & LineasLogIn(I) & vbCrLf & String(128, "_") & vbCrLf & vbCrLf
  Next I
  If LineasLogIn(2) <> "." And Len(LineasLogIn(2)) = 10 Then MBPeriodo = LineasLogIn(2)
  MBPeriodo.SetFocus
End Sub

Private Sub Form_Activate()
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FActivar
End Sub

Private Sub MBPeriodo_GotFocus()
  MarcarTexto MBPeriodo
End Sub

Private Sub MBPeriodo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBPeriodo_LostFocus()
  FechaValida MBPeriodo
End Sub
