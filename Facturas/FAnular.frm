VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FAnulacion 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ANULACION DE FACTURAS"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "&Anulación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8610
      Picture         =   "FAnular.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   105
      Width           =   1065
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
      Height          =   855
      Left            =   8610
      Picture         =   "FAnular.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1050
      Width           =   1065
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   7350
      TabIndex        =   2
      Top             =   105
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
   End
   Begin VB.Label LblSaldo 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   6405
      TabIndex        =   12
      Top             =   1365
      Width           =   2115
   End
   Begin VB.Label LblTotal 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   4305
      TabIndex        =   10
      Top             =   1365
      Width           =   2010
   End
   Begin VB.Label LblIVA 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   2310
      TabIndex        =   8
      Top             =   1365
      Width           =   1905
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SALDO PENDIENTE"
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
      Left            =   6405
      TabIndex        =   11
      Top             =   1050
      Width           =   2115
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TOTAL"
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
      Left            =   4305
      TabIndex        =   9
      Top             =   1050
      Width           =   2010
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " I.V.A"
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
      TabIndex        =   7
      Top             =   1050
      Width           =   1905
   End
   Begin VB.Label LblSubTotal 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Top             =   1365
      Width           =   2115
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SUBTOTAL"
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
      TabIndex        =   5
      Top             =   1050
      Width           =   2115
   End
   Begin VB.Label LblAnular 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1485
      Left            =   105
      TabIndex        =   13
      Top             =   1785
      Width           =   8415
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6405
      TabIndex        =   1
      Top             =   105
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   6930
      TabIndex        =   4
      Top             =   525
      Width           =   1590
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      TabIndex        =   3
      Top             =   525
      Width           =   6840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cliente"
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
      Width           =   6210
   End
End
Attribute VB_Name = "FAnulacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Idx As Long

Private Sub Command2_Click()
  FechaValida MBoxFecha
  If FA.Factura > 0 Then
     FA.Fecha_NC = MBoxFecha
     FA.Serie_NC = NC.Serie
     FA.Autorizacion_NC = NC.Autorizacion
     Titulo = "FORMULARIO DE ANULACION"
     Mensajes = "Esta seguro que desea proceder," & vbCrLf _
              & "con la Factura No. " & FA.Factura
     If BoxMensaje = vbYes Then
        RatonReloj
        Control_Procesos "A", "Anulación de " & FA.TC & "-" & FA.Serie & " No. " & Format(FA.Factura, "000000000")
        ConceptoComp = "Anulación de la " & FA.TC & "-" & FA.Serie & " No. " & FA.Factura & ", del Cliente: " & NombreCliente
        sSQL = "DELETE * " _
             & "FROM Trans_Abonos " _
             & "WHERE Factura = " & FA.Factura & " " _
             & "AND TP = '" & FA.TC & "' " _
             & "AND Serie = '" & FA.Serie & "' " _
             & "AND Autorizacion = '" & FA.Autorizacion & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' "
        Ejecutar_SQL_SP sSQL
           
       'Borramos las facturas del kardex de anulacion
        sSQL = "DELETE * " _
             & "FROM Trans_Kardex " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND TC = '" & FA.TC & "' " _
             & "AND Serie = '" & FA.Serie & "' " _
             & "AND Factura = " & FA.Factura & " "
        Ejecutar_SQL_SP sSQL
           
        sSQL = "UPDATE Facturas " _
             & "SET T = 'A', Nota = '" & ConceptoComp _
             & "',Fecha_C = #" & BuscarFecha(MBoxFecha) & "#," _
             & "Saldo_MN = " & Total_Saldos_ME & " " _
             & "WHERE Factura = " & FA.Factura & " " _
             & "AND TC = '" & FA.TC & "' " _
             & "AND Serie = '" & FA.Serie & "' " _
             & "AND Autorizacion = '" & FA.Autorizacion & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' "
        Ejecutar_SQL_SP sSQL

        sSQL = "UPDATE Detalle_Factura " _
             & "SET T = 'A' " _
             & "WHERE Factura = " & FA.Factura & " " _
             & "AND TC = '" & FA.TC & "' " _
             & "AND Serie = '" & FA.Serie & "' " _
             & "AND Autorizacion = '" & FA.Autorizacion & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' "
        Ejecutar_SQL_SP sSQL
     End If
  End If
  RatonNormal
  Unload FAnulacion
End Sub

Private Sub Command3_Click()
  Unload FAnulacion
End Sub

Private Sub Form_Activate()
   'MsgBox FA.T & vbCrLf & FA.Sin_IVA & vbCrLf & FA.Con_IVA & vbCrLf & FA.Saldo_MN
    If Len(FA.Cta_CxP) > 2 Then NC.Cta_CxP = FA.Cta_CxP
    LblSubTotal.Caption = Format$(FA.SubTotal - FA.Descuento - FA.Descuento2, "#,##0.00")
    LblIVA.Caption = Format$(FA.Total_IVA, "#,##0.00")
    LblTotal.Caption = Format$(FA.Total_MN, "#,##0.00")
    LblSaldo.Caption = Format$(FA.Saldo_Actual, "#,##0.00")
    Label1.Caption = FA.Cliente
    Select Case FA.TC
      Case "FA": Label3.Caption = "FACTURA " & FA.Serie & "-"
      Case "NV": Label3.Caption = "NOTA DE VENTA " & FA.Serie & "-"
      Case "DO": Label3.Caption = "DONACION " & FA.Serie & "-"
      Case Else: Label3.Caption = "DOCUMENTO " & FA.Serie & "-"
    End Select
    Label2.Caption = Format$(FA.Factura, "000000000")
   'Si es anulacion directa o por NC
    LblAnular.Caption = "SI YA REALIZO EL CIERRE DE CAJA, AL ANULAR ESTA FACTURA TENDRA QUE " _
                      & "VOLVER A REALIZAR EL CIERRE DEL DIA DE EMISION DE LA FACTURA." & vbCrLf _
                      & "SOLO SE PUEDE ANULAR FACTURAS SI SE TIENE PRESENTE LA FACTURA ORIGINAL Y COPIA, " _
                      & "SI ES ELECTRONICA SE DEBE COMUNICAR AL CLIENTE DE LA ANULACION." & vbCrLf
    Command2.Caption = "&Anulación"
    Label2.Caption = FA.Factura
    Label1.Caption = FA.Cliente
    Select Case FA.TC
      Case "PV": Label3.Caption = " Punto de Venta No."
      Case "NV": Label3.Caption = " Nota de Venta No. " & FA.Autorizacion & "-" & FA.Serie & "-"
      Case Else: Label3.Caption = " Factura No. " & FA.Autorizacion & "-" & FA.Serie & "-"
    End Select
End Sub

Private Sub Form_Load()
  CentrarForm FAnulacion
End Sub

Private Sub MBoxFecha_GotFocus()
  MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_KeyDown(KeyCode As Integer, Shift As Integer)
 Keys_Especiales Shift
 PresionoEnter KeyCode
' If CtrlDown And KeyCode = vbKeyF12 Then
'    FA.Porc_NC = 0.12
'    MsgBox "Proceda a registrar la NC"
' End If
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha, True
  Validar_Porc_IVA MBoxFecha
  NC.Fecha = MBoxFecha
End Sub

