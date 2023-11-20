VERSION 5.00
Begin VB.Form FChangeValores 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtDetalle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5880
      MaxLength       =   60
      TabIndex        =   10
      Text            =   "0"
      Top             =   2310
      Width           =   3060
   End
   Begin VB.TextBox TxtConcepto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   105
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1155
      Width           =   8835
   End
   Begin VB.TextBox TxtHaber 
      Alignment       =   1  'Right Justify
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
      Left            =   3780
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   2310
      Width           =   2010
   End
   Begin VB.TextBox TxtDeposito 
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
      MaxLength       =   16
      TabIndex        =   4
      Text            =   "0"
      Top             =   2310
      Width           =   1485
   End
   Begin VB.TextBox TxtDebe 
      Alignment       =   1  'Right Justify
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
      Left            =   1680
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2310
      Width           =   2010
   End
   Begin VB.CommandButton Command1 
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
      Height          =   855
      Left            =   9030
      Picture         =   "ChangeValores.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command2 
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
      Left            =   9030
      Picture         =   "ChangeValores.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1050
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "DETALLE AUXILIAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   5880
      TabIndex        =   9
      Top             =   1995
      Width           =   3060
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FFFF&
      Caption         =   " CONCEPTO DEL COMPROBANTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   840
      Width           =   8835
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "VALOR DEL HABER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   3780
      TabIndex        =   7
      Top             =   1995
      Width           =   2010
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Cheq./Dep.No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   105
      TabIndex        =   3
      Top             =   1995
      Width           =   1485
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "VALOR DEL DEBE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   1680
      TabIndex        =   5
      Top             =   1995
      Width           =   2010
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ELIJA LA EMPRESA A COPIAR EL CATALOGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   8835
   End
End
Attribute VB_Name = "FChangeValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
     'MsgBox Producto
     Cadena = ""
     If TxtConcepto <> NomCta Then
        sSQL = "UPDATE Comprobantes " _
             & "SET Concepto = '" & TxtConcepto & "' " _
             & "WHERE TP = '" & Co.TP & "' " _
             & "AND Numero = " & Co.Numero & " " _
             & "AND Item = '" & Co.Item & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' "
        Ejecutar_SQL_SP sSQL
        Cadena = Cadena & "Concepto del Comprobante" & vbCrLf
     End If

     If CCur(TxtDebe) <> Debe Then
        sSQL = "UPDATE Transacciones " _
             & "SET Debe = " & Val(TxtDebe) & " " _
             & "WHERE TP = '" & Co.TP & "' " _
             & "AND Numero = " & Co.Numero & " " _
             & "AND Item = '" & Co.Item & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND ID = " & Asiento & " " _
             & "AND Cta = '" & Cta & "' "
        Ejecutar_SQL_SP sSQL
        Cadena = Cadena & "El Valor del Debe" & vbCrLf
     End If
     
     If CCur(TxtHaber) <> Haber Then
        sSQL = "UPDATE Transacciones " _
             & "SET Haber = " & Val(TxtHaber) & " " _
             & "WHERE TP = '" & Co.TP & "' " _
             & "AND Numero = " & Co.Numero & " " _
             & "AND Item = '" & Co.Item & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND ID = " & Asiento & " " _
             & "AND Cta = '" & Cta & "' "
        Ejecutar_SQL_SP sSQL
        Cadena = Cadena & "El Valor del Haber" & vbCrLf
     End If
     
     If TxtDeposito <> NoCheque Then
        sSQL = "UPDATE Transacciones " _
             & "SET Cheq_Dep = '" & TxtDeposito & "' " _
             & "WHERE TP = '" & Co.TP & "' " _
             & "AND Numero = " & Co.Numero & " " _
             & "AND Item = '" & Co.Item & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND ID = " & Asiento & " " _
             & "AND Cta = '" & Cta & "' "
        Ejecutar_SQL_SP sSQL
        Cadena = Cadena & "El Valor del Cheque o Deposito" & vbCrLf
     End If
         
     If TxtDetalle <> NomCtaSup Then
        sSQL = "UPDATE Transacciones " _
             & "SET Detalle = '" & TxtDetalle & "' " _
             & "WHERE TP = '" & Co.TP & "' " _
             & "AND Numero = " & Co.Numero & " " _
             & "AND Item = '" & Co.Item & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND ID = " & Asiento & " " _
             & "AND Cta = '" & Cta & "' "
        Ejecutar_SQL_SP sSQL
        Cadena = Cadena & "El Detalle de la Transaccion" & vbCrLf
     End If
     If Cadena = "" Then
        MsgBox "No se ha realizado ningun cambio"
     Else
        Actualiza_Procesado_Tabla "Transacciones", True
        Actualiza_Procesado_Tabla "Trans_SubCtas", True
        Actualiza_Procesado_Tabla "Trans_Kardex", True

        MsgBox "Proceso realizado, se actualizaron: " & vbCrLf _
             & Cadena _
             & "vuelva a listar el Comprobante para " _
             & "verificar los cambios realizados"
     End If
     Unload FChangeValores
End Sub

Private Sub Command2_Click()
  Unload FChangeValores
End Sub

Private Sub Form_Activate()
  If Cta <> Ninguno Then
     Label6.Caption = "FECHA: " & Co.Fecha & " DEL COMPROBANTE " & Co.TP & "-" & Co.Numero & vbCrLf & Cta & " - " & Cuenta_No
     TxtConcepto = NomCta
     TxtDebe = Debe
     TxtHaber = Haber
     TxtDetalle = NomCtaSup
     TxtDeposito = NoCheque
     TxtConcepto.SetFocus
  Else
     MsgBox "No existe Cuenta para cambiar"
     Unload FChangeValores
  End If
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FChangeValores
End Sub

Private Sub TxtConcepto_GotFocus()
   MarcarTexto TxtConcepto
End Sub

Private Sub TxtConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtConcepto_LostFocus()
   TextoValido TxtConcepto
End Sub

Private Sub TxtDebe_GotFocus()
   MarcarTexto TxtDebe
End Sub

Private Sub TxtDebe_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtDebe_LostFocus()
   TextoValido TxtDebe, True, , 2
End Sub

Private Sub TxtDetalle_GotFocus()
   MarcarTexto TxtDetalle
End Sub

Private Sub TxtDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtDetalle_LostFocus()
   MarcarTexto TxtDetalle
End Sub

Private Sub TxtHaber_GotFocus()
   MarcarTexto TxtHaber
End Sub

Private Sub TxtHaber_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtHaber_LostFocus()
   TextoValido TxtHaber, True, , 2
End Sub

Private Sub TxtDeposito_GotFocus()
   MarcarTexto TxtDeposito
End Sub

Private Sub TxtDeposito_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

