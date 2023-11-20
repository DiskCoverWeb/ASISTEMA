VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FAutorizarNotas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AUTORIZACION DE NOTAS"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox LstAutoriza 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3300
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   420
      Width           =   5370
   End
   Begin MSAdodcLib.Adodc AdoAutorizar 
      Height          =   330
      Left            =   4830
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
      Caption         =   "Autorizar"
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
      Height          =   750
      Left            =   5565
      Picture         =   "FAutNota.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   945
      Width           =   1170
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Grabar"
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
      Left            =   5565
      Picture         =   "FAutNota.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   105
      Width           =   1170
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Ingresar Notas de"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Width           =   5370
   End
End
Attribute VB_Name = "FAutorizarNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
  Unload FAutorizarNotas
End Sub

Private Sub Command5_Click()
Dim IAut As Integer
Dim Visto As Boolean
   With AdoAutorizar.Recordset
    If .RecordCount > 0 Then
        For IAut = 0 To LstAutoriza.ListCount - 1
           .Fields(Autoizar_Notas(IAut)) = VerdadFalso(LstAutoriza.Selected(IAut))
        Next IAut
       .Update
    End If
   End With
   RatonNormal
   Unload FAutorizarNotas
End Sub

Private Sub Form_Activate()
Dim IAut As Integer
   For IAut = 0 To 11
       Autoizar_Notas(IAut) = Ninguno
   Next IAut
   IAut = 0
   sSQL = "SELECT * " _
        & "FROM Catalogo_Periodo_Lectivo " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' "
   SelectAdodc AdoAutorizar, sSQL
   With AdoAutorizar.Recordset
    If .RecordCount > 0 Then
        If Mid$(FormatoLibreta, 1, 9) = "TRIMESTRE" Then
           LstAutoriza.AddItem "Primer Trimestre Primer Periodo", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NPQP1"))
           Autoizar_Notas(IAut) = "NPQP1": IAut = IAut + 1
           LstAutoriza.AddItem "Primer Trimestre Segundo Periodo", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NPQP2"))
           Autoizar_Notas(IAut) = "NPQP2": IAut = IAut + 1
           LstAutoriza.AddItem "Examen Primer Trimestre", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NPQEX"))
           Autoizar_Notas(IAut) = "NPQEX": IAut = IAut + 1
           
           LstAutoriza.AddItem "Segundo Trimestre Primer Periodo", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NSQP1"))
           Autoizar_Notas(IAut) = "NSQP1": IAut = IAut + 1
           LstAutoriza.AddItem "Segundo Trimestre Segundo Periodo", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NSQP2"))
           Autoizar_Notas(IAut) = "NSQP2": IAut = IAut + 1
           LstAutoriza.AddItem "Examen Segundo Trimestre", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NSQEX"))
           Autoizar_Notas(IAut) = "NSQEX": IAut = IAut + 1
           
           LstAutoriza.AddItem "Tercer Trimestre Primer Periodo", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NTQP1"))
           Autoizar_Notas(IAut) = "NTQP1": IAut = IAut + 1
           LstAutoriza.AddItem "Tercer Trimestre Segundo Periodo", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NTQP2"))
           Autoizar_Notas(IAut) = "NTQP2": IAut = IAut + 1
           LstAutoriza.AddItem "Examen Tercer Trimestre", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NTQEX"))
           Autoizar_Notas(IAut) = "NTQEX": IAut = IAut + 1
        ElseIf Mid$(FormatoLibreta, 1, 9) = "QUIMESTRE" Then
           LstAutoriza.AddItem "Primer Quimestre Primer Parcial", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NPQP1"))
           Autoizar_Notas(IAut) = "NPQP1": IAut = IAut + 1
           LstAutoriza.AddItem "Primer Quimestre Segundo Parcial", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NPQP2"))
           Autoizar_Notas(IAut) = "NPQP2": IAut = IAut + 1
           LstAutoriza.AddItem "Primer Quimestre Tercer Parcial", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NPQP3"))
           Autoizar_Notas(IAut) = "NPQP3": IAut = IAut + 1
           LstAutoriza.AddItem "Examen Primer Quimestre", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NPQEX"))
           Autoizar_Notas(IAut) = "NPQEX": IAut = IAut + 1
           
           LstAutoriza.AddItem "Segundo Quimestre Primer Parcial", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NSQP1"))
           Autoizar_Notas(IAut) = "NSQP1": IAut = IAut + 1
           LstAutoriza.AddItem "Segundo Quimestre Segundo Parcial", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NSQP2"))
           Autoizar_Notas(IAut) = "NSQP2": IAut = IAut + 1
           LstAutoriza.AddItem "Segundo Quimestre Tercer Parcial", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NSQP3"))
           Autoizar_Notas(IAut) = "NSQP3": IAut = IAut + 1
           LstAutoriza.AddItem "Examen Segundo Quimestre", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NSQEX"))
           Autoizar_Notas(IAut) = "NSQEX": IAut = IAut + 1
        Else
           LstAutoriza.AddItem "Primer Quimestre Primer Periodo", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NPQP1"))
           Autoizar_Notas(IAut) = "NPQP1": IAut = IAut + 1
           LstAutoriza.AddItem "Primer Quimestre Segundo Periodo", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NPQP2"))
           Autoizar_Notas(IAut) = "NPQP2": IAut = IAut + 1
           LstAutoriza.AddItem "Examen Primer Quimestre", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NPQEX"))
           Autoizar_Notas(IAut) = "NPQEX": IAut = IAut + 1
           
           LstAutoriza.AddItem "Segundo Quimestre Primer Periodo", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NSQP1"))
           Autoizar_Notas(IAut) = "NSQP1": IAut = IAut + 1
           LstAutoriza.AddItem "Segundo Quimestre Segundo Periodo", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NSQP2"))
           Autoizar_Notas(IAut) = "NSQP2": IAut = IAut + 1
           LstAutoriza.AddItem "Examen Segundo Quimestre", IAut
           LstAutoriza.Selected(IAut) = CBool(.Fields("NSQEX"))
           Autoizar_Notas(IAut) = "NSQEX": IAut = IAut + 1
        End If
        LstAutoriza.AddItem "Supletorio", IAut
        LstAutoriza.Selected(IAut) = CBool(.Fields("NSUPL"))
        Autoizar_Notas(IAut) = "NSUPL": IAut = IAut + 1
        LstAutoriza.AddItem "Remedial", IAut
        LstAutoriza.Selected(IAut) = CBool(.Fields("NREME"))
        Autoizar_Notas(IAut) = "NREME": IAut = IAut + 1
        LstAutoriza.AddItem "Grado", IAut
        LstAutoriza.Selected(IAut) = CBool(.Fields("NGRADO"))
        Autoizar_Notas(IAut) = "NGRADO": IAut = IAut + 1
    End If
   End With
   RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm FAutorizarNotas
   ConectarAdodc AdoAutorizar
End Sub

