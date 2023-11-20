VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FFileMat 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   19.42
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   26.882
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Generar &Paralelos"
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
      Left            =   105
      Picture         =   "FFileMat.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1155
      Width           =   1065
   End
   Begin VB.PictureBox PictLetrasH 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5685
      Left            =   1260
      ScaleHeight     =   10.028
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   5.398
      TabIndex        =   2
      Top             =   525
      Width           =   3060
   End
   Begin VB.PictureBox PictLetrasV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5685
      Left            =   4725
      ScaleHeight     =   10.028
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   5.768
      TabIndex        =   1
      Top             =   525
      Width           =   3270
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Generar &Materias"
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
      Left            =   105
      Picture         =   "FFileMat.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   105
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   3045
      Top             =   105
      Visible         =   0   'False
      Width           =   1725
      _ExtentX        =   3043
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
      Caption         =   "Aux"
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
   Begin MSAdodcLib.Adodc AdoAux1 
      Height          =   330
      Left            =   1260
      Top             =   105
      Visible         =   0   'False
      Width           =   1725
      _ExtentX        =   3043
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
      Caption         =   "Aux1"
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
Attribute VB_Name = "FFileMAt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim PointColor
Dim PictTexto, Texto, CampoTexto, LineaDeTexto As String
Dim AnchoDeLinea As Single
Dim IJ As Integer
Dim AnchoMaximo As Single
Dim AltoMaximo As Single
Dim TipoLetra As String
Dim PorteLetra As String
' Empezamos la escritura de las letras
PictLetrasH.Visible = False
PictLetrasV.Visible = False
PictLetrasH.AutoRedraw = True
PictLetrasV.AutoRedraw = True
sSQL = "SELECT * " _
     & "FROM Catalogo_Estudiantil " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND TC = 'M' " _
     & "ORDER BY CodigoE  "
SelectAdodc AdoAux, sSQL
RatonReloj

PictTexto = ""
Contador = 0
AnchoMaximo = 0
AltoMaximo = 0

With AdoAux.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     TipoLetra = TipoTimes
     PorteLetra = 8
     AltoMaximo = 15.5
     AnchoMaximo = 3
     'MsgBox Codigo & vbCrLf & PictTexto & vbCrLf & AnchoMaximo & vbCrLf & AltoMaximo
     PictTexto = ""
     Contador = 0
     Codigo4 = Mid(.Fields("CodigoE"), 1, Len(.Fields("CodigoE")) - 3)
     Codigo = Mid(Codigo4, 1, 1) & Mid(Codigo4, 3, 2) & Mid(Codigo4, 6, 2)
     
     Do While Not .EOF
        NomCta = .Fields("Detalle")
        If Codigo4 <> Mid(.Fields("CodigoE"), 1, Len(.Fields("CodigoE")) - 3) Then
           Codigo = Mid(Codigo4, 1, 1) & Mid(Codigo4, 3, 2) & Mid(Codigo4, 6, 2)
           'MsgBox Codigo
           PictLetrasH.AutoRedraw = True
           PictLetrasH.Picture = LoadPicture()
           PictLetrasV.AutoRedraw = True
           PictLetrasV.Picture = LoadPicture()
           PictLetrasH.FontBold = False: PictLetrasV.FontBold = False
           PictLetrasH.Font = TipoLetra: PictLetrasV.Font = TipoLetra
           PictLetrasH.FontSize = PorteLetra: PictLetrasV.FontSize = PorteLetra
           
           PictLetrasH.Width = AnchoMaximo
           PictLetrasH.Height = AltoMaximo
           PictLetrasV.Height = PictLetrasH.Width
           PictLetrasV.Width = PictLetrasH.Height
           'MsgBox PictTexto & "...."
           AltoLetra = PictLetrasH.TextHeight(Mid(PictTexto, 1, 1))
           PCol = 0.1: Msg = "": PFil = 0.01
           For I = 1 To Len(PictTexto)
             If Mid(PictTexto, I, 1) = vbCr Then
                Texto = SinEspaciosIzq(Msg)
                PictLetrasH.CurrentX = PCol: PictLetrasH.CurrentY = PFil + 0.1
                PictLetrasH.Print Texto
                PictLetrasH.CurrentX = PCol: PictLetrasH.CurrentY = PFil + 0.4
                PictLetrasH.Print Mid(Msg, Len(Texto) + 2, Len(Msg))
                PictLetrasH.Line (0, PFil + 0.8)-(AnchoMaximo, PFil + 0.8), QBColor(0)
                Msg = ""
                PFil = PFil + 0.4 + AltoLetra
                I = I + 2
             End If
             Msg = Msg & Mid(PictTexto, I, 1)
           Next I
          RatonReloj
          JR = 0
          Do While JR <= PictLetrasH.ScaleHeight
             IR = 0
             Do While IR <= PictLetrasH.ScaleWidth
                PointColor = PictLetrasH.Point(IR, JR)
                If PointColor <> CLng(&HFFFFFF) Then PictLetrasV.PSet (JR, PictLetrasH.ScaleWidth - IR), PointColor
                IR = IR + (PorteLetra / 720)
             Loop
             JR = JR + (PorteLetra / 720)
          Loop
          Beep
          FFileMAt.Caption = Codigo & ": (" & IR & ")(" & JR & ")(" & PointColor & ")"
          'MsgBox "..."
          PictLetrasV.Line (0.01, 0.01)-(PictLetrasV.Width - 0.01, PictLetrasV.Height - 0.01), QBColor(0), B
          SavePicture PictLetrasV.Image, RutaSistema & "\FORMATOS\PARALELO\P" & Codigo & ".BMP"
          Codigo4 = Mid(.Fields("CodigoE"), 1, Len(.Fields("CodigoE")) - 3)
          PictTexto = ""
          Contador = 0
          
          
'''          RatonNormal
'''          PictLetrasH.Visible = True
'''          PictLetrasV.Visible = True
'''          Exit Sub
        End If
        Contador = Contador + 1
        PictTexto = PictTexto & .Fields("Detalle") & vbCrLf
        RatonNormal
       .MoveNext
     Loop
     
                Codigo = Mid(Codigo4, 1, 1) & Mid(Codigo4, 3, 2) & Mid(Codigo4, 6, 2)
           'MsgBox Codigo
           PictLetrasH.AutoRedraw = True
           PictLetrasH.Picture = LoadPicture()
           PictLetrasV.AutoRedraw = True
           PictLetrasV.Picture = LoadPicture()
           PictLetrasH.FontBold = False: PictLetrasV.FontBold = False
           PictLetrasH.Font = TipoLetra: PictLetrasV.Font = TipoLetra
           PictLetrasH.FontSize = PorteLetra: PictLetrasV.FontSize = PorteLetra
           
           PictLetrasH.Width = AnchoMaximo
           PictLetrasH.Height = AltoMaximo
           PictLetrasV.Height = PictLetrasH.Width
           PictLetrasV.Width = PictLetrasH.Height
           'MsgBox PictTexto & "...."
           AltoLetra = PictLetrasH.TextHeight(Mid(PictTexto, 1, 1))
           PCol = 0.1: Msg = "": PFil = 0.01
           For I = 1 To Len(PictTexto)
             If Mid(PictTexto, I, 1) = vbCr Then
                Texto = SinEspaciosIzq(Msg)
                PictLetrasH.CurrentX = PCol: PictLetrasH.CurrentY = PFil + 0.1
                PictLetrasH.Print Texto
                PictLetrasH.CurrentX = PCol: PictLetrasH.CurrentY = PFil + 0.4
                PictLetrasH.Print Mid(Msg, Len(Texto) + 2, Len(Msg))
                PictLetrasH.Line (0, PFil + 0.8)-(AnchoMaximo, PFil + 0.8), QBColor(0)
                Msg = ""
                PFil = PFil + 0.4 + AltoLetra
                I = I + 2
             End If
             Msg = Msg & Mid(PictTexto, I, 1)
           Next I
          RatonReloj
          IR = 0
          Do While IR <= PictLetrasH.ScaleWidth
             JR = 0
             Do While JR <= PictLetrasH.ScaleHeight
                PointColor = PictLetrasH.Point(IR, JR)
                If PointColor <> CLng(&HFFFFFF) Then PictLetrasV.PSet (JR, PictLetrasH.ScaleWidth - IR), PointColor
                JR = JR + (PorteLetra / 720)
             Loop
             IR = IR + (PorteLetra / 720)
          Loop
          Beep
          FFileMAt.Caption = Codigo & ": (" & IR & ")(" & JR & ")(" & PointColor & ")"
          'MsgBox "..."
          PictLetrasV.Line (0.01, 0.01)-(PictLetrasV.Width - 0.01, PictLetrasV.Height - 0.01), QBColor(0), B
          SavePicture PictLetrasV.Image, RutaSistema & "\FORMATOS\PARALELO\P" & Codigo & ".BMP"

 End If
End With
PictLetrasH.Visible = True
PictLetrasV.Visible = True
RatonNormal
End Sub

Private Sub Command18_Click()
Dim PointColor
Dim PictTexto, Texto, CampoTexto, LineaDeTexto As String
Dim AnchoDeLinea As Single
Dim IJ As Integer
' Empezamos la escritura de las letras
PictLetrasH.Visible = False
PictLetrasV.Visible = False
PictLetrasH.AutoRedraw = True
PictLetrasV.AutoRedraw = True
sSQL = "SELECT * " _
     & "FROM Catalogo_Materias " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "ORDER BY CodMat "
SelectAdodc AdoAux, sSQL
RatonReloj
With AdoAux.Recordset
 If .RecordCount > 0 Then
     Do While Not .EOF
        Codigo = .Fields("CodMat"): NomCta = .Fields("Materia")
        FFileMAt.Caption = NomCta & ":"
        PictLetrasH.AutoRedraw = True
        PictLetrasH.Picture = LoadPicture()
        PictLetrasV.AutoRedraw = True
        PictLetrasV.Picture = LoadPicture()
        PictLetrasH.FontBold = True: PictLetrasV.FontBold = True
        PictLetrasH.Font = TipoComicSans: PictLetrasV.Font = TipoComicSans
        PictLetrasH.FontSize = 10: PictLetrasV.FontSize = 10
        
        PictTexto = "Quimestres" & vbCrLf _
                  & "Promedio Global" & vbCrLf _
                  & "Examen Supletorio" & vbCrLf _
                  & "Promedio Total" & vbCrLf
        AltoLetra = PictLetrasH.TextHeight(PictTexto)
        AnchoMax = PictLetrasH.TextWidth(PictTexto)
        PictLetrasH.Width = Round(AnchoMax) + 1.45
        PictLetrasH.Height = Round(AltoLetra) + 0.85
        PictLetrasV.Height = PictLetrasH.Width
        PictLetrasV.Width = PictLetrasH.Height
                
        PictLetrasH.FontBold = True: PictLetrasV.FontBold = True
        PictLetrasH.Font = TipoComicSans: PictLetrasV.Font = TipoComicSans
        PictLetrasH.FontSize = 10: PictLetrasV.FontSize = 10
          
        AltoLetra = PictLetrasH.TextHeight(Mid(PictTexto, 1, 1))
        PictLetrasH.Line (AnchoMax + 0.2, 0)-(AnchoMax + 0.2, PictLetrasH.Height), QBColor(0)
        PCol = 0.1: Msg = "": PFil = 0.1
        For I = 1 To Len(PictTexto)
         If Mid(PictTexto, I, 1) = vbCr Then
            PictLetrasH.CurrentX = PCol: PictLetrasH.CurrentY = PFil
            PictLetrasH.Print Msg
            PictLetrasH.Line (0, PFil + 0.8)-(AnchoMax + 0.2, PFil + 0.8), QBColor(0)
           'MsgBox Msg
            Msg = ""
            PFil = PFil + 0.45 + AltoLetra
            I = I + 2
         End If
         Msg = Msg & Mid(PictTexto, I, 1)
        Next I
        RatonReloj
        IR = 0
        Do While IR < PictLetrasH.ScaleWidth
           JR = 0
           Do While JR < PictLetrasH.ScaleHeight
              PointColor = PictLetrasH.Point(IR, JR)
              If PointColor <> CLng(&HFFFFFF) Then PictLetrasV.PSet (JR, PictLetrasH.ScaleWidth - IR - 0.1), PointColor
              JR = JR + 0.005
           Loop
           IR = IR + 0.005
        Loop
        FFileMAt.Caption = NomCta & ": (" & IR & ")(" & JR & ")(" & PointColor & ")"
        AnchoDeLinea = PictLetrasH.Width + 0.1
        PosLinea = 0.01
        Texto = SinEspaciosIzq(NomCta)
        PictLetrasV.CurrentX = 0.1
        PictLetrasV.CurrentY = 0.01
        PictLetrasV.Print Texto
        
        PictLetrasV.CurrentX = 0.1
        PictLetrasV.CurrentY = 0.45
        PictLetrasV.Print Mid(NomCta, Len(Texto) + 2, Len(NomCta))
        PictLetrasV.Line (0.01, 0.01)-(PictLetrasV.Width - 0.18, PictLetrasV.Height - 0.16), QBColor(0), B
        
        SavePicture PictLetrasV.Image, RutaSistema & "\FORMATOS\MATERIAS\M" & Codigo & ".BMP"
        RatonNormal
       .MoveNext
     Loop
 End If
End With
PictLetrasH.Visible = True
PictLetrasV.Visible = True
'MsgBox "Ok"
RatonNormal
End Sub

Private Sub Form_Activate()
  RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoAux1
End Sub
