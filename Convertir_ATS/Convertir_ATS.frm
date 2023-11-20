VERSION 5.00
Begin VB.Form Convertir_ATS 
   Caption         =   "CONVERTIDOR DE ARCHIVOS XML AL NUEVO DIMM ATS"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8760
   Icon            =   "Convertir_ATS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   7665
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   210
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox CmbXML 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3540
      ItemData        =   "Convertir_ATS.frx":1297D
      Left            =   5040
      List            =   "Convertir_ATS.frx":1297F
      Style           =   1  'Simple Combo
      TabIndex        =   5
      Text            =   "CmbXML"
      Top             =   945
      Visible         =   0   'False
      Width           =   5370
   End
   Begin VB.ListBox LstXML 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3420
      Left            =   105
      TabIndex        =   4
      Top             =   945
      Visible         =   0   'False
      Width           =   4845
   End
   Begin VB.CommandButton Command3 
      Caption         =   " &Salir del Convertido"
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
      Left            =   4935
      TabIndex        =   2
      Top             =   105
      Width           =   2325
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Grabar Resultado >>"
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
      Left            =   2520
      TabIndex        =   1
      Top             =   105
      Width           =   2325
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Leer Archivo Anterior >>"
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
      Width           =   2325
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   3
      Top             =   525
      Width           =   7155
   End
End
Attribute VB_Name = "Convertir_ATS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LineaFile As String
Dim CaracFile As String
Dim AuxLinea As String
Dim NombreArchivo As String
Dim RutaGeneraFile As String
Dim I As Long
Dim NumFile As Long
Dim Limpiar As Boolean
Dim EsVenta As Boolean
Dim ContCR As Byte
Dim ContLF As Byte
Dim ContSlay As Byte
Dim ContMayor As Byte
Dim ContIncognita As Byte
Dim TotalVentas As Currency
Dim TotalCompras As Currency

Private Sub Command1_Click()
Dim MaxCar As Integer
Dim PrimerLinea As Long
Dim NI As Byte
Dim NF As Byte
  LstXML.Visible = False
  CmbXML.Visible = False
  Dir_Dialog.Filter = "Todos los archivos|*.xml"
  Dir_Dialog.Filename = Abrir_Archivo(Me.hwnd, Dir_Dialog, OpenFile)
  LstXML.Clear
  CmbXML.Clear
  NombreArchivo = Dir_Dialog.File
  RutaGeneraFile = Dir_Dialog.Filename
  Label1.Caption = RutaGeneraFile
  Screen.MousePointer = vbHourglass
  TotalVentas = 0
  TotalCompras = 0
  PrimerLinea = 0
  ContCR = 0
  ContLF = 0
  ContSlay = 0
  ContMayor = 0
  ContIncognita = 0
  Limpiar = False
  EsVenta = False
  If NombreArchivo <> "" Then
    'MsgBox RutaGeneraFile
     LineaFile = ""
     NumFile = FreeFile
     Open RutaGeneraFile For Input As #NumFile
       Do While Not EOF(NumFile)
          CaracFile = Input(1, #NumFile)
          If CaracFile = vbTab Then CaracFile = ""
          If CaracFile = " " Then CaracFile = ""
          LineaFile = LineaFile & CaracFile
          If CaracFile = "/" Then ContSlay = ContSlay + 1
          If CaracFile = ">" Then ContMayor = ContMayor + 1
          If CaracFile = "?" Then ContIncognita = ContIncognita + 1
          If CaracFile = vbCr Then ContCR = ContCR + 1
          If CaracFile = vbLf Then ContLF = ContLF + 1
          If ContCR = 1 Then Limpiar = True
          If ContLF = 1 Then Limpiar = True
          If ContIncognita = 2 And ContMayor = 1 Then Limpiar = True
          If ContSlay = 1 And ContMayor = 2 Then Limpiar = True
          Select Case LineaFile
            Case "<iva>", "<compras>", "<detalleCompras>", "<air>", "<detalleAir>", "<ventas>", "<detalleVentas>", _
                 "</iva>", "</compras>", "</detalleCompras>", "</air>", "</detalleAir>", "</ventas>", "</detalleVentas>"
                 Limpiar = True
          End Select
          If LineaFile = "<ventas>" Then EsVenta = True
          If LineaFile = "<detalleCompras>" Then TotalCompras = 0
          If Limpiar Then
             If Len(LineaFile) < 5 Then LineaFile = ""
             If Len(LineaFile) > 1 Then
                LstXML.AddItem Trim(LineaFile)
                LineaFile = Replace(LineaFile, "numeroRuc", "IdInformante")
                LineaFile = Replace(LineaFile, "anio", "Anio")
                LineaFile = Replace(LineaFile, "mes", "Mes")
                LineaFile = Replace(LineaFile, "<estabRetencion1>000</estabRetencion1>", "")
                LineaFile = Replace(LineaFile, "<ptoEmiRetencion1>000</ptoEmiRetencion1>", "")
                LineaFile = Replace(LineaFile, "<secRetencion1>0</secRetencion1>", "")
                LineaFile = Replace(LineaFile, "<autRetencion1>000</autRetencion1>", "")
                LineaFile = Replace(LineaFile, "<fechaEmiRet1>00/00/0000</fechaEmiRet1>", "")
                LineaFile = Replace(LineaFile, "<estabRetencion2>000</estabRetencion2>", "")
                LineaFile = Replace(LineaFile, "<ptoEmiRetencion2>000</ptoEmiRetencion2>", "")
                LineaFile = Replace(LineaFile, "<secRetencion2>0</secRetencion2>", "")
                LineaFile = Replace(LineaFile, "<autRetencion2>000</autRetencion2>", "")
                LineaFile = Replace(LineaFile, "<fechaEmiRet2>00/00/0000</fechaEmiRet2>", "")
                LineaFile = Replace(LineaFile, "<docModificado>0</docModificado>", "")
                LineaFile = Replace(LineaFile, "<estabModificado>000</estabModificado>", "")
                LineaFile = Replace(LineaFile, "<ptoEmiModificado>000</ptoEmiModificado>", "")
                LineaFile = Replace(LineaFile, "<secModificado>0</secModificado>", "")
                LineaFile = Replace(LineaFile, "<autModificado>000</autModificado>", "")
                LineaFile = Replace(LineaFile, "<autModificado>000</autModificado>", "")
                For I = 1 To 9
                    LineaFile = Replace(LineaFile, "<tipoComprobante>" & CStr(I) & "</tipoComprobante>", "<tipoComprobante>" & Format(I, "00") & "</tipoComprobante>")
                Next I
                If PrimerLinea = 0 Then
                   CmbXML.AddItem "<?xml version='1.0' encoding='UTF-8'?>"
                ElseIf InStr(LineaFile, "IdInformante") Then
                   CmbXML.AddItem "<TipoIDInformante>R</TipoIDInformante>"
                   CmbXML.AddItem Trim(LineaFile)
                ElseIf InStr(LineaFile, "Mes") Then
                   CmbXML.AddItem Trim(LineaFile)
                   CmbXML.AddItem "<numEstabRuc>001</numEstabRuc>"
                   CmbXML.AddItem "<totalVentas>0.00</totalVentas>"
                   CmbXML.AddItem "<codigoOperativo>IVA</codigoOperativo>"
                Else
                   CmbXML.AddItem Trim(LineaFile)
                End If
               'Sumatoria de la Compra
                If InStr(LineaFile, "baseNoGraIva") And Not EsVenta Then
                   NI = InStr(LineaFile, ">") + 1
                   NF = InStr(LineaFile, "</")
                   TotalCompras = TotalCompras + Val(Mid(LineaFile, NI, NF - NI))
                End If
                If InStr(LineaFile, "baseImponible") And Not EsVenta Then
                   NI = InStr(LineaFile, ">") + 1
                   NF = InStr(LineaFile, "</")
                   TotalCompras = TotalCompras + Val(Mid(LineaFile, NI, NF - NI))
                End If
                If InStr(LineaFile, "baseImpGrav") And Not EsVenta Then
                   NI = InStr(LineaFile, ">") + 1
                   NF = InStr(LineaFile, "</")
                   TotalCompras = TotalCompras + Val(Mid(LineaFile, NI, NF - NI))
                End If
                If InStr(LineaFile, "montoIva") And Not EsVenta Then
                   NI = InStr(LineaFile, ">") + 1
                   NF = InStr(LineaFile, "</")
                   TotalCompras = TotalCompras + Val(Mid(LineaFile, NI, NF - NI))
                End If
                If InStr(LineaFile, "valRetServ100") Then
                   CmbXML.AddItem "<pagoExterior>"
                   CmbXML.AddItem "<pagoLocExt>01</pagoLocExt>"
                   CmbXML.AddItem "<paisEfecPago>NA</paisEfecPago>"
                   CmbXML.AddItem "<aplicConvDobTrib>NA</aplicConvDobTrib>"
                   CmbXML.AddItem "<pagExtSujRetNorLeg>NA</pagExtSujRetNorLeg>"
                   CmbXML.AddItem "</pagoExterior>"
                   If TotalCompras > 1000 Then
                      CmbXML.AddItem "<formasDePago>"
                      CmbXML.AddItem "<formaPago>01</formaPago>"
                      CmbXML.AddItem "</formasDePago>"
                      TotalCompras = 0
                   End If
                End If
               'Sumatoria de las Ventas
                If InStr(LineaFile, "baseNoGraIva") And EsVenta Then
                   NI = InStr(LineaFile, ">") + 1
                   NF = InStr(LineaFile, "</")
                   TotalVentas = TotalVentas + Val(Mid(LineaFile, NI, NF - NI))
                End If
                If InStr(LineaFile, "baseImponible") And EsVenta Then
                   NI = InStr(LineaFile, ">") + 1
                   NF = InStr(LineaFile, "</")
                   TotalVentas = TotalVentas + Val(Mid(LineaFile, NI, NF - NI))
                End If
                If InStr(LineaFile, "baseImpGrav") And EsVenta Then
                   NI = InStr(LineaFile, ">") + 1
                   NF = InStr(LineaFile, "</")
                   TotalVentas = TotalVentas + Val(Mid(LineaFile, NI, NF - NI))
                End If
                If InStr(LineaFile, "montoIva") And EsVenta Then
                   NI = InStr(LineaFile, ">") + 1
                   NF = InStr(LineaFile, "</")
                   TotalVentas = TotalVentas + Val(Mid(LineaFile, NI, NF - NI))
                End If
                If LineaFile = "</ventas>" Then
                   CmbXML.AddItem "<ventasEstablecimiento>"
                   CmbXML.AddItem "<ventaEst>"
                   CmbXML.AddItem "<codEstab>001</codEstab>"
                   CmbXML.AddItem "<ventasEstab>" & Format(TotalVentas, "#0.00") & "</ventasEstab>"
                   CmbXML.AddItem "</ventaEst>"
                   CmbXML.AddItem "</ventasEstablecimiento>"
                End If
             End If
             ContCR = 0
             ContLF = 0
             ContSlay = 0
             ContMayor = 0
             ContIncognita = 0
             Limpiar = False
             LineaFile = ""
             PrimerLinea = PrimerLinea + 1
          End If
       Loop
     Close #NumFile
  Else
     MsgBox "Seleccione un archivo"
  End If
  If InStr(CmbXML.List(8), "totalVentas") Then CmbXML.List(8) = Replace(CmbXML.List(8), "0.00", Format(TotalVentas, "#0.00"))
  LstXML.Width = (Me.ScaleWidth / 2) - 200
  CmbXML.Width = (Me.ScaleWidth / 2) - 200
  CmbXML.Left = LstXML.Width + 200
  LstXML.Height = Me.ScaleHeight - 1500
  CmbXML.Height = Me.ScaleHeight - 1500
  LstXML.Visible = True
  CmbXML.Visible = True
  Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Dim PosI As Integer
Dim PosF As Integer
Dim PosCar As String
    Screen.MousePointer = vbHourglass
    PosF = Len(Label1.Caption)
    PosCar = Mid(Label1.Caption, PosF, 1)
    Do While PosF > 0 And PosCar <> "\"
       PosCar = Mid(Label1.Caption, PosF, 1)
       PosF = PosF - 1
    Loop
    RutaGeneraFile = Mid(Label1.Caption, 1, PosF + 1) & Replace(Mid(Label1.Caption, PosF + 2, Len(Label1.Caption)), "AT", "DISKCOVER_AT")
    NumFile = FreeFile
    Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
    For I = 0 To CmbXML.ListCount - 1
        LineaFile = CmbXML.List(I)
        LineaFile = Replace(LineaFile, vbCr, "")
        LineaFile = Replace(LineaFile, vbLf, "")
        If LineaFile <> "" And Len(LineaFile) > 4 Then Print #NumFile, LineaFile
    Next I
    Close #NumFile
    Screen.MousePointer = vbDefault
    MsgBox "Nombre del Archivo Final" & vbCrLf & vbCrLf & RutaGeneraFile
End Sub

Private Sub Command3_Click()
  End
End Sub

Private Sub Form_Activate()
  Convertir_ATS.WindowState = vbMaximized
End Sub

