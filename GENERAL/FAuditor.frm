VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FAuditoria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MODULO DE AUDITORIA"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "&Consultar"
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
      Left            =   9135
      TabIndex        =   13
      Top             =   105
      Width           =   1170
   End
   Begin MSDataListLib.DataCombo DCModulos 
      Bindings        =   "FAuditor.frx":0000
      DataSource      =   "AdoModulos"
      Height          =   360
      Left            =   1155
      TabIndex        =   10
      Top             =   1155
      Visible         =   0   'False
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCUsuario 
      Bindings        =   "FAuditor.frx":0019
      DataSource      =   "AdoUsuario"
      Height          =   360
      Left            =   1155
      TabIndex        =   9
      Top             =   630
      Visible         =   0   'False
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Modulos"
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
      TabIndex        =   12
      Top             =   1155
      Width           =   1065
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Usuario"
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
      TabIndex        =   11
      Top             =   630
      Width           =   1065
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Eliminar Registro de Auditoria"
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
      Left            =   10395
      TabIndex        =   8
      Top             =   105
      Width           =   1170
   End
   Begin MSAdodcLib.Adodc AdoAuditoria 
      Height          =   330
      Left            =   105
      Top             =   1575
      Width           =   4110
      _ExtentX        =   7250
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
      Caption         =   "Auditoria"
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
      Height          =   855
      Left            =   10395
      TabIndex        =   5
      Top             =   1050
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
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
      Left            =   9135
      TabIndex        =   4
      Top             =   1050
      Width           =   1170
   End
   Begin VB.TextBox TxtAuditoria 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4740
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1995
      Width           =   11460
   End
   Begin VB.ListBox LstAuditoria 
      Height          =   1815
      Left            =   4305
      TabIndex        =   3
      Top             =   105
      Width           =   4740
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   2835
      TabIndex        =   2
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   1470
      TabIndex        =   1
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin MSDataGridLib.DataGrid DGAuditoria 
      Bindings        =   "FAuditor.frx":0032
      Height          =   4740
      Left            =   105
      TabIndex        =   6
      Top             =   1995
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   8361
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoUsuario 
      Height          =   330
      Left            =   420
      Top             =   2205
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Usuario"
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
   Begin MSAdodcLib.Adodc AdoModulos 
      Height          =   330
      Left            =   420
      Top             =   2625
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Modulos"
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Periodo"
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
      Width           =   1380
   End
End
Attribute VB_Name = "FAuditoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Imprimir_Auditoria()
Dim IniX As Single
Dim IniY As Single
Dim Texto As String
Dim LineTexto As String
Dim CampoTexto As String
Dim AnchoDeLinea As Single
On Error GoTo Errorhandler
RatonReloj
SQLMsg1 = UCaseStrg(LstAuditoria.Text)
SQLMsg2 = "Fecha: " & FechaSistema
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
Escala_Centimetro 1, TipoTimes, 8
'Iniciamos la impresion
AnchoDeLinea = 18: Pagina = 1
IniX = 1: IniY = 1
Encabezado_Documento IniX, IniY, 19
LineTexto = TxtAuditoria.Text
Texto = LineTexto
J = Len(Texto)
I = 1: K = 1
LineTexto = ""
Do While I < J
   Caracter = MidStrg(Texto, I, 1)
   CampoTexto = MidStrg(Texto, I, 3)
   LineTexto = LineTexto & Caracter
   If Printer.TextWidth(LineTexto) > AnchoDeLinea Or Asc(Caracter) = 13 Or Asc(Caracter) = 10 Then
      If Printer.TextWidth(LineTexto) > AnchoDeLinea Then
         K = Len(LineTexto)
         If K > 0 Then
            Do
               K = K - 1
               I = I - 1
            Loop Until K < 2 Or MidStrg(LineTexto, K, 1) = " "
            LineTexto = MidStrg(LineTexto, 1, K)
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
Unload Me
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Unload Me
    Exit Sub
Else
    RatonNormal
    Unload Me
End If
End Sub

Private Sub Check1_Click()
  If Check1.value <> 0 Then DCUsuario.Visible = True Else DCUsuario.Visible = False
End Sub

Private Sub Check2_Click()
  If Check2.value <> 0 Then DCModulos.Visible = True Else DCModulos.Visible = False
End Sub

Private Sub Command1_Click()
  MensajeEncabData = "MODULO DE AUDITORIA: " & LstAuditoria.Text
  If Opcion = 1 Then
     ImprimirAdodc AdoAuditoria, 1, 8
  Else
     Imprimir_Auditoria
  End If
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Command3_Click()
  Mensajes = "Eliminar Registros de Auditoria"
  Titulo = "ELIMINACION"
  If BoxMensaje = vbYes Then
     sSQL = "DELETE * " _
          & "FROM Trans_Entrada_Salida " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND ES <> 'H' "
     Ejecutar_SQL_SP sSQL
     Control_Procesos Normal, "Eliminar Registros de Auditoria"
  End If
End Sub

Private Sub Command4_Click()
   DGAuditoria.Visible = False
   TxtAuditoria.Text = ""
   TxtAuditoria.Visible = False
   Codigo1 = DCModulos.Text
   Codigos = DCUsuario.Text
   If Codigos = "" Then Codigos = Ninguno
   FechaValida MBFechaI
   FechaValida MBFechaF
   FechaIni = BuscarFecha(MBFechaI.Text)
   FechaFin = BuscarFecha(MBFechaF.Text)
   Select Case Val(SinEspaciosIzq(LstAuditoria.Text))
    Case 1 '- Cuentas de Audoria por Usuario
          sSQL = "SELECT Nombre_Completo,Usuario,Clave " _
               & "FROM Accesos " _
               & "WHERE MidStrg(Codigo,1,6) <> 'ACCESO' " _
               & "ORDER BY Codigo,Nombre_Completo "
          Select_Adodc_Grid DGAuditoria, AdoAuditoria, sSQL
          DGAuditoria.Visible = True
          Opcion = 1
    Case 2 '- Registro de rutas de auditoria
          sSQL = "SELECT TES.Codigo As Modulos,TES.Fecha,TES.Hora,TES.Tarea,TES.Proceso,A.Nombre_Completo As Nombre_Usuario " _
               & "FROM Trans_Entrada_Salida TES,Accesos As A " _
               & "WHERE TES.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
               & "AND TES.Item = '" & NumEmpresa & "' " _
               & "AND TES.Periodo = '" & Periodo_Contable & "' " _
               & "AND TES.ES <> 'H' "
          If DCModulos.Visible Then sSQL = sSQL & "AND TES.Codigo = '" & Codigo1 & "' "
          If DCUsuario.Visible Then sSQL = sSQL & "AND A.Nombre_Completo = '" & Codigos & "' "
          sSQL = sSQL & "AND TES.CodigoU = A.Codigo " _
               & "ORDER BY TES.Codigo,A.Nombre_Completo,TES.Fecha,TES.Hora "
          Select_Adodc_Grid DGAuditoria, AdoAuditoria, sSQL
          Opcion = 1
          DGAuditoria.Visible = True
    Case 3 '- Anulacion de Transacciones
          RatonReloj
          FechaIni = BuscarFecha(MBFechaI.Text)
          FechaFin = BuscarFecha(MBFechaF.Text)
          Cadena = "Numero de Transacciones Anuladas: " & vbCrLf
          sSQL = "SELECT * " _
               & "FROM Facturas " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
               & "AND T = 'A' " _
               & "AND TC NOT IN ('C','P') " _
               & "ORDER BY TC,Factura "
          Select_Adodc AdoAuditoria, sSQL
          If AdoAuditoria.Recordset.RecordCount > 0 Then
             Contador = 0
             I = AdoAuditoria.Recordset.fields("Factura")
             J = AdoAuditoria.Recordset.fields("Factura")
             TipoDoc = AdoAuditoria.Recordset.fields("TC")
             Do While Not AdoAuditoria.Recordset.EOF
                If TipoDoc <> AdoAuditoria.Recordset.fields("TC") Then
                   Cadena = Cadena & "(" & TipoDoc & ") Factura Desde: " & I & vbCrLf _
                          & "(" & TipoDoc & ") Factura Hasta: " & J & vbCrLf _
                          & "        Cantidad de Transacciones: " & Contador & vbCrLf & vbCrLf
                   Contador = 0
                   TipoDoc = AdoAuditoria.Recordset.fields("TC")
                   I = AdoAuditoria.Recordset.fields("Factura")
                End If
                J = AdoAuditoria.Recordset.fields("Factura")
                Contador = Contador + 1
                AdoAuditoria.Recordset.MoveNext
             Loop
             Cadena = Cadena & "(" & TipoDoc & ") Factura Desde: " & I & vbCrLf _
                    & "(" & TipoDoc & ") Factura Hasta: " & J & vbCrLf _
                    & "        Cantidad de Transacciones: " & Contador & vbCrLf & vbCrLf
          End If
          RatonNormal
          Opcion = 2
          TxtAuditoria.Text = Cadena
          TxtAuditoria.Visible = True
          
    Case 4 '- Registro de Transacciones Erradas
          Mensajes = "Verificar Transacciones Erradas"
          Titulo = "VERIFICACION DE ERRORES"
          If BoxMensaje = vbYes Then Mayorizar_Cuentas_SP
          Opcion = 2
          TxtAuditoria.Visible = True
    Case 5 '- Registros de Transacciones anuladas
          sSQL = "SELECT TES.Fecha,TES.Hora,TES.Tarea,A.Nombre_Completo As Nombre_Usuario,TES.Codigo As Modulos " _
               & "FROM Trans_Entrada_Salida TES,Accesos As A " _
               & "WHERE TES.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
               & "AND TES.Item = '" & NumEmpresa & "' " _
               & "AND TES.Periodo = '" & Periodo_Contable & "' " _
               & "AND TES.ES = 'A' "
          If DCModulos.Visible Then sSQL = sSQL & "AND TES.Codigo = '" & Codigo1 & "' "
          If DCUsuario.Visible Then sSQL = sSQL & "AND A.Nombre_Completo = '" & Codigos & "' "
          sSQL = sSQL & "AND TES.CodigoU = A.Codigo " _
               & "ORDER BY TES.Fecha,TES.Hora,A.Nombre_Completo "
          Select_Adodc_Grid DGAuditoria, AdoAuditoria, sSQL
          Opcion = 1
          DGAuditoria.Visible = True
    Case 6 '- Numero de Transacciones
          FechaIni = BuscarFecha(MBFechaI.Text)
          FechaFin = BuscarFecha(MBFechaF.Text)
          Cadena = "Numero de Transacciones: " & vbCrLf
          sSQL = "SELECT * " _
               & "FROM Facturas " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND T <> 'A' " _
               & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
               & "AND TC NOT IN ('C','P') " _
               & "ORDER BY TC,Factura "
           Select_Adodc AdoAuditoria, sSQL
           If AdoAuditoria.Recordset.RecordCount > 0 Then
              Contador = 1
              I = AdoAuditoria.Recordset.fields("Factura")
              J = AdoAuditoria.Recordset.fields("Factura")
              TipoDoc = AdoAuditoria.Recordset.fields("TC")
              Do While Not AdoAuditoria.Recordset.EOF
                 If TipoDoc <> AdoAuditoria.Recordset.fields("TC") Then
                    Contador = J - I + 1
                    Cadena = Cadena & "(" & TipoDoc & ") Factura Desde: " _
                           & Format$(I, "0000000") & " Hasta: " & Format$(J, "0000000") & vbCrLf _
                           & "   - Cantidad de Transacciones: " & Contador & vbCrLf & vbCrLf
                    TipoDoc = AdoAuditoria.Recordset.fields("TC")
                    I = AdoAuditoria.Recordset.fields("Factura")
                    J = AdoAuditoria.Recordset.fields("Factura")
                 End If
                 J = AdoAuditoria.Recordset.fields("Factura")
                 AdoAuditoria.Recordset.MoveNext
              Loop
              Contador = J - I + 1
              Cadena = Cadena & "(" & TipoDoc & ") Factura Desde: " _
                           & Format$(I, "0000000") & " Hasta: " & Format$(J, "0000000") & vbCrLf _
                           & "   - Cantidad de Transacciones: " & Contador & vbCrLf & vbCrLf
           End If
           sSQL = "SELECT * " _
               & "FROM Comprobantes " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
               & "AND TP IN ('CD','CI','CE') " _
               & "ORDER BY TP DESC,Numero "
           Select_Adodc AdoAuditoria, sSQL
           If AdoAuditoria.Recordset.RecordCount > 0 Then
              I = AdoAuditoria.Recordset.fields("Numero")
              J = AdoAuditoria.Recordset.fields("Numero")
              TipoDoc = AdoAuditoria.Recordset.fields("TP")
              Do While Not AdoAuditoria.Recordset.EOF
                 If TipoDoc <> AdoAuditoria.Recordset.fields("TP") Then
                    Contador = J - I + 1
                    Cadena = Cadena & "(" & TipoDoc & ") Comprobante Desde: " _
                           & Format$(I, "0000000") & " Hasta: " & Format$(J, "0000000") & vbCrLf _
                           & "   - Cantidad de Transacciones: " & Contador & vbCrLf & vbCrLf
                    TipoDoc = AdoAuditoria.Recordset.fields("TP")
                    I = AdoAuditoria.Recordset.fields("Numero")
                    J = AdoAuditoria.Recordset.fields("Numero")
                 End If
                 J = AdoAuditoria.Recordset.fields("Numero")
                 AdoAuditoria.Recordset.MoveNext
              Loop
              Contador = J - I + 1
              Cadena = Cadena & "(" & TipoDoc & ") Comprobante Desde: " _
                           & Format$(I, "0000000") & " Hasta: " & Format$(J, "0000000") & vbCrLf _
                           & "   - Cantidad de Transacciones: " & Contador & vbCrLf & vbCrLf
           End If
           TxtAuditoria.Text = Cadena
           Opcion = 2
           TxtAuditoria.Visible = True
    Case 7 '- Proteccion de Claves
          sSQL = "SELECT Codigo,Usuario,Nombre_Completo " _
               & "FROM Accesos " _
               & "WHERE MidStrg(Codigo,1,6) = 'ACCESO' " _
               & "ORDER BY Codigo,Nombre_Completo "
          Select_Adodc_Grid DGAuditoria, AdoAuditoria, sSQL
          Opcion = 1
          DGAuditoria.Visible = True
    Case 8 '08 - Cuentas de Auditoria
          sSQL = "SELECT T_No,DC,Detalle,Codigo,Item " _
               & "FROM Ctas_Proceso " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "ORDER BY T_No "
          Select_Adodc_Grid DGAuditoria, AdoAuditoria, sSQL
          Opcion = 1
          DGAuditoria.Visible = True
    Case 9 '- Totales
          Opcion = 2
          TxtAuditoria.Visible = True
          RatonReloj
          Control_Procesos Normal, "Cierre Diario de Caja"
          sSQL = "SELECT T_No,DC,Detalle,Codigo,Item " _
               & "FROM Ctas_Proceso " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "ORDER BY T_No "
          Select_Adodc_Grid DGAuditoria, AdoAuditoria, sSQL
          'FCierreCaja.Show
   End Select
End Sub

Private Sub DGAuditoria_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGAuditoria.Visible = False
     GenerarDataTexto FAuditoria, AdoAuditoria
     DGAuditoria.Visible = True
  End If
End Sub

Private Sub Form_Activate()
  LstAuditoria.Clear
  LstAuditoria.AddItem "01 - Cuentas de Audoria por Usuario"
  LstAuditoria.AddItem "02 - Registro de rutas de auditoria"
  LstAuditoria.AddItem "03 - Anulacion de Transacciones"
  LstAuditoria.AddItem "04 - Registro de Transacciones Erradas"
  LstAuditoria.AddItem "05 - Registros de Transacciones anuladas"
  LstAuditoria.AddItem "06 - Numero de Transacciones"
  LstAuditoria.AddItem "07 - Proteccion de Claves"
  LstAuditoria.AddItem "08 - Cuentas de Auditoria"
  LstAuditoria.AddItem "09 - Totales"
  sSQL = "SELECT Nombre_Completo " _
       & "FROM Accesos " _
       & "WHERE Usuario <> '.' " _
       & "GROUP BY Nombre_Completo " _
       & "ORDER BY Nombre_Completo "
  SelectDB_Combo DCUsuario, AdoUsuario, sSQL, "Nombre_Completo"
  
  sSQL = "SELECT Codigo " _
       & "FROM Trans_Entrada_Salida " _
       & "WHERE Codigo <> '.' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "GROUP BY Codigo " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCModulos, AdoModulos, sSQL, "Codigo"
End Sub

Private Sub LstAuditoria_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

Private Sub Form_Load()
  CentrarForm FAuditoria
  ConectarAdodc AdoUsuario
  ConectarAdodc AdoModulos
  ConectarAdodc AdoAuditoria
End Sub

