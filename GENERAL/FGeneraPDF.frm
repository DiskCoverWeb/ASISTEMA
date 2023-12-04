VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSChrt20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Begin VB.Form FGeneraPDF 
   Caption         =   "PDF"
   ClientHeight    =   11115
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   17280
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17280
      _ExtentX        =   30480
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   18
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PDF"
            Object.ToolTipText     =   "Crear PDF"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Acces_UDET"
            Object.ToolTipText     =   "Acceso a Microsoft Access"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SP_Almacenado"
            Object.ToolTipText     =   "Procedimientos almacenas"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Exportar_Excel"
            Object.ToolTipText     =   "Exportar a Excel un grafico"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Estadistica"
            Object.ToolTipText     =   "Estadistica"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Subir_CC"
            Object.ToolTipText     =   "Subir Centro de Costos"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Kardex_Trans"
            Object.ToolTipText     =   "Actualiza Kardex con Transacciones"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Kardex_Facturas"
            Object.ToolTipText     =   "Igualar Facturas con Kardex"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Emails"
            Object.ToolTipText     =   "Envio por mail"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SRI"
            Object.ToolTipText     =   "Leer Documento Autorizado del SRI"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "LeerExcel"
            Object.ToolTipText     =   "Leer Excel en un datagrid"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "LeerXML"
            Object.ToolTipText     =   "Leer archivo XML"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PictureBox"
            Object.ToolTipText     =   "Picture Box"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Navegador"
            Object.ToolTipText     =   "Navegador WEB interno"
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "XML_Firmado"
            Object.ToolTipText     =   "Leer XML Firmado del SRI"
            Object.Tag             =   ""
            ImageIndex      =   19
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Enviar_Json"
            Object.ToolTipText     =   "Prueba de Json post"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
      EndProperty
      MousePointer    =   1
      Begin VB.Frame Frame1 
         Caption         =   "TIPO DE CONTROL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   10710
         TabIndex        =   4
         Top             =   0
         Width           =   7575
         Begin VB.CommandButton Command2 
            Caption         =   "Command1"
            Height          =   330
            Left            =   3360
            TabIndex        =   7
            Top             =   210
            Width           =   1065
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Boton1"
            Height          =   330
            Left            =   2205
            TabIndex        =   6
            Top             =   210
            Width           =   1065
         End
         Begin VB.ComboBox CTipoCtrl 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   105
            TabIndex        =   5
            Top             =   210
            Width           =   2010
         End
         Begin MSAdodcLib.Adodc AdoDataGrid 
            Height          =   330
            Left            =   4515
            Top             =   210
            Width           =   2955
            _ExtentX        =   5212
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
            Caption         =   "DataGrid"
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
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   8400
      Top             =   4725
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1065
      Left            =   105
      ScaleHeight     =   1005
      ScaleWidth      =   6150
      TabIndex        =   8
      Top             =   735
      Visible         =   0   'False
      Width           =   6210
   End
   Begin VB.Timer msgTimer 
      Left            =   6195
      Top             =   840
   End
   Begin MSChart20Lib.MSChart MSChart 
      Height          =   2745
      Left            =   105
      OleObjectBlob   =   "FGeneraPDF.frx":0000
      TabIndex        =   2
      Top             =   5775
      Width           =   5790
   End
   Begin MSDataGridLib.DataGrid DataGrid 
      Bindings        =   "FGeneraPDF.frx":1BAC
      Height          =   2010
      Left            =   210
      TabIndex        =   1
      Top             =   8820
      Visible         =   0   'False
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   3545
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
            LCID            =   12298
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
            LCID            =   12298
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
   Begin MSAdodcLib.Adodc AdoUDET 
      Height          =   330
      Left            =   9240
      Top             =   945
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Caption         =   "UDET"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11760
      Top             =   945
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   435
      Left            =   9030
      Top             =   1365
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   767
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
      Caption         =   "Ctas"
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
   Begin MSAdodcLib.Adodc AdoAsiento 
      Height          =   435
      Left            =   9030
      Top             =   1890
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   767
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
      Caption         =   "Asiento"
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
   Begin VB.Label msgLabel 
      Caption         =   "Label1"
      Height          =   225
      Left            =   7035
      TabIndex        =   3
      Top             =   945
      Width           =   1695
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   11130
      Top             =   945
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   19
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":1BC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":1EE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":2FD92
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":2FFD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":302EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":387BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":38AD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":38DF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":3910A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":39424
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":39716
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":39A30
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":39D4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":3E90C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":3EC26
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":3F878
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":3FB92
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":42BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FGeneraPDF.frx":45C36
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FGeneraPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim obj_Excel As Object
Dim Obj_Libro As Object
Dim Obj_Hoja As Object
Dim cImp As cImpresion

Private Sub CTipoCtrl_Click()
   SiguienteControl
End Sub

Private Sub CTipoCtrl_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub CTipoCtrl_LostFocus()
    MSChart.Visible = False
    DataGrid.Visible = False
'    WebBrowser.Visible = False
    Select Case CTipoCtrl.Text
    Case "PDF"
    Case "Malla"
         DataGrid.Top = 800
         DataGrid.width = MDI_X_Max - 100
         DataGrid.Height = MDI_Y_Max - DataGrid.Top - 50
         DataGrid.Visible = True
    Case "PictureBox"
         Picture1.Top = 800
         Picture1.width = MDI_X_Max - 100
         Picture1.Height = MDI_Y_Max - Picture1.Top - 50
         Picture1.Visible = True
    Case "Estadistica"
         MSChart.Top = 800
         MSChart.width = MDI_X_Max - 100
         MSChart.Height = MDI_Y_Max - MSChart.Top - 50
         MSChart.Visible = True
    Case "Navegador"
'''         With WebBrowser1
'''             .Top = 800
'''             .width = MDI_X_Max - 100
'''             .Height = MDI_Y_Max - WebBrowser1.Top - 50
'''         End With
    Case ""
        
    End Select
End Sub

Private Sub DataGrid_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyP Then
     DataGrid.Visible = False
     GenerarDataTexto FGeneraPDF, AdoDataGrid
     DataGrid.Visible = True
  End If
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoCtas
  ConectarAdodc AdoAsiento
  ConectarAdodc AdoDataGrid
  
  MSChart.Visible = False
  Picture1.Visible = False
  DataGrid.Visible = False
'  WebBrowser.Visible = False
  
  CTipoCtrl.Clear
  CTipoCtrl.AddItem "PDF"
  CTipoCtrl.AddItem "Malla"
  CTipoCtrl.AddItem "PictureBox"
  CTipoCtrl.AddItem "Estadistica"
  CTipoCtrl.AddItem "Navegador"
  CTipoCtrl.Text = "PDF"
    
 'En la carga del formulario, ajustamos los valores y deshabilitamos el Timer
  msgTimer.Enabled = False    ' Timer detenido
  msgTimer.Interval = 5000    ' Pausa 5 segundos
  msgLabel.Caption = ""       ' Mensaje Borrado
  msgLabel.Visible = False    ' Mensaje Oculto
  
  'SRI_Presenta_PDF FGeneraPDF, "C:\SYSBASES\TEMP\Archivo de prueba.pdf"
  
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim cn As New ADODB.Connection
Dim rs As ADODB.Recordset
Dim TipoSRI As Tipo_Contribuyente
Dim SQL_Server1 As Boolean
Dim AdoStrCnn1 As String
Dim RutaPDF As String

Dim v1 As Byte
Dim v2 As Byte
Dim v3 As Integer
Dim v4 As Integer
Dim v5 As Currency
Dim v6 As Currency
Dim v7 As Long
Dim v8 As Long
Dim v9 As Date
Dim v10 As Date

'''Dim Sales(5, 2) As Object
'''
'''Sales = New Object(, ){ _
'''      {"Company", "Company A", "Company B"}, _
'''      {"June", 20, 10}, _
'''      {"July", 10, 5}, _
'''      {"August", 30, 15}, _
'''      {"September", 14, 7}}
'MsgBox Button.key
 RatonReloj
 Select Case Button.key
   Case "Salir"
        RatonNormal
        Unload FGeneraPDF
   Case "PDF"
        With CommonDialog1
             If .Filename = "" Then .Filename = "Seleccione un Archivo"
            .Flags = cdlOFNFileMustExist + cdlOFNNoChangeDir + cdlOFNHideReadOnly
            .Filter = "Archivos PDF|*.pdf"
            .DialogTitle = "Abrir Archivo"
            .Action = 1
             RatonReloj
             
             If .Filename <> "" Then
             
                 'WebBrowser1.Navigate .Filename
             
                 RatonReloj
'''                 fPDF.setZoom 125
'''                 fPDF.setShowScrollbars True
'''                 fPDF.setShowToolbar False
'''                 fPDF.LoadFile .Filename
                 RatonNormal
             End If
            'RoundRect AcroPDF1, 1, 1, 100, 100, 45, 45
             RatonNormal
            'MsgBox "..."
             
        End With
   Case "Imprimir"
        RatonReloj
        Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
        Titulo = "IMPRESION"
        Bandera = False
        'SetPrinters.Show 1
        'If PonImpresoraDefecto(SetNombrePRN) Then
           'Generamos el documento
            tPrint.TipoImpresion = Es_PDF
            tPrint.NombreArchivo = "Archivo de prueba"
            tPrint.TituloArchivo = "PRUEBA DE PDF"
            tPrint.TipoLetra = TipoArial
            tPrint.OrientacionPagina = 1
            tPrint.PaginaA4 = True
            tPrint.EsCampoCorto = False
            tPrint.VerDocumento = True

            Set cPrint = New cImpresion
            cPrint.iniciaImpresion
            cPrint.printImagen RutaSistema & "\LOGOS\DiskCover.gif", 1, 1, 3, 1.5
            cPrint.printCuadro 5, 1, 8, 3, Magenta, "BF"
            cPrint.printCuadro 8.5, 1, 11, 3, Rojo, "B"
            cPrint.printCuadro 5, 5, 10, 5.5, Rojo, "B"
            cPrint.printCuadro 11, 3.5, 15, 4.5, Rojo, "B"
            cPrint.printLinea 8.5, 1, 11, 3, Magenta
            cPrint.printLinea 11.5, 1, 11.5, 3, Azul
            cPrint.printLinea 11.5, 3, 13, 3, Azul
            
            cPrint.generarBarras String(37, "9"), cC128_B, 8.5, 6, 14, 1, True
            cPrint.PorteDeLetra = 14
            cPrint.printTexto 5, 3.5, "20 Hola mundo...", , , Amarillo_Claro
            cPrint.PorteDeLetra = 16
            Cadena = "16 Hola amigos"
            cPrint.printTexto 5, 8, Cadena, , , Amarillo_Claro
            cPrint.printTexto 5, 9, "16 ANCHO CADENA: " & Format(cPrint.anchoTexto(Cadena), "#,##0.00"), , , Amarillo_Claro
            cPrint.PorteDeLetra = 10
            cPrint.colorDeLetra = Rojo
            cPrint.printTexto 5, 10, "10 123112", , 2, Magenta
            cPrint.printTexto 5, 11, "10 1231.34", , 2, Verde_Claro
            cPrint.printTexto 5, 12, "10 182,121,231.34", , 2
            cPrint.PorteDeLetra = 10
            v1 = 200
            v2 = 250
            v3 = 23432
            v4 = 12212
            v5 = 123.56
            v6 = 43454.78
            v7 = 65000
            v8 = 75000
            v9 = "13/12/2014"
            v10 = "20/01/2016"
            cPrint.printVariable 5, 13, v1, 10, , Amarillo_Claro, 2
''            cPrint.printVariables 2, 3.5, v2, 10, , Amarillo, 2
''            cPrint.printVariables 2, 4, v3, 10, , Amarillo, 2
''            cPrint.printVariables 2, 4.5, v4, 10, , vbYellow, 2
''            cPrint.printVariables 2, 5, v5, 10, , Magenta, 2
''            cPrint.printVariables 2, 5.5, v6, 10, , , 2
''            cPrint.printVariables 2, 6, v7, 10, , Amarillo, 2
''            cPrint.printVariables 2, 6.5, v8, 10, , Amarillo, 2
''            cPrint.printVariables 2, 7, v9, 10, , Amarillo, 2
''            cPrint.printVariables 2, 7.5, v10, 10, , Amarillo, 2
            cPrint.finalizaImpresion
            'MsgBox "..."
            'Presentar_PDF fPDF, RutaSysBases & "\TEMP\" & tPrint.NombreArchivo & ".pdf"
        'End If
   Case "Acces_UDET"
        'Cuadro_Impresora
'''        Contador = 0
'''        If Dato_DBF.Tipo_Base = "ACCESS" And Len(Dato_DBF.Actuales) > 1 And Len(Dato_DBF.Carpeta) > 1 Then
'''           sSQL = "SELECT PAGCRE.mat_alu,PAGCRE.cod_sem,PAGCRE.cod_carr,CARR.des_carr,A.nom_alu,A.ape_alu,PAGCRE.comp_no,PAGCRE.efec,PAGCRE.total " _
'''                & "FROM pag_cre AS PAGCRE,carrera AS CARR, alumno AS A " _
'''                & "Where (PAGCRE.mat_alu = a.mat_alu) And (PAGCRE.cod_carr = CARR.cod_carr) "
'''
'''
'''           Set Rs = New ADODB.Recordset
'''           Cn.open "Provider=Microsoft.Jet.OLEDB.4.0;" _
'''                 & "Data Source=" & Dato_DBF.Carpeta & Dato_DBF.Actuales & ".MDB"
'''           Rs.Source = "datosprov"
'''           Rs.CursorType = adOpenKeyset
'''           Rs.LockType = adLockOptimistic
'''           Rs.open sSQL, Cn
'''           Cadena = ""
'''           Do While Not Rs.EOF
'''              Cadena = Cadena & Rs.Fields("mat_alu") & " - " & Rs.Fields("ape_alu") & " - " & Rs.Fields("comp_no") & vbCrLf
'''
'''              Contador = Contador + 1
'''              If Contador > 10 Then Rs.MoveLast
'''              Rs.MoveNext
'''           Loop
'''           Rs.Close
'''           Cn.Close
'''
'''           MsgBox Cadena

'''            SQL_Server1 = SQL_Server
'''            SQL_Server = False
'''            '" & Dato_DBF.Carpeta & Dato_DBF.Actuales & "
'''            AdoStrCnn1 = "Provider=Microsoft.Jet.OLEDB.4.0;" _
'''                       & "Data Source=" & Dato_DBF.Carpeta & Dato_DBF.Actuales & ";" _
'''                       & "USer id=admin; password="
'''            MsgBox AdoStrCnn1
'''            sSQL = "Select A.nom_alu,A.ape_alu,ACCM.des_carr " _
'''                 & "From alumno As A, alumnoxcarreraxcob_mat As ACCM " _
'''                 & "Where a.mat_alu = ACCM.mat_alu "
'''            sSQL = CompilarSQL(sSQL)
'''            AdoUDET.ConnectionString = AdoStrCnn1
'''            AdoUDET.RecordSource = sSQL
'''            AdoUDET.Refresh
'''            AdoUDET.Recordset.Close

'''            SQL_Server = SQL_Server1
'''        End If
   Case "SP_Almacenado"
        'SP_Almacenado
   Case "Exportar_Excel"
        'Exportar_Excel
        sSQL = "SELECT F.CodigoC,C.Actividad,C.Cliente,CI_RUC,C.Direccion,C.Grupo,SUM(Saldo_MN) As Saldo_Pend " _
               & "FROM Facturas As F, Clientes As C " _
               & "WHERE F.Item = '" & NumEmpresa & "' " _
               & "AND F.Periodo = '" & Periodo_Contable & "' " _
               & "AND F.Fecha < #" & BuscarFecha(FechaSistema) & "# " _
               & "AND F.T = 'P' " _
               & "AND NOT F.TC IN ('C','P') " _
               & "AND F.CodigoC = C.Codigo " _
               & "GROUP BY F.CodigoC,C.Actividad,C.Cliente,CI_RUC,C.Direccion,C.Grupo " _
               & "HAVING SUM(Saldo_MN) > 0 " _
               & "ORDER BY F.CodigoC,C.Actividad,C.Cliente,CI_RUC,C.Direccion,C.Grupo "
        Select_Adodc AdoDataGrid, sSQL
        Exportar_AdoDB_Excel AdoDataGrid.Recordset
   Case "Estadistica"
        Estadisticas
   Case "Subir_CC"
        Subir_CC
   Case "Kardex_Facturas"
        Kardex_Facturas
   Case "Kardex_Trans"
        Kardex_Trans
   Case "Emails"
        TMail.ListaMail = 255
        TMail.TipoDeEnvio = "CO"
       'MsgBox RutaBackup
        Cadena = Leer_Archivo_Texto(RutaSistema & "\FORMATOS\credenciales.html")
        TMail.Asunto = "Prueba de Mails por smtp.diskcoversystem.com"
        TMail.MensajeHTML = "" 'Cadena
        TMail.Mensaje = "Esta es una prueba de Correo Electronico enviado por DNS-EXIT, " _
                      & "mensaje enviado desde el PC: " & IP_PC.Nombre_PC & ", a las: " & Time & ", " _
                      & "de la empresa: " & Empresa & "."
        TMail.MensajeHTML = Leer_Archivo_Texto(RutaSistema & "\FONDOS\index1.html")
                      
        TMail.Adjunto = ""
        
        TMail.para = ""
        Insertar_Mail TMail.para, "diskcover.system@yahoo.com"
        Insertar_Mail TMail.para, "diskcover.system@gmail.com"
        Insertar_Mail TMail.para, "informacion@diskcoversystem.com"
        FEnviarCorreos.Show 1
        TMail.para = ""
        TMail.ListaMail = 255
   Case "SRI"
        SRI_Autorizacion = SRI_Leer_XML_Autorizado(RutaSysBases & "\SRI\Comprobantes Recibidos\0110202101179185161700120010050000244741234567814.xml", RutaSysBases & "\SRI\Comprobantes no Autorizados\0110202101179185161700120010050000244741234567814.xml")
        MsgBox SRI_Autorizacion.Estado_SRI & vbCrLf & SRI_Autorizacion.Fecha_Autorizacion & vbCrLf & SRI_Autorizacion.Hora_Autorizacion & vbCrLf & SRI_Autorizacion.Documento_XML
   Case "LeerExcel"
        LeerExcel
   Case "LeerXML"
        LeerXML
   Case "PictureBox"
       'Dimensionar Objeto Picture
        Dim Pic As PictureBox
       'Cargarlo, con una imagen
        'Pic.PaintPicture = LoadPicture(RutaSistema & "\INICIO.jpg")
        'Pic.CurrentX = 10
        'Pic.CurrentY = 10
        'Pic.Print "hola"
       'Asignar la imagen a un Picture para ver que funciona
        Picture1.Picture = Pic
       'Borrar la imagen del Objeto Picture
        Set Pic = Nothing
   Case "Navegador"
       'Codigo_HTML = Inet1.OpenURL("https://srienlinea.sri.gob.ec/sri-catastro-sujeto-servicio-internet/rest/ConsolidadoContribuyente/existePorNumeroRuc?numeroRuc=1793133435001")
       '1793133436001
       '1792164710001
       '0702164179001
       TipoSRI = consulta_RUC_SRI("0702164179001")
       TipoSRI = consulta_RUC_SRI("1792164715001")
       TipoSRI = consulta_RUC_SRI("1793133435001")
       Mensajes = ""
       With TipoSRI
        If Len(.RUC_SRI) > 1 Then Mensajes = Mensajes & "R.U.C.: " & .RUC_SRI & vbCrLf
        If Len(.RazonSocial) > 1 Then Mensajes = Mensajes & "RAZON SOCIAL: " & .RazonSocial & vbCrLf
        If Len(.NombreComercial) > 1 Then Mensajes = Mensajes & "NOMBRE COMERCIAL: " & .NombreComercial & vbCrLf
        If Len(.TipoRUC) > 1 Then Mensajes = Mensajes & UCaseStrg(.TipoRUC) & ", "
        If Len(.Obligado) > 1 Then Mensajes = Mensajes & .Obligado & " OBLIGADO A LLEVAR CONTABILIDAD" & vbCrLf
        If Len(.ActividadEconomica) > 1 Then Mensajes = Mensajes & "ACTIVIDAD ECONOMICA: " & .ActividadEconomica & vbCrLf
        If Len(.FechaInicio) > 1 Then Mensajes = Mensajes & "INICIO SU ACTIVIDAD EL " & .FechaInicio & vbCrLf
        If Len(.FechaActualización) > 1 Then Mensajes = Mensajes & "R.U.C. ACTUALIZADO EL " & .FechaActualización & vbCrLf
        If Len(.FechaReinicio) > 1 Then Mensajes = Mensajes & "REINICIO DE ACTIVIDADES: " & .FechaReinicio & vbCrLf
        If Len(.Categoria) > 1 And Len(.ClaseRUC) > 1 Then Mensajes = Mensajes & "CATEGORIA: " & .Categoria & ", CLASE: " & .ClaseRUC & vbCrLf
        If Len(.FechaCese) > 1 Then Mensajes = Mensajes & "CESE DE ACTIVIDADES: " & .FechaCese & vbCrLf
        If Len(.MicroEmpresa) > 1 Then Mensajes = Mensajes & "TIPO DE CONTRIBUYENTE: """ & UCaseStrg(.MicroEmpresa) & """ " & vbCrLf
        If Len(.AgenteRetencion) > 1 Then Mensajes = Mensajes & "AGENTE DE RETENCION: """ & UCaseStrg(.AgenteRetencion) & """ " & vbCrLf
        If Len(.Estado) > 1 Then Mensajes = Mensajes & "ESTADO DEL CONTRIBUYENTE: """ & UCaseStrg(.Estado) & """ "
       End With
       MsgBox Mensajes
   Case "XML_Firmado"
        Codigo = InputBox("INGRESE LA CLAVE DE ACCESO", "LEER ARCHIVO FIRMADO", "1234567890")
        Recepcion_SRI_XML Codigo
   Case "Enviar_Json"
        MsgBox "1752765675 = " & post_URL_JSon("1752765675", 0, 0)
  End Select
  RatonNormal
End Sub

Public Sub Cuadro_Impresora()
Dim NPRinter, BeginPage, EndPage, NumCopies, Orientation, I
' Establece Cancel a True.
CommonDialog1.CancelError = True
On Error GoTo errHandler
' Presenta el cuadro de diálogo Imprimir.
CommonDialog1.ShowPrinter
' Obtiene los valores seleccionados por el usuario ' en el cuadro de diálogo.
BeginPage = CommonDialog1.FromPage
EndPage = CommonDialog1.ToPage
NumCopies = CommonDialog1.Copies
Orientation = CommonDialog1.Orientation

MsgBox NPRinter & vbCrLf & BeginPage & vbCrLf & EndPage & vbCrLf & NumCopies & vbCrLf & Orientation
For I = 1 To NumCopies
' Escriba aquí código para enviar los datos a la ' impresora.
Next
Exit Sub
errHandler:
' El usuario hizo clic en el botón Cancelar.
Exit Sub
End Sub

Public Sub SP_Almacenado()
Dim AdoReg As ADODB.Recordset
Dim ContFile As Long
Dim NumFile As Long
Dim ArchivoSQL As String
Dim ListaArchivo() As String

'''   Cadena = Dir(RutaSistema & "\BASES\UPDATE_DB\fn_*.sql", vbNormal) 'Recupera la primera entrada.
'''   Do While Cadena <> ""
'''      If Cadena <> "." And Cadena <> ".." Then
'''         ArchivoSQL = MidStrg(Cadena, 1, Len(Cadena) - 4)
'''         sSQL = "IF OBJECT_ID('" & ArchivoSQL & "') IS NOT NULL DROP FUNCTION " & ArchivoSQL & ";"
'''         Ejecutar_SQL_SP sSQL
'''      End If
'''      Cadena = Dir
'''   Loop
'''
'''   Cadena = Dir(RutaSistema & "\BASES\UPDATE_DB\sp_*.sql", vbNormal) 'Recupera la primera entrada.
'''   Do While Cadena <> ""
'''      If Cadena <> "." And Cadena <> ".." Then
'''         ArchivoSQL = MidStrg(Cadena, 1, Len(Cadena) - 4)
'''         sSQL = "IF OBJECT_ID('" & ArchivoSQL & "') IS NOT NULL DROP PROCEDURE " & ArchivoSQL & ";"
'''         Ejecutar_SQL_SP sSQL
'''      End If
'''      Cadena = Dir
'''   Loop
'''   RatonReloj
'''   NumFile = 0
'''   ArchivoSQL = Dir(RutaSistema & "\BASES\UPDATE_DB\*.sql", vbNormal) 'Recupera la primera entrada.
'''   Do While ArchivoSQL <> ""
'''      If ArchivoSQL <> "." And ArchivoSQL <> ".." Then
'''         NumFile = NumFile + 1
'''      End If
'''      ArchivoSQL = Dir
'''   Loop
'''
'''   ReDim ListaArchivo(NumFile) As String
'''   NumFile = 0
'''   ArchivoSQL = Dir(RutaSistema & "\BASES\UPDATE_DB\*.sql", vbNormal) 'Recupera la primera entrada.
'''   Do While ArchivoSQL <> ""
'''      If ArchivoSQL <> "." And ArchivoSQL <> ".." Then
'''         ListaArchivo(NumFile) = ArchivoSQL
'''         NumFile = NumFile + 1
'''      End If
'''      ArchivoSQL = Dir
'''   Loop
'''   For ContFile = 0 To NumFile - 1
'''       sSQL = Leer_Archivo_SQL(RutaSistema & "\BASES\UPDATE_DB\" & ListaArchivo(ContFile))
'''       Ejecutar_SQL_SP sSQL
'''   Next ContFile
   RatonNormal
   Parametros = CBool(OpcCoop) & "," & CBool(ConSucursal) & ",'" & NumEmpresa & "','" & Periodo_Contable & "' "
   Ejecutar_SP "sp_Mayorizar_Cuentas", Parametros
   
   MsgBox "Proceso Terminado"

'Para el Store Procedure los Siguiente si es consulta es de esta manera, ojo en el mismo Modulo :
    'Set DataGrid1.DataSource = EJEMPLO_Listar
    
End Sub

Public Function EJEMPLO_Listar() As ADODB.Recordset
Dim cn As New ADODB.Connection
Dim COMR As ADODB.Command
On Error GoTo Mal
'"Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sA;pWD=sa;Initial Catalog=NOMBREDELABASEDEDATOS"
    RatonReloj
    Set cn = New ADODB.Connection
    cn.ConnectionString = AdoStrCnn
    cn.open
    
    Set COMR = New Command
    With COMR
        .ActiveConnection = cn
        .ActiveConnection.CursorLocation = adUseClient
        .CommandType = adCmdStoredProc
        .CommandText = "sp_Actualiza_Clientes"
        .Parameters.Append COMR.CreateParameter("Item", adVarChar, adParamInput, 3)
        .Parameters("Item").value = NumEmpresa
        .Parameters.Append COMR.CreateParameter("TD", adVarChar, adParamInput, 2)
        .Parameters("TD").value = "C"
        .CommandTimeout = 0
        .Prepared = True
        Set EJEMPLO_Listar = New ADODB.Recordset
        Set EJEMPLO_Listar = .Execute
        Set COMR = Nothing
    End With
    RatonNormal
    MsgBox "Proceso terminado con exito"
Mal:
    Select Case Err.Number
      Case 3021: MsgBox "VerificaTransaccion"
    End Select
End Function

Private Sub Exportar_Excel()

Dim oXL As Object        ' Excel application
Dim oBook As Object      ' Excel workbook
Dim oSheet As Object     ' Excel Worksheet
Dim oChart As Object     ' Excel Chart

Dim iRow As Integer      ' Index variable for the current Row
Dim iCol As Integer      ' Index variable for the current Row

Const cNumCols = 10      ' Number of points in each Series
Const cNumRows = 2       ' Number of Series

ReDim aTemp(1 To cNumRows, 1 To cNumCols)

'Insertar Excel y crear un nuevo libro
Set oXL = CreateObject("Excel.application")
Set oBook = oXL.Workbooks.Add
Set oSheet = oBook.Worksheets.Item(1)

' Insertar datos al azar en las dos series
Randomize Now()
For iRow = 1 To cNumRows
For iCol = 1 To cNumCols
aTemp(iRow, iCol) = Int(Rnd * 50) + 1
Next iCol
Next iRow
oSheet.Range("A1").Resize(cNumRows, cNumCols).value = aTemp

'Añadir gráfico en la primera hoja del libro
Set oChart = oSheet.ChartObjects.Add(150, 140, 300, 200).Chart
oChart.SetSourceData Source:=oSheet.Range("A1").Resize(cNumRows, cNumCols)

' Hacer visible Exccel:
oXL.Visible = True

oXL.UserControl = True

End Sub

Public Sub Estadisticas()
Dim Dato() As Currency

    MSChart.width = 10000
    MSChart.Height = 6000
   'In this example, the sales for two companies are shown for four months.
   'Create the data array and bind it to the ChartData property.
   ''   MSChart.ChartData = Sales
  
    ReDim Dato(10, 2)
     
    MSChart.columnCount = 10 'Numero de columnas
    MSChart.RowCount = 10

    'Add titles to the axes.
    MSChart.Plot.axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).AxisTitle.Text = "Vendedores"
    MSChart.Plot.axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).AxisTitle.Text = "Miles en USD"
    MSChart.chartType = VtChChartType2dBar
    MSChart.RowLabel = "Origenes"
     For I = 0 To 9
        MSChart.Column = I + 1
        MSChart.Row = 1
        MSChart.RowLabel = CStr(I + 1)
        Dato(I, 1) = I * 100
        Dato(I, 2) = I * 200
        MSChart.ChartData = Dato
     Next I
    MSChart.ShowLegend = True
     
         
   'Set custom colors for the bars.
''   With MSChart.Plot
''      'Yellow for Company A
''      ' -1 selects all the datapoints.
''      .SeriesCollection(1).DataPoints(-1).Brush.FillColor.Set(250, 250, 0)
''      'Purple for Company B
''      .SeriesCollection(2).DataPoints(-1).Brush.FillColor.Set(200, 50, 200)
''   End With
        
'''
'''        MSChart.RowCount = 2
'''        For I = 0 To 2 - 1
'''            MSChart.Row = I + 1
'''            MSChart.Data = I
'''            MSChart.RowLabel = "Texto: " & I
'''        Next I
'''
'''
''        MSChart.RowCount = 1
''        MSChart.TitleText = "xxxxx"
''
'''        MSChart.RowLabel = "xxx"
''
''        MSChart.RowCount = 1
''
''        MSChart.ColumnCount = 10
''
''        MSChart.Column = 1
''        MSChart.Data = 1500
''        MSChart.Column = 2
''        MSChart.Data = 3000
''        MSChart.ColumnLabel = "B"
''        MSChart.Column = 3
''        MSChart.Data = 4500
''        MSChart.ColumnLabel = "C"
''        MSChart.Column = 4
''        MSChart.Data = 45500
''        MSChart.Column = 5
''        MSChart.Data = 14500
''        MSChart.Column = 10
''        MSChart.Data = 14500
''
End Sub
 
Private Sub Command1_Click()
  msgLabel = "Mensaje de prueba en el Command Uno"
End Sub
 
Private Sub Command2_Click()
  msgLabel = "Mensaje de prueba en el Command Dos"
End Sub
  
Private Sub msgLabel_Change()
  If msgLabel.Caption = "" Then
    ' Mensaje en blanco, detener temporizador, ocultar mensaje
    msgTimer.Enabled = False
    msgLabel.Visible = False
  Else
    ' Mensaje con datos, activar temporizador, mostrar mensaje
    msgTimer.Enabled = False  'Primero desactivamos para detenerlo
    msgTimer.Enabled = True
    msgLabel.Visible = True
  End If
End Sub
 
Private Sub msgTimer_Timer()
  ' Llegado al tiempo, borramos mensaje
  msgLabel.Caption = ""
End Sub

Public Sub Kardex_Trans()
Dim Cont As Integer
Dim ContIDX As Integer
Dim A_No As Long

   Progreso_Barra.Mensaje_Box = "Consultando el Kardex con Transacciones"
   Progreso_Iniciar
   RatonReloj
   
   NoCheque = Ninguno
   CodigoCC = Ninguno
   Cont = 0
   A_No = 0
   NumComp = 0
   ContIDX = 0
   Trans_No = 197
   
   Progreso_Barra.Mensaje_Box = "Determinando que comprobantes se va ha procesar"
   Progreso_Esperar True
   'MsgBox "..."
   Actualiza_Transacciones_Kardex_SP
   MsgBox "."
   sSQL = "SELECT Fecha,TP, Numero, Codigo_P, Cta_Inv, Contra_Cta, Valor_Total, Cta_Inv, Valor_Total " _
        & "FROM Trans_Kardex " _
        & "WHERE Periodo = '" & Periodo_Contable & "' " _
        & "AND Item  = '" & NumEmpresa & "' " _
        & "AND X = 'M' " _
        & "AND T <> 'A' " _
        & "ORDER BY Fecha, TP, Numero "
   Select_Adodc_Grid DataGrid, AdoDataGrid, sSQL
   With AdoDataGrid.Recordset
    If .RecordCount > 0 Then
        Mensajes = "Quiere Insertar Cta en Comprobante de Kardex"
        Titulo = "PREGUNTA DE ACTUALIZACION"
        If BoxMensaje = vbYes Then
           RatonReloj
           Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
           Do While Not .EOF
              Mifecha = .fields("Fecha")
              CodigoB = .fields("Codigo_P")
              TipoDoc = .fields("TP")
              Numero = .fields("Numero")
              Cta = .fields("Cta_Inv")
              Contra_Cta = .fields("Contra_Cta")
              Progreso_Barra.Mensaje_Box = "[" & ContIDX & "] Ins. Cta. Trans." & TipoDoc & " = " & Numero & ": " & Mifecha & ", " & CodigoCli
              Progreso_Esperar

              SetAdoAddNew "Transacciones"
              SetAdoFields "T", "N"
              SetAdoFields "TP", TipoDoc
              SetAdoFields "Numero", Numero
              SetAdoFields "Fecha", Mifecha
              SetAdoFields "Cta", Contra_Cta
              SetAdoFields "Debe", 1
              SetAdoFields "Codigo_C", CodigoB
              SetAdoUpdate
              
              SetAdoAddNew "Transacciones"
              SetAdoFields "T", "N"
              SetAdoFields "TP", TipoDoc
              SetAdoFields "Numero", Numero
              SetAdoFields "Fecha", Mifecha
              SetAdoFields "Cta", Cta
              SetAdoFields "Haber", 1
              SetAdoFields "Codigo_C", CodigoB
              SetAdoUpdate
              
             .MoveNext
           Loop
        End If
    End If
   End With
   
   Progreso_Barra.Mensaje_Box = "Consultando Comprobantes"
   Progreso_Esperar True
   sSQL = "SELECT C.T, C.Fecha, C.TP, C.Numero, C.Codigo_B, CL.Cliente, C.X " _
        & "FROM Comprobantes As C, Clientes As CL " _
        & "WHERE C.Periodo = '" & Periodo_Contable & "' " _
        & "AND C.Item  = '" & NumEmpresa & "' " _
        & "AND C.X = 'M' " _
        & "AND C.T <> 'A' " _
        & "AND C.Codigo_B = CL.Codigo " _
        & "ORDER BY C.Fecha, C.TP, C.Numero "
   Select_Adodc_Grid DataGrid, AdoDataGrid, sSQL
   
   Mensajes = "Quiere Actualizar Comprobantes en Kardex"
   Titulo = "PREGUNTA DE ACTUALIZACION"
   If BoxMensaje = vbYes Then
      With AdoDataGrid.Recordset
       If .RecordCount > 0 Then
           RatonReloj
           Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
           'MsgBox "...."
           Do While Not .EOF
              Mifecha = .fields("Fecha")
              CodigoB = .fields("Codigo_B")
              TipoDoc = .fields("TP")
              Numero = .fields("Numero")
              CodigoCli = .fields("Cliente")
              Progreso_Barra.Mensaje_Box = "[" & ContIDX & "] " & TipoDoc & " = " & Numero & ": " & Mifecha & ", " & CodigoCli
              Progreso_Esperar
              sSQL = "SELECT Cta_Inv, SUM(Valor_Total) As TotalInv " _
                   & "FROM Trans_Kardex " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND TP = '" & TipoDoc & "' " _
                   & "AND Numero = '" & Numero & "' " _
                   & "GROUP BY Cta_Inv " _
                   & "ORDER BY Cta_Inv "
              Select_Adodc AdoCtas, sSQL
              If AdoCtas.Recordset.RecordCount > 0 Then
                 Do While Not AdoCtas.Recordset.EOF
                    Haber = Redondear(AdoCtas.Recordset.fields("TotalInv"), 2)
                    Cta = AdoCtas.Recordset.fields("Cta_Inv")
                    sSQL = "UPDATE Transacciones " _
                         & "SET Haber = " & Haber & " " _
                         & "WHERE Periodo = '" & Periodo_Contable & "' " _
                         & "AND Item = '" & NumEmpresa & "' " _
                         & "AND TP = '" & TipoDoc & "' " _
                         & "AND Numero = '" & Numero & "' " _
                         & "AND Cta = '" & Cta & "' "
                    Ejecutar_SQL_SP sSQL
                    AdoCtas.Recordset.MoveNext
                 Loop
              End If
              
              sSQL = "SELECT Contra_Cta, SUM(Valor_Total) As TotalInv " _
                   & "FROM Trans_Kardex " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND TP = '" & TipoDoc & "' " _
                   & "AND Numero = '" & Numero & "' " _
                   & "GROUP BY Contra_Cta " _
                   & "ORDER BY Contra_Cta "
              Select_Adodc AdoCtas, sSQL
              If AdoCtas.Recordset.RecordCount > 0 Then
                 Do While Not AdoCtas.Recordset.EOF
                    Debe = Redondear(AdoCtas.Recordset.fields("TotalInv"), 2)
                    Cta = AdoCtas.Recordset.fields("Contra_Cta")
                    
                    sSQL = "UPDATE Transacciones " _
                         & "SET Debe = " & Debe & " " _
                         & "WHERE Periodo = '" & Periodo_Contable & "' " _
                         & "AND Item = '" & NumEmpresa & "' " _
                         & "AND TP = '" & TipoDoc & "' " _
                         & "AND Numero = '" & Numero & "' " _
                         & "AND Cta = '" & Cta & "' "
                    Ejecutar_SQL_SP sSQL
                    
                    sSQL = "DELETE * " _
                         & "FROM Trans_SubCtas " _
                         & "WHERE Item = '" & NumEmpresa & "' " _
                         & "AND Periodo = '" & Periodo_Contable & "' " _
                         & "AND TP = '" & TipoDoc & "' " _
                         & "AND Numero = '" & Numero & "' " _
                         & "AND Cta = '" & Cta & "' " _
                         & "AND Debitos > 0 "
                    Ejecutar_SQL_SP sSQL
                    AdoCtas.Recordset.MoveNext
                 Loop
              End If
              
              sSQL = "SELECT TK.Contra_Cta, TK.CodigoL, TK.Valor_Total, CSC.TC " _
                   & "FROM Trans_Kardex As TK, Catalogo_SubCtas As CSC " _
                   & "WHERE TK.Item = '" & NumEmpresa & "' " _
                   & "AND TK.Periodo = '" & Periodo_Contable & "' " _
                   & "AND TK.TP = '" & TipoDoc & "' " _
                   & "AND TK.Numero = '" & Numero & "' " _
                   & "AND TK.Periodo = CSC.Periodo " _
                   & "AND TK.Item = CSC.Item " _
                   & "AND TK.CodigoL = CSC.Codigo " _
                   & "ORDER BY TK.TP, TK.Numero, TK.Contra_Cta "
              Select_Adodc AdoCtas, sSQL
              If AdoCtas.Recordset.RecordCount > 0 Then
                 Do While Not AdoCtas.Recordset.EOF
                    Debe = Redondear(AdoCtas.Recordset.fields("Valor_Total"), 2)
                    Cta = AdoCtas.Recordset.fields("Contra_Cta")
                    SubCta = AdoCtas.Recordset.fields("TC")
                    CodigoL = AdoCtas.Recordset.fields("CodigoL")
                    If Len(CodigoL) > 1 Then
                       SetAdoAddNew "Trans_SubCtas"
                       SetAdoFields "T", "N"
                       SetAdoFields "TC", SubCta
                       SetAdoFields "Cta", Cta
                       SetAdoFields "TP", TipoDoc
                       SetAdoFields "Numero", Numero
                       SetAdoFields "Codigo", CodigoL
                       SetAdoFields "Debitos", Debe
                       SetAdoFields "Fecha", Mifecha
                       SetAdoFields "Fecha_V", Mifecha
                       SetAdoFields "Detalle_SubCta", CodigoCli
                       SetAdoUpdate
                    End If
                    AdoCtas.Recordset.MoveNext
                 Loop
              End If
              If ContIDX >= 400 Then
                 RatonNormal
                 Mensajes = "Desea seguir procesando"
                 Titulo = "PREGUNTA DE CONFIRMACION"
                 If BoxMensaje = vbYes Then ContIDX = 0 Else GoTo Termino
              End If
              ContIDX = ContIDX + 1
             .MoveNext
           Loop
       End If
      End With
   End If
Termino:
   RatonNormal
   Progreso_Final
   DataGrid.Visible = True
   MsgBox "Proceso Terminado"
End Sub

Public Sub Kardex_Facturas()
Dim Cont As Integer
Dim ContIDX As Integer
Dim A_No As Long

   Progreso_Barra.Mensaje_Box = "Consultando el Kardex con Transacciones"
   Progreso_Iniciar
   RatonReloj
   
   Parametros = "'" & NumEmpresa & "','" & Periodo_Contable & "' "
   Ejecutar_SP "sp_Reindexar_Periodo", Parametros
   Mayorizar_Cuentas_SP
   Mayorizar_Inventario_SP
   
   NoCheque = Ninguno
   CodigoCC = Ninguno
   Cont = 0
   A_No = 0
   NumComp = 0
   ContIDX = 0
   Trans_No = 197
   
   Progreso_Barra.Mensaje_Box = "Determinando que comprobantes se va ha procesar"
   Progreso_Esperar True
   
   sSQL = "UPDATE Detalle_Factura " _
        & "SET Corte = 0 " _
        & "WHERE Periodo = '" & Periodo_Contable & "' " _
        & "AND Item = '" & NumEmpresa & "' "
   Ejecutar_SQL_SP sSQL
   
   sSQL = "UPDATE Trans_Kardex " _
        & "SET X = '.' " _
        & "WHERE Periodo = '" & Periodo_Contable & "' " _
        & "AND Item = '" & NumEmpresa & "' "
   Ejecutar_SQL_SP sSQL
   
   sSQL = "UPDATE Detalle_Factura " _
        & "SET Corte = (SELECT ROUND(SUM(TK.Salida),2,0) " _
        & "             FROM Trans_Kardex As TK " _
        & "             WHERE TK.Item = '" & NumEmpresa & "' " _
        & "             AND TK.Periodo = '" & Periodo_Contable & "' " _
        & "             AND TK.Salida <> 0 " _
        & "             AND Detalle_Factura.Item = TK.Item " _
        & "             AND Detalle_Factura.Periodo = TK.Periodo " _
        & "             AND Detalle_Factura.Fecha = TK.Fecha " _
        & "             AND Detalle_Factura.TC = TK.TC " _
        & "             AND Detalle_Factura.Serie = TK.Serie " _
        & "             AND Detalle_Factura.Factura = TK.Factura " _
        & "             AND Detalle_Factura.Codigo = TK.Codigo_Inv) " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND T <> 'A' "
   Ejecutar_SQL_SP sSQL
   
   sSQL = "UPDATE Detalle_Factura " _
        & "SET Corte = 0 " _
        & "WHERE Periodo = '" & Periodo_Contable & "' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Corte IS NULL "
   Ejecutar_SQL_SP sSQL
   
   sSQL = "SELECT C.Cliente, DF.* " _
        & "FROM Detalle_Factura As DF, Catalogo_Productos As CP, Clientes As C " _
        & "WHERE DF.Periodo = '" & Periodo_Contable & "' " _
        & "AND DF.Item  = '" & NumEmpresa & "' " _
        & "AND (DF.Corte - DF.Cantidad) <> 0.0 " _
        & "AND DF.Corte = 0 " _
        & "AND DF.T <> 'A' " _
        & "AND DF.Item = CP.Item " _
        & "AND DF.Periodo = CP.Periodo " _
        & "AND DF.Codigo = CP.Codigo_Inv " _
        & "AND DF.CodigoC = C.Codigo " _
        & "ORDER BY DF.Corte, DF.Fecha, DF.TC, DF.Serie, DF.Factura, DF.Codigo "
   Select_Adodc_Grid DataGrid, AdoDataGrid, sSQL
   With AdoDataGrid.Recordset
    If .RecordCount > 0 Then
        Mensajes = "Quiere actualizar el detalle con el Kardex"
        Titulo = "PREGUNTA DE ACTUALIZACION"
        If BoxMensaje = vbYes Then
           RatonReloj
           Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
           Do While Not .EOF
              Mifecha = .fields("Fecha")
              TipoDoc = .fields("TC")
              Factura_No = .fields("Factura")
              CodigoInv = .fields("Codigo")
              Cod_Bodega = .fields("CodBodega")
              Cod_Marca = .fields("CodMarca")
              Progreso_Barra.Mensaje_Box = "[" & ContIDX & "] Ins. Cta. Trans." & TipoDoc & " = " & Factura_No & ": " & Mifecha & ", " & CodigoInv
              Progreso_Esperar

             'Grabamos en el Kardex la factura
              'MsgBox CodigoInv
              If Leer_Codigo_Inv(CodigoInv, Mifecha, Cod_Bodega, Cod_Marca) Then
                 'MsgBox DatInv.Costo
                 If DatInv.Costo > 0 And Len(DatInv.Cta_Inventario) > 1 And Len(DatInv.Cta_Costo_Venta) > 1 Then
                    SetAdoAddNew "Trans_Kardex"
                    SetAdoFields "T", "@" ' Normal
                    SetAdoFields "TC", .fields("TC")
                    SetAdoFields "Serie", .fields("Serie")
                    SetAdoFields "Fecha", .fields("Fecha")
                    SetAdoFields "Factura", .fields("Factura")
                    SetAdoFields "Codigo_P", .fields("CodigoC")
                    SetAdoFields "CodBodega", .fields("CodBodega")
                    SetAdoFields "CodMarca", .fields("CodMarca")
                    SetAdoFields "Codigo_Inv", .fields("Codigo")
                    SetAdoFields "CodigoL", .fields("CodigoL")
                    SetAdoFields "Lote_No", .fields("Lote_No")
                    SetAdoFields "Fecha_Fab", .fields("Fecha_Fab")
                    SetAdoFields "Fecha_Exp", .fields("Fecha_Exp")
                    SetAdoFields "Procedencia", .fields("Procedencia")
                    SetAdoFields "Modelo", .fields("Modelo")
                    SetAdoFields "Orden_No", .fields("Orden_No")
                    SetAdoFields "Serie_No", .fields("Serie_No")
                    SetAdoFields "Total_IVA", .fields("Total_IVA")
                    SetAdoFields "Porc_C", .fields("Porc_C")
                    SetAdoFields "Salida", .fields("Cantidad")
                    SetAdoFields "PVP", .fields("Precio")
                    SetAdoFields "Valor_Unitario", .fields("Precio")
                    SetAdoFields "Costo", DatInv.Costo
                    SetAdoFields "Valor_Total", Redondear(.fields("Cantidad") * .fields("Precio"), 2)
                    SetAdoFields "Total", Redondear(.fields("Cantidad") * DatInv.Costo, 2)
                    SetAdoFields "Detalle", "FA: " + MidStrg(.fields("Cliente"), 1, 96)
                    SetAdoFields "Codigo_Barra", DatInv.Codigo_Barra
                    SetAdoFields "Cta_Inv", DatInv.Cta_Inventario
                    SetAdoFields "Contra_Cta", DatInv.Cta_Costo_Venta
                    SetAdoFields "Item", NumEmpresa
                    SetAdoFields "Periodo", Periodo_Contable
                    SetAdoFields "CodigoU", CodigoUsuario
                    SetAdoUpdate
                 End If
              End If
              Cont = Cont + 1
              If Cont > 200 Then
                 Mensajes = "Quiere seguir actualizando el detalle con el Kardex"
                 Titulo = "PREGUNTA DE ACTUALIZACION"
                 If BoxMensaje = vbNo Then GoTo Termino
                 Cont = 0
              End If
             .MoveNext
           Loop
        End If
    End If
   End With

'''   Cont = 0
'''   sSQL = "SELECT C.Cliente, DF.TC, DF.Serie, DF.Factura, TK.Codigo_Inv, DF.Producto, MAX(TK.ID) AS Borrar " _
'''        & "FROM Detalle_Factura As DF, Trans_Kardex As TK, Clientes As C " _
'''        & "WHERE DF.Periodo = '" & Periodo_Contable & "' " _
'''        & "AND DF.Item  = '" & NumEmpresa & "' " _
'''        & "AND (DF.Corte - DF.Cantidad) <> 0.0 " _
'''        & "AND DF.Corte <> 0 " _
'''        & "AND DF.T <> 'A' " _
'''        & "AND DF.Item = TK.Item " _
'''        & "AND DF.Periodo = TK.Periodo " _
'''        & "AND DF.Codigo = TK.Codigo_Inv " _
'''        & "AND DF.TC = TK.TC " _
'''        & "AND DF.Serie = TK.Serie " _
'''        & "AND DF.Factura = TK.Factura " _
'''        & "AND DF.CodigoC = C.Codigo " _
'''        & "GROUP BY C.Cliente, DF.TC, DF.Serie, DF.Factura, TK.Codigo_Inv, DF.Producto " _
'''        & "ORDER BY C.Cliente, DF.TC, DF.Serie, DF.Factura, TK.Codigo_Inv, DF.Producto "
'''   Select_Adodc_Grid DataGrid, AdoDataGrid, sSQL
'''   With AdoDataGrid.Recordset
'''    If .RecordCount > 0 Then
'''        Mensajes = "Quiere eliminar el detalle duplicado del Kardex"
'''        Titulo = "PREGUNTA DE ELIMINACION"
'''        If BoxMensaje = vbYes Then
'''           RatonReloj
'''           Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
'''           Do While Not .EOF
'''              TipoDoc = .Fields("TC")
'''              Factura_No = .Fields("Factura")
'''              CodigoInv = .Fields("Codigo_Inv")
'''              Progreso_Barra.Mensaje_Box = "[" & ContIDX & "] Ins. Cta. Trans." & TipoDoc & " = " & Factura_No & ", " & CodigoInv
'''              Progreso_Esperar
'''
'''             'Grabamos en el Kardex la factura
'''              sSQL = "DELETE * " _
'''                   & "FROM Trans_Kardex " _
'''                   & "WHERE Periodo = '" & Periodo_Contable & "' " _
'''                   & "AND Item = '" & NumEmpresa & "' " _
'''                   & "AND ID = " & .Fields("Borrar") & " "
'''              Ejecutar_SQL_SP sSQL
'''              Cont = Cont + 1
'''              If Cont > 100 Then
'''                 Mensajes = "Quiere seguir actualizando el detalle con el Kardex"
'''                 Titulo = "PREGUNTA DE ACTUALIZACION"
'''                 If BoxMensaje = vbNo Then GoTo Termino
'''                 Cont = 0
'''              End If
'''             .MoveNext
'''           Loop
'''        End If
'''    End If
'''   End With

'''   Parametros = "'" & NumEmpresa & "','" & Periodo_Contable & "' "
'''   Ejecutar_SP "sp_Reindexar_Periodo", Parametros
'''   Mayorizar_Cuentas_SP
'''   Mayorizar_Inventario_SP
Termino:
   RatonNormal
   Progreso_Final
   DataGrid.Visible = True
   MsgBox "Proceso Terminado"
End Sub

Public Sub Subir_CC()
Dim Cont As Integer
Dim ContIDX As Integer
Dim A_No As Long
   Progreso_Barra.Mensaje_Box = "Consultando el Balance"
   Progreso_Iniciar
   RatonReloj
   
   sSQL = "UPDATE Trans_Kardex " _
        & "SET Procesado = 0 " _
        & "WHERE Periodo = '" & Periodo_Contable & "' " _
        & "AND Item = '" & NumEmpresa & "' "
   Ejecutar_SQL_SP sSQL
   
   NoCheque = Ninguno
   CodigoCC = Ninguno
   Cont = 0
   A_No = 0
   NumComp = 0
   ContIDX = 0
   Trans_No = 197
   IniciarAsientosAdo AdoAsiento
      
   Mifecha = BuscarFecha(FechaSistema)
   sSQL = "SELECT T, IDX, FECHA, CODIGO_INVENTARIO, PRODUCTO, CANTIDAD_SALIDA, CODIGO_CC, CENTRO_DE_COSTOS, CATEGORIA, CTA_INVENTARIO, CODIGO_CONTABLE, DESCRIPCION_CUENTA, " _
        & "CEDULA, APELLIDOS_NOMBRES, FICHA_CLINICA, CI_RUC, CODIGO " _
        & "FROM A_Salida_Centro_Costos " _
        & "WHERE T = 'N' " _
        & "ORDER BY IDX, FECHA, CI_RUC, APELLIDOS_NOMBRES, CODIGO_CONTABLE, CTA_INVENTARIO, CODIGO_CC, CODIGO_INVENTARIO "
   Select_Adodc_Grid DataGrid, AdoDataGrid, sSQL
   RatonNormal
   Mensajes = "Quiere procesar Comprobantes automaticos"
   Titulo = "PREGUNTA DE ACTUALIZACION"
   If BoxMensaje = vbYes Then
      DataGrid.Visible = False
      Mayorizar_Inventario_SP
      With AdoDataGrid.Recordset
       If .RecordCount > 0 Then
           RatonReloj
           Progreso_Barra.Valor_Maximo = .RecordCount
           Cont = .fields("IDX")
           Mifecha = .fields("FECHA")
           CodigoA = .fields("CI_RUC")
           CodigoB = TrimStrg(.fields("APELLIDOS_NOMBRES"))
           CodigoC = TrimStrg(.fields("FICHA_CLINICA"))
           Codigo1 = TrimStrg(.fields("CATEGORIA"))
           CodigoCli = .fields("CODIGO")
           
           sSQL = "UPDATE Trans_Kardex " _
                & "SET X = '.' " _
                & "WHERE Periodo = '" & Periodo_Contable & "' " _
                & "AND Item = '" & NumEmpresa & "' "
           Ejecutar_SQL_SP sSQL

           sSQL = "UPDATE Trans_Kardex " _
                & "SET X = 'D' " _
                & "FROM Trans_Kardex AS TK, A_Salida_Centro_Costos As CC " _
                & "WHERE TK.Periodo = '" & Periodo_Contable & "' " _
                & "AND TK.Item = '" & NumEmpresa & "' " _
                & "AND TK.Entrada = 0 " _
                & "AND TK.Salida > 0 " _
                & "AND CC.T = 'N' " _
                & "AND TK.Codigo_Inv = CC.CODIGO_INVENTARIO " _
                & "AND TK.Fecha = CC.FECHA " _
                & "AND TK.Codigo_P = CC.CODIGO " _
                & "AND TK.CodigoL = CC.CODIGO_CC " _
                & "AND TK.Contra_Cta = CC.CODIGO_CONTABLE " _
                & "AND TK.Cta_Inv = CC.CTA_INVENTARIO "
           Ejecutar_SQL_SP sSQL
           
           sSQL = "DELETE * " _
                & "FROM Trans_Kardex " _
                & "WHERE Periodo = '" & Periodo_Contable & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND T = 'D' "
           Ejecutar_SQL_SP sSQL
           
           Do While Not .EOF
              Progreso_Barra.Mensaje_Box = "[" & ContIDX & "] CD = " & NumComp & ": " & Mifecha & ", " & CodigoB & " -> (" & A_No & ") " & .fields("CODIGO_INVENTARIO")
              Progreso_Esperar
              If Cont <> .fields("IDX") Then
                 Ln_No = 0
                 sSQL = "SELECT CONTRA_CTA, SUM(VALOR_TOTAL) As TotCta " _
                      & "FROM Asiento_K " _
                      & "WHERE Item = '" & NumEmpresa & "' " _
                      & "AND CodigoU = '" & CodigoUsuario & "' " _
                      & "AND T_No = " & Trans_No & " " _
                      & "GROUP BY CONTRA_CTA " _
                      & "ORDER BY CONTRA_CTA "
                 Select_Adodc AdoCtas, sSQL
                 If AdoCtas.Recordset.RecordCount > 0 Then
                    Do While Not AdoCtas.Recordset.EOF
                       InsertarAsientos AdoAsiento, AdoCtas.Recordset.fields("CONTRA_CTA"), 0, AdoCtas.Recordset.fields("TotCta"), 0
                       Ln_No = Ln_No + 1
                       AdoCtas.Recordset.MoveNext
                    Loop
                 End If
              
                 sSQL = "SELECT CTA_INVENTARIO, SUM(VALOR_TOTAL) As TotCta " _
                      & "FROM Asiento_K " _
                      & "WHERE Item = '" & NumEmpresa & "' " _
                      & "AND CodigoU = '" & CodigoUsuario & "' " _
                      & "AND T_No = " & Trans_No & " " _
                      & "GROUP BY CTA_INVENTARIO " _
                      & "ORDER BY CTA_INVENTARIO "
                 Select_Adodc AdoCtas, sSQL
                 If AdoCtas.Recordset.RecordCount > 0 Then
                    Do While Not AdoCtas.Recordset.EOF
                       InsertarAsientos AdoAsiento, AdoCtas.Recordset.fields("CTA_INVENTARIO"), 0, 0, AdoCtas.Recordset.fields("TotCta")
                       Ln_No = Ln_No + 1
                       AdoCtas.Recordset.MoveNext
                    Loop
                 End If
                 RatonReloj
                 Factura_No = 0
                 FechaTexto = Mifecha
                 FechaComp = Mifecha
                 NumComp = ReadSetDataNum("Diario", True, True)
                 Co.T = Normal
                 Co.TP = "CD"
                 Co.Numero = NumComp
                 Co.Fecha = FechaTexto
                 Co.Concepto = "Salida por Farmacia de: " & CodigoB & ", centro de costos "
                 Co.CodigoB = CodigoCli
                 Co.Efectivo = 0
                 Co.Monto_Total = 0
                 Co.Usuario = CodigoUsuario
                 Co.T_No = Trans_No
                 Co.Item = NumEmpresa
                 
                 If Len(CodigoC) > 1 Then Co.Concepto = Co.Concepto & ", Ficha Clinica: " & CodigoC
                 If Len(Codigo1) > 1 Then Co.Concepto = Co.Concepto & ", Categoria: " & Codigo1
                 GrabarComprobante Co
                 Mayorizar_Inventario_SP
              
                 sSQL = "UPDATE A_Salida_Centro_Costos " _
                      & "SET T = 'P' " _
                      & "WHERE IDX = " & Cont & " "
                 Ejecutar_SQL_SP sSQL
                'MsgBox sSQL
                 Cont = .fields("IDX")
                 ContIDX = ContIDX + 1
                 Mifecha = .fields("FECHA")
                 CodigoA = .fields("CI_RUC")
                 CodigoB = TrimStrg(.fields("APELLIDOS_NOMBRES"))
                 CodigoC = TrimStrg(.fields("FICHA_CLINICA"))
                 Codigo1 = TrimStrg(.fields("CATEGORIA"))
                 CodigoCli = .fields("CODIGO")
                'MsgBox "CD = " & NumComp & " .."
                 IniciarAsientosAdo AdoAsiento
                 A_No = 0
              End If
              
              If ContIDX >= 300 Then
                 RatonNormal
                 Mensajes = "Desea seguir procesando"
                 Titulo = "PREGUNTA DE CONFIRMACION"
                 If BoxMensaje = vbYes Then ContIDX = 0 Else GoTo Termino
              End If

             'Averiguamos el costo promedio de salida
              Stock_Actual_Inventario .fields("FECHA"), .fields("CODIGO_INVENTARIO"), "01"
              ValorTotal = Redondear(ValorUnit * .fields("CANTIDAD_SALIDA"), 2)
              RatonReloj
              
              '.Fields ("FICHA_CLINICA")
              SetAdoAddNew "Asiento_K"
              SetAdoFields "DH", "2"
              SetAdoFields "CODIGO_INV", .fields("CODIGO_INVENTARIO")
              SetAdoFields "PRODUCTO", .fields("PRODUCTO")
              SetAdoFields "CANT_ES", .fields("CANTIDAD_SALIDA")
              SetAdoFields "VALOR_UNIT", ValorUnit
              SetAdoFields "VALOR_TOTAL", ValorTotal
              SetAdoFields "CTA_INVENTARIO", .fields("CTA_INVENTARIO")
              SetAdoFields "CONTRA_CTA", .fields("CODIGO_CONTABLE")
              SetAdoFields "CANTIDAD", .fields("CANTIDAD_SALIDA")
              SetAdoFields "SUBCTA", .fields("CODIGO_CC")
              SetAdoFields "UNIDAD", "UNIDAD"
              SetAdoFields "Codigo_B", CodigoCli
              SetAdoFields "CodBod", "01"
              SetAdoFields "Item", NumEmpresa
              SetAdoFields "CodigoU", CodigoUsuario
              SetAdoFields "T_No", Trans_No
              SetAdoFields "A_No", A_No
              SetAdoUpdate
              
              SetAdoAddNew "Asiento_SC"
              SetAdoFields "TM", "1"
              SetAdoFields "DH", "1"
              SetAdoFields "Factura", 0
              SetAdoFields "Codigo", .fields("CODIGO_CC")
              SetAdoFields "FECHA_V", .fields("FECHA")
              SetAdoFields "Cta", .fields("CODIGO_CONTABLE")
              SetAdoFields "Detalle_SubCta", CodigoB
              SetAdoFields "TC", "CC"
              SetAdoFields "T_No", Trans_No
              SetAdoFields "SC_No", A_No
              SetAdoFields "Valor", ValorTotal
              SetAdoUpdate
                            
              A_No = A_No + 1
             .MoveNext
           Loop
           Ln_No = 0
           sSQL = "SELECT CONTRA_CTA, SUM(VALOR_TOTAL) As TotCta " _
                & "FROM Asiento_K " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND CodigoU = '" & CodigoUsuario & "' " _
                & "AND T_No = " & Trans_No & " " _
                & "GROUP BY CONTRA_CTA " _
                & "ORDER BY CONTRA_CTA "
           Select_Adodc AdoCtas, sSQL
           If AdoCtas.Recordset.RecordCount > 0 Then
              Do While Not AdoCtas.Recordset.EOF
                 InsertarAsientos AdoAsiento, AdoCtas.Recordset.fields("CONTRA_CTA"), 0, AdoCtas.Recordset.fields("TotCta"), 0
                 Ln_No = Ln_No + 1
                 AdoCtas.Recordset.MoveNext
              Loop
           End If
        
           sSQL = "SELECT CTA_INVENTARIO, SUM(VALOR_TOTAL) As TotCta " _
                & "FROM Asiento_K " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND CodigoU = '" & CodigoUsuario & "' " _
                & "AND T_No = " & Trans_No & " " _
                & "GROUP BY CTA_INVENTARIO " _
                & "ORDER BY CTA_INVENTARIO "
           Select_Adodc AdoCtas, sSQL
           If AdoCtas.Recordset.RecordCount > 0 Then
              Do While Not AdoCtas.Recordset.EOF
                 InsertarAsientos AdoAsiento, AdoCtas.Recordset.fields("CTA_INVENTARIO"), 0, 0, AdoCtas.Recordset.fields("TotCta")
                 Ln_No = Ln_No + 1
                 AdoCtas.Recordset.MoveNext
              Loop
           End If
           RatonReloj
           Factura_No = 0
           FechaTexto = Mifecha
           FechaComp = Mifecha
           NumComp = ReadSetDataNum("Diario", True, True)
           Co.T = Normal
           Co.TP = "CD"
           Co.Numero = NumComp
           Co.Fecha = FechaTexto
           Co.Concepto = "Salida por Farmacia de: " & CodigoB & ", centro de costos "
           Co.CodigoB = CodigoCli
           Co.Efectivo = 0
           Co.Monto_Total = 0
           Co.Usuario = CodigoUsuario
           Co.T_No = Trans_No
           Co.Item = NumEmpresa
           
           If Len(CodigoC) > 1 Then Co.Concepto = Co.Concepto & ", Ficha Clinica: " & CodigoC
           If Len(Codigo1) > 1 Then Co.Concepto = Co.Concepto & ", Categoria: " & Codigo1
           GrabarComprobante Co
           Mayorizar_Inventario_SP
        
           sSQL = "UPDATE A_Salida_Centro_Costos " _
                & "SET T = 'P' " _
                & "WHERE IDX = " & Cont & " "
           Ejecutar_SQL_SP sSQL
       End If
      End With
   End If
'   Eliminar_Nulos_SP "Clientes"
Termino:
   RatonNormal
   Progreso_Final
   DataGrid.Visible = True
   MsgBox "Proceso Terminado"
End Sub

Public Sub LeerExcel()
Dim rsExcel As ADODB.Recordset

  Set rsExcel = New ADODB.Recordset
 
      Dim sFileName As String
      Dim sfilter As String
      
      sfilter = "*.xls|*.xlsx|*.csv|"
      CommonDialog1.Filter = sfilter
      CommonDialog1.ShowOpen
 
      If Trim(CommonDialog1.Filename) <> "" Then
 
         sFileName = CommonDialog1.Filename
 
         'Set rsExcel = Importar_Excel_AdoDB(Inet1, sFileName)
 
         'Set DataGrid.DataSource = rsExcel
         Set AdoDataGrid.Recordset = rsExcel
         
       End If
       If AdoDataGrid.Recordset.RecordCount > 0 Then
       AdoDataGrid.Recordset.MoveLast
       DataGrid.Caption = AdoDataGrid.Recordset.RecordCount
       Cadena = "Registros: " & AdoDataGrid.Recordset.RecordCount & vbCrLf
       For I = 0 To AdoDataGrid.Recordset.fields.Count - 1
           Cadena = Cadena & AdoDataGrid.Recordset.fields(I).Name & " = " & AdoDataGrid.Recordset.fields(I) & vbCrLf
       Next I
       'MsgBox Cadena
       End If
End Sub

Public Sub LeerXML()
   Dim doc As New MSXML2.DOMDocument
   Dim nodeList As MSXML2.IXMLDOMNodeList
   Dim nodeList1 As MSXML2.IXMLDOMNodeList
   Dim node As MSXML2.IXMLDOMNode
   Dim node1 As MSXML2.IXMLDOMNode
   Dim success As Boolean
   Dim IdXML As Long
   Dim IdXML1 As Long
   Dim nodeName As String

   success = doc.Load("C:\SYSBASES\CE\CE999\Comprobantes Generados\0309202101070216417900110010030001222791234567815.xml")
   If success = False Then
      MsgBox doc.parseError.reason
   Else
      Set nodeList = doc.selectNodes("/factura/infoTributaria")
      Cadena = ""
      If Not nodeList Is Nothing Then
         For Each node In nodeList
             For IdXML = 0 To node.childNodes.Length - 1
                 nodeName = node.childNodes.Item(IdXML).nodeName
                 Cadena = Cadena & node.selectSingleNode(nodeName).nodeName & " = " & node.selectSingleNode(nodeName).Text & vbCrLf
             Next IdXML
         Next node
      End If
      Cadena = Cadena & vbCrLf
      
      Set nodeList = doc.selectNodes("/factura/infoFactura")
      If Not nodeList Is Nothing Then
         For Each node In nodeList
             For IdXML = 0 To node.childNodes.Length - 1
                 nodeName = node.childNodes.Item(IdXML).nodeName
                 'MsgBox node.childNodes.Item(IdXML).nodeType
                 If nodeName = "totalConImpuestos" Then
                    
                    Set nodeList1 = doc.selectNodes("/factura/infoFactura/totalConImpuestos/totalImpuesto")
                    If Not nodeList1 Is Nothing Then
                        For Each node1 In nodeList1
                            For IdXML1 = 0 To node1.childNodes.Length - 1
                                nodeName = node1.childNodes.Item(IdXML1).nodeName
                                Cadena = Cadena & vbTab & node1.selectSingleNode(nodeName).nodeName & " = " & node1.selectSingleNode(nodeName).Text & vbCrLf
                            Next IdXML1
                        Next node1
                    End If
                 Else
                    Cadena = Cadena & node.selectSingleNode(nodeName).nodeName & " = " & node.selectSingleNode(nodeName).Text & vbCrLf
                 End If
                 
             Next IdXML
         Next node
         
      End If
      MsgBox Cadena
   End If
End Sub

'''Sub xyz()
'''
'''Dim Browser As SHDocVw.InternetExplorer 'Microsoft Internet Controls
'''Dim HTMLdoc As MSHTML.HTMLDocument 'Microsoft HTML Object Library
'''Dim URL As String
'''
'''  URL = "http://www.bbc.co.uk/news"
'''  Set Browser = New InternetExplorer
'''    Browser.Silent = True
'''    Browser.Navigate URL
'''    Browser.Visible = True
'''  Do
'''    Loop Until Browser.readyState = READYSTATE_COMPLETE
'''
'''    Set HTMLdoc = Browser.Document
'''
'''End Sub

Public Sub Recepcion_SRI_XML(ClaveAcceso As String)
Dim RutaXMLAutorizado As String
Dim RutaXMLRechazado As String
Dim Documento As String
    MsgBox Len(ClaveAcceso) & "-" & ClaveAcceso
    If Len(ClaveAcceso) = 49 Then
      'RutaXML = RutaDocumentos & "\Comprobantes Generados\" & ClaveDeAcceso & ".xml"
      'RutaXMLFirmado = RutaDocumentos & "\Comprobantes Firmados\" & ClaveDeAcceso & ".xml"
       RutaXMLAutorizado = RutaDocumentos & "\Comprobantes Autorizados\" & ClaveAcceso & ".xml"
       RutaXMLRechazado = RutaDocumentos & "\Comprobantes no Autorizados\" & ClaveAcceso & ".xml"
       SRI_Autorizacion = SRI_Leer_XML_Autorizado(RutaXMLAutorizado, RutaXMLRechazado)
       TextoFileEmp = SRI_Autorizacion.Documento_XML
       MsgBox SRI_Autorizacion.Documento_XML
       I = InStr(TextoFileEmp, "<![CDATA[")
       F = InStr(TextoFileEmp, "]]></comprobante>")
       If I > 0 And F > 0 Then
          I = I + 9
          Documento = TrimStrg(MidStrg(TextoFileEmp, I, F - I))
          Escribir_Archivo RutaXMLAutorizado, Documento
          
       End If
    End If

End Sub
