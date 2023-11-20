VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FRelacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RELACION DE DEPENDENCIA"
   ClientHeight    =   4044
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   9036
   Icon            =   "FRelacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4044
   ScaleWidth      =   9036
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Salir"
      Height          =   765
      Left            =   7560
      Picture         =   "FRelacion.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Salir"
      Top             =   324
      Width           =   990
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   750
      Left            =   6480
      Picture         =   "FRelacion.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Grabar"
      Top             =   324
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Height          =   1956
      Left            =   108
      TabIndex        =   16
      Top             =   1944
      Width           =   8796
      Begin VB.TextBox TxtNumSerieUno 
         Height          =   336
         Left            =   324
         MaxLength       =   3
         TabIndex        =   30
         Text            =   "001"
         ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
         Top             =   1296
         Width           =   645
      End
      Begin VB.TextBox TxtNumSerieDos 
         Height          =   336
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   31
         Text            =   "001"
         ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
         Top             =   1296
         Width           =   645
      End
      Begin VB.TextBox TxtNumSerietres 
         Height          =   336
         Left            =   1836
         MaxLength       =   7
         TabIndex        =   32
         Text            =   "0000001"
         ToolTipText     =   $"FRelacion.frx":0B8E
         Top             =   1296
         Width           =   750
      End
      Begin VB.TextBox TxtBaseImp 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   5940
         TabIndex        =   26
         ToolTipText     =   "Verificador (1 caracter)"
         Top             =   432
         Width           =   1308
      End
      Begin VB.TextBox TxtNumAutor 
         Alignment       =   1  'Right Justify
         Height          =   336
         Left            =   2808
         MaxLength       =   10
         TabIndex        =   34
         Text            =   "0000000001"
         Top             =   1296
         Width           =   1416
      End
      Begin VB.TextBox TxtValorTot 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5940
         TabIndex        =   38
         ToolTipText     =   "Verificador (1 caracter)"
         Top             =   1296
         Width           =   1308
      End
      Begin VB.TextBox TxtNumComp 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   4536
         MaxLength       =   1
         TabIndex        =   36
         ToolTipText     =   "Verificador (1 caracter)"
         Top             =   1296
         Width           =   876
      End
      Begin VB.TextBox TxtOtrosIng 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   4536
         TabIndex        =   24
         ToolTipText     =   "Verificador (1 caracter)"
         Top             =   432
         Width           =   1308
      End
      Begin VB.TextBox TxtAporte 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   3024
         TabIndex        =   22
         ToolTipText     =   "Verificador (1 caracter)"
         Top             =   432
         Width           =   1308
      End
      Begin VB.TextBox TxtPorcenAporte 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1728
         MaxLength       =   5
         TabIndex        =   20
         Text            =   "9.35"
         Top             =   432
         Width           =   660
      End
      Begin VB.TextBox TxtSalario 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   324
         TabIndex        =   18
         ToolTipText     =   "Verificador (1 caracter)"
         Top             =   432
         Width           =   1308
      End
      Begin MSDataListLib.DataCombo DCValorRet 
         Bindings        =   "FRelacion.frx":0C31
         DataSource      =   "AdoPorIva"
         Height          =   288
         Left            =   7452
         TabIndex        =   28
         ToolTipText     =   $"FRelacion.frx":0C49
         Top             =   432
         Width           =   1176
         _ExtentX        =   2074
         _ExtentY        =   508
         _Version        =   393216
         Text            =   ""
      End
      Begin MSForms.Label Label12 
         Height          =   228
         Left            =   336
         TabIndex        =   29
         Top             =   1080
         Width           =   2112
         Caption         =   "No. de Serie y Secuencial"
         Size            =   "3725;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label8 
         Height          =   228
         Left            =   5940
         TabIndex        =   25
         Top             =   216
         Width           =   1524
         Caption         =   "Base Imponible"
         Size            =   "2688;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label10 
         Height          =   228
         Left            =   7452
         TabIndex        =   27
         Top             =   216
         Width           =   1308
         Caption         =   "% Valor Reten."
         Size            =   "2307;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   228
         Left            =   5940
         TabIndex        =   37
         Top             =   1080
         Width           =   1632
         Caption         =   "Valor Total Retenido"
         Size            =   "2879;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   228
         Left            =   2808
         TabIndex        =   33
         Top             =   1080
         Width           =   1308
         Caption         =   "No. Autorización"
         Size            =   "2307;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label5 
         Height          =   228
         Left            =   4536
         TabIndex        =   35
         Top             =   1080
         Width           =   1308
         Caption         =   "No. Comp. Ret."
         Size            =   "2307;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label4 
         Height          =   228
         Left            =   4536
         TabIndex        =   23
         Top             =   216
         Width           =   1308
         Caption         =   "Otros Ingresos"
         Size            =   "2307;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label3 
         Height          =   228
         Left            =   3024
         TabIndex        =   21
         Top             =   216
         Width           =   1308
         Caption         =   "Aporte Personal"
         Size            =   "2307;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   228
         Left            =   1728
         TabIndex        =   19
         Top             =   216
         Width           =   1308
         Caption         =   "% Aporte Pers."
         Size            =   "2307;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   228
         Left            =   324
         TabIndex        =   17
         Top             =   216
         Width           =   1308
         Caption         =   "Salario Básico"
         Size            =   "2307;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame FrmRetencion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   108
      TabIndex        =   0
      Top             =   216
      Width           =   3840
      Begin VB.OptionButton OpcSi 
         Caption         =   "SI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2376
         TabIndex        =   4
         ToolTipText     =   $"FRelacion.frx":0CDB
         Top             =   540
         Width           =   552
      End
      Begin VB.OptionButton OpcNo 
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3024
         TabIndex        =   5
         Top             =   540
         Value           =   -1  'True
         Width           =   660
      End
      Begin MSDataListLib.DataCombo DCMes 
         Bindings        =   "FRelacion.frx":0D65
         DataSource      =   "AdoMes"
         Height          =   288
         Left            =   2268
         TabIndex        =   2
         Top             =   216
         Width           =   1416
         _ExtentX        =   2498
         _ExtentY        =   508
         _Version        =   393216
         Text            =   ""
      End
      Begin MSForms.Label Label11 
         Height          =   228
         Left            =   108
         TabIndex        =   1
         Top             =   216
         Width           =   2028
         Caption         =   "Mes de la Retención"
         Size            =   "3577;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label9 
         Height          =   228
         Left            =   108
         TabIndex        =   3
         Top             =   648
         Width           =   2028
         Caption         =   "Sistema del Salario Neto"
         Size            =   "3577;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame FrmTipoComprob 
      Height          =   960
      Left            =   3888
      TabIndex        =   6
      Top             =   216
      Width           =   2010
      Begin VB.TextBox TxtNumeroC 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   972
         MaxLength       =   7
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "FRelacion.frx":0D7A
         ToolTipText     =   "En este campo se debe ingresar el número del comprobante, el cual no excedera los siete caracteres"
         Top             =   540
         Width           =   984
      End
      Begin VB.ComboBox CTP 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   8
         ToolTipText     =   "En este combo se despliega una lista con lo stipos de comprobantes existentes tales como: Comprobante Diario, Ingreso o Egreso"
         Top             =   540
         Width           =   765
      End
      Begin MSForms.Label Label40 
         Height          =   228
         Left            =   972
         TabIndex        =   9
         Top             =   216
         Width           =   768
         Caption         =   "Número"
         Size            =   "1355;402"
         FontName        =   "Arial"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label39 
         Height          =   228
         Left            =   108
         TabIndex        =   7
         Top             =   216
         Width           =   444
         Caption         =   "Tipo"
         Size            =   "783;402"
         FontName        =   "Arial"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSDataListLib.DataCombo DCProveedor 
      Bindings        =   "FRelacion.frx":0D7E
      DataSource      =   "AdoClientes"
      Height          =   288
      Left            =   108
      TabIndex        =   12
      Top             =   1620
      Width           =   5736
      _ExtentX        =   10118
      _ExtentY        =   508
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoAsientoAir 
      Height          =   336
      Left            =   3672
      Top             =   4428
      Visible         =   0   'False
      Width           =   2532
      _ExtentX        =   4466
      _ExtentY        =   572
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
      Caption         =   "AsientoAir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoMes 
      Height          =   336
      Left            =   6264
      Top             =   4752
      Visible         =   0   'False
      Width           =   2208
      _ExtentX        =   3895
      _ExtentY        =   593
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
      Caption         =   "Mes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoAsientoDependencia 
      Height          =   336
      Left            =   6264
      Top             =   5076
      Visible         =   0   'False
      Width           =   2208
      _ExtentX        =   3895
      _ExtentY        =   593
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
      Caption         =   "Asiento Dependencia"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoTransAir 
      Height          =   336
      Left            =   3672
      Top             =   4752
      Visible         =   0   'False
      Width           =   2532
      _ExtentX        =   4466
      _ExtentY        =   572
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
      Caption         =   "TransAir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   336
      Left            =   6264
      Top             =   4428
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   593
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
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoDependencia 
      Height          =   336
      Left            =   3672
      Top             =   5076
      Visible         =   0   'False
      Width           =   2532
      _ExtentX        =   4466
      _ExtentY        =   572
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
      Caption         =   "Dependencia"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   336
      Left            =   1944
      Top             =   4428
      Visible         =   0   'False
      Width           =   1668
      _ExtentX        =   2942
      _ExtentY        =   593
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
      Caption         =   "Clientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label LblTD 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   336
      Left            =   5940
      TabIndex        =   13
      Top             =   1620
      Width           =   336
   End
   Begin MSForms.Label Label41 
      Height          =   228
      Left            =   108
      TabIndex        =   11
      Top             =   1404
      Width           =   876
      Caption         =   "Proveedor"
      Size            =   "1545;402"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label17 
      Height          =   228
      Left            =   6372
      TabIndex        =   14
      Top             =   1404
      Width           =   1740
      Caption         =   "No. de Identificación"
      Size            =   "3069;402"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label LblNumIdent 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   6264
      TabIndex        =   15
      Top             =   1620
      Width           =   1848
   End
End
Attribute VB_Name = "FRelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DirecNum, SalN, Telef As String
Dim DirecProv As Byte

Private Sub CmdCerrar_Click()
  Unload Me
End Sub

Private Sub CmdGrabar_Click()
  'Valido por si acaso exista algun valor con 0
  TextoValido TxtBaseImp, True, , 2
  TextoValido TxtOtrosIng, True, , 2
  TextoValido TxtBaseImp, True, , 2
  TextoValido TxtAporte, True, , 2
  TextoValido TxtValorTot, True, , 2
  'Pregunto antes de grabar
  Titulo = "GRABAR RELACION DE DEPENDENCIA"
  Mensajes = "Desea Grabar los Datos"
  If BoxMensaje = vbYes Then
     'Borrar todas las transacciones de compras que tengan la misma factura y la misma retencion
     'del mismo proveedor
     sSQL = "DELETE * " _
          & "FROM Trans_Dependencia " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND IdeRDEP = '" & CodigoCliente & "' "
     ConectarAdoExecute sSQL
     
     '& "AND Fecha = #" & BuscarFecha(MBFechaRegis) & "# " _
     'Si existe la misma retencion la borramos para quede la actual
     sSQL = "DELETE * " _
          & "FROM Trans_Air " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND IdProv = '" & CodigoCliente & "' " _
          & "AND EstabRetencion = '" & TxtNumSerieUno & "' " _
          & "AND PtoEmiRetencion = '" & TxtNumSerieDos & "' " _
          & "AND SecRetencion = " & Convertir_Numero(TxtNumSerietres) & " " _
          & "AND AutRetencion = '" & TxtNumAutor & "' " _
          & "AND Tipo_Trans = 'D' "
     ConectarAdoExecute sSQL
     Grabacion
     Mensajes = "Los Datos fueron grabados correctamente" & vbCrLf _
              & "Desea ingresar otra transacción"""
     If BoxMensaje = vbYes Then
        Ln_No = 1
        Limpiar_Controles
        Listar_Air
        DCMes.SetFocus
     Else
        Unload FRelacion
     End If
   Else
      DCMes.SetFocus
   End If

End Sub

Public Sub Grabacion()
   'Grabo en el Asiento_Compras e implicito Asiento_Air
   'MsgBox Max_ID("Trans_Compras")
    SetAdoAddNew "Asiento_Dependencia"
    SetAdoFields "IdeRDEP", CodigoCliente
    SetAdoFields "TipDocRDEP", 2 'Ojo no se de donde sale este valor
    SetAdoFields "DirCalle", DireccionCli
    SetAdoFields "DirNumCalle", DirecNum
    SetAdoFields "DirCiu", 0 'No se de donde sale este codigo
    SetAdoFields "DirProv", DirecProv
    SetAdoFields "Telef", Telef
    SetAdoFields "SisSalNet", SalN
    SetAdoFields "ValIngLiq", Convertir_Numero(TxtSalario, 2)
    SetAdoFields "ApoPerIess", Convertir_Numero(TxtAporte, 2)
    SetAdoFields "BasImp", Convertir_Numero(TxtBaseImp, 2)
    SetAdoFields "PorValRet", DCValorRet
    SetAdoFields "OtroIng", Convertir_Numero(TxtOtrosIng, 2)
    SetAdoFields "ValRet", Convertir_Numero(TxtValorTot, 2)
    SetAdoFields "AñoRet", Year(Date)
    SetAdoFields "NumRet", TxtNumComp
    SetAdoFields "A_No", Ln_No
    SetAdoFields "T_No", Trans_No
    SetAdoUpdate
    'Grabamos los datos de la transaccion en la tabla definitiva de almacenamiento
    ID_Trans = Maximo_De("Trans_Dependencia", "ID")  'va a tener el indice de transaccion unico para que no exista duplicados en a base
    sSQL = "SELECT * " _
         & "FROM Asiento_Dependencia " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "ORDER BY T_No "
    SelectAdodc AdoAsientoDependencia, sSQL
    With AdoAsientoDependencia.Recordset
     If .RecordCount > 0 Then
         FechaTexto = Date
         SetAdoAddNew "Trans_Dependencia"
         SetAdoFields "T", Normal
         SetAdoFields "IdeRDEP", .Fields("IdeRDEP")
         SetAdoFields "TipDocRDEP", .Fields("TipDocRDEP")
         SetAdoFields "DirCalle", .Fields("DirCalle")
         SetAdoFields "DirNumCalle", .Fields("DirNumCalle")
         SetAdoFields "DirCiu", .Fields("DirCiu")
         SetAdoFields "DirProv", .Fields("DirProv")
         SetAdoFields "Telef", .Fields("Telef")
         SetAdoFields "SisSalNet", .Fields("SisSalNet")
         SetAdoFields "ValIngLiq", .Fields("ValIngLiq")
         SetAdoFields "ApoPerIess", .Fields("ApoPerIess")
         SetAdoFields "BasImp", .Fields("BasImp")
         SetAdoFields "PorValRet", .Fields("PorValRet")
         SetAdoFields "OtroIng", .Fields("OtroIng")
         SetAdoFields "ValRet", .Fields("ValRet")
         SetAdoFields "AñoRet", .Fields("AñoRet")
         SetAdoFields "NumRet", .Fields("NumRet")
         SetAdoFields "TP", CTP
         SetAdoFields "Numero", TxtNumeroC
         SetAdoFields "Fecha", FechaTexto
         SetAdoFields "ID", ID_Trans
         SetAdoFields "Linea_SRI", Ln_No
         SetAdoUpdate
     End If
    End With
    'Grabo en el Asiento Air
''    Insertar_AsientoAir
''   'Selecciona el numero mayor para continuar la secuencia en el
''    ID_Trans = Maximo_De("Trans_Air", "ID")
''    sSQL = "SELECT * " _
''         & "FROM Asiento_Air " _
''         & "WHERE Item = '" & NumEmpresa & "' " _
''         & "AND CodigoU = '" & CodigoUsuario & "' " _
''         & "AND T_No = " & Trans_No & " " _
''         & "AND Tipo_Trans = 'D' " _
''         & "ORDER BY A_No "
''    SelectAdodc AdoTransAir, sSQL
''    With AdoAsientoAir.Recordset
''     If .RecordCount > 0 Then
''         Do While Not .EOF
''            '& "AND FechaEmiRet = #" & FechaTexto & "# " _ '
''            SetAdoAddNew "Trans_Air"
''            SetAdoFields "T", Normal
''            SetAdoFields "CodRet", .Fields("CodRet")
''            SetAdoFields "BaseImp", .Fields("BaseImp")
''            SetAdoFields "Porcentaje", .Fields("Porcentaje")
''            SetAdoFields "ValRet", .Fields("ValRet")
''            SetAdoFields "EstabRetencion", .Fields("EstabRetencion")
''            SetAdoFields "PtoEmiRetencion", .Fields("PtoEmiRetencion")
''            SetAdoFields "SecRetencion", .Fields("SecRetencion")
''            SetAdoFields "AutRetencion", .Fields("AutRetencion")
''            SetAdoFields "Tipo_Trans", .Fields("Tipo_Trans")
''            SetAdoFields "Numero", TxtNumeroC
''            SetAdoFields "IdProv", CodigoCliente
''            SetAdoFields "TP", CTP
''            SetAdoFields "Cta_Retencion", .Fields("Cta_Retencion")
''            SetAdoFields "EstabFactura", .Fields("EstabFactura")
''            SetAdoFields "PuntoEmiFactura", .Fields("PuntoEmiFactura")
''            SetAdoFields "Factura_No", .Fields("Factura_No")
''            SetAdoFields "Fecha", FechaTexto
''            SetAdoFields "ID", ID_Trans
''            SetAdoFields "Linea_SRI", 0
''            SetAdoUpdate
''           .MoveNext
''         Loop
''      End If
''    End With
End Sub


Private Sub CMesRet_Click()
  NumMeses = 1
  With AdoMes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Mes = '" & DCMes & "' ")
       If Not .EOF Then NumMeses = .Fields("NoMes")
   End If
  End With
End Sub

Private Sub CMesRet_DblClick()

End Sub

Private Sub CTP_GotFocus()
  MarcarTexto CTP
End Sub

Private Sub CTP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CTP_LostFocus()
  If IsNumeric(CTP.Text) Then
     MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
     CTP.Text = ""
     CTP.Text = "CD"
     CTP.SetFocus
  End If
End Sub

Private Sub DCMes_LostFocus()
  NumMeses = 1
  With AdoMes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Mes = '" & DCMes & "' ")
       If Not .EOF Then NumMeses = .Fields("NoMes")
   End If
  End With
End Sub

Private Sub DCProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCProveedor_LostFocus()
  If IsNumeric(DCProveedor.Text) Then
     MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCProveedor.Text = ""
     Leer_Clientes
     DCProveedor.SetFocus
  Else
     NombreCliente = UCase(DCProveedor)
     With AdoClientes.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Cliente = '" & NombreCliente & "' ")
          If Not .EOF Then
            'Busca y captura el codigo de Porcentaje IVA
            CodigoCliente = .Fields("Codigo")
            DireccionCli = .Fields("DireccionT")
            DirecNum = .Fields("DirNumero")
            DirecProv = .Fields("Prov")
            Telef = .Fields("Telefono")
            CICliente = .Fields("CI_RUC")
            TipoBenef = .Fields("TD")
            LblNumIdent = CICliente
            LblTD.Caption = TipoBenef
              
            TxtNumSerietres = "0000001"
            'Aqui despliego el ultimo numero de la Transaccion
            sSQL = "SELECT TOP 1 * " _
                 & "FROM Trans_Compras " _
                 & "WHERE IdProv = '" & CodigoCliente & "' " _
                 & "AND Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "ORDER BY Secuencial DESC "
            SelectAdodc AdoAux, sSQL
            With AdoAux.Recordset
             If .RecordCount > 0 Then TxtNumSerietres = .Fields("Secuencial")
            End With
         Else
            FClientesFlash.Show
            Leer_Clientes
         End If
      Else
         FClientesFlash.Show
         Leer_Clientes
      End If
    End With
  End If
End Sub

Private Sub Form_Activate()
  Ln_No = 1
  CTP.Clear
  CTP.AddItem "CE"
  CTP.AddItem "CI"
  CTP.AddItem "CD"
  CTP.Text = "CE"
  Carga_Datos_Iniciales
End Sub

Public Sub Carga_Datos_Iniciales()
  Encerar_Var
  Limpiar_Controles
  TxtNumSerieUno = "001"
  TxtNumSerieDos = "001"
  TxtNumSerietres = "0000001"
  'Cargo los meses
  Carga_Meses
  'Carga en el Data Combo los Clientes con su RUC
  Leer_Clientes
End Sub

Public Sub Leer_Clientes()
  'Carga en el Data Combo los Clientes con su RUC
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Cliente <> '.' " _
       & "AND TD <>  'E' " _
       & "ORDER BY Cliente "
  SelectDBCombo DCProveedor, AdoClientes, sSQL, "Cliente"
End Sub

Private Sub Form_Load()
  CentrarForm FImportaciones
  ConectarAdodc AdoAux
  ConectarAdodc AdoMes
  ConectarAdodc AdoAsientoDependencia
  ConectarAdodc AdoDependencia
  ConectarAdodc AdoAsientoAir
  ConectarAdodc AdoTransAir
  ConectarAdodc AdoClientes
End Sub

Private Sub OpcNo_Click()
  If OpcNo.Value = True Then SalN = "N"
End Sub

Private Sub OpcSi_Click()
  If OpcSi.Value = True Then SalN = "S"
  
End Sub

Private Sub TxtBaseImp_GotFocus()
  MarcarTexto TxtBaseImp
End Sub

Private Sub TxtBaseImp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImp_LostFocus()
  TextoValido TxtBaseImp, True, , 0
  
  'Cargo los Porcentajes de la Tabla Tipo_Renta
  Carga_Porcen_Ret
End Sub

Private Sub TxtNumAutor_GotFocus()
   MarcarTexto TxtNumAutor
End Sub

Private Sub TxtNumAutor_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumAutor_LostFocus()
  If Val(TxtNumAutor) <= 0 Then TxtNumAutor = "0000000001"
  TxtNumAutor = Format(Val(Round(TxtNumAutor)), String(10, "0"))
End Sub

Private Sub TxtNumComp_LostFocus()
  'Calculo el Valor Retenido de acuerdo a la Tabla
  If SaldoInic <> 0 Then
     TxtValorTot = Saldo - SaldoInic
  Else
     TxtValorTot = 0
  End If
End Sub

Private Sub TxtNumeroC_GotFocus()
  MarcarTexto TxtNumeroC
End Sub

Private Sub TxtNumeroC_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumeroC_LostFocus()
  If Not IsNumeric(TxtNumeroC.Text) Then
     MsgBox "No ingrese caracteres alfabéticos. Vuelva a ingresar.", vbInformation, "Aviso"
     TxtNumeroC.Text = ""
     TxtNumeroC.SetFocus
  End If
  If TxtNumeroC.Text = "." Or TxtNumeroC.Text = " " Then
     MsgBox "Ingrese el Número de Comprobante", vbInformation, "Aviso"
  End If
  If Val(TxtNumeroC) <= 0 Then TxtNumeroC = "0"
  TxtNumeroC = Format(Convertir_Numero(TxtNumeroC), String(7, "0"))
End Sub

Private Sub TxtNumSerieDos_GotFocus()
  MarcarTexto TxtNumSerieDos
End Sub

Private Sub TxtNumSerieDos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieDos_LostFocus()
  TextoValido TxtNumSerieDos, True, , 0
  If Val(TxtNumSerieDos) <= 0 Then TxtNumSerieDos = "001"
  TxtNumSerieDos = Format(Val(TxtNumSerieDos), "000")
End Sub

Private Sub TxtNumSerietres_GotFocus()
  MarcarTexto TxtNumSerietres
End Sub

Private Sub TxtNumSerietres_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerietres_LostFocus()
  If Val(TxtNumSerietres) <= 0 Then TxtNumSerietres = "0000001"
  TxtNumSerietres = Format(Val(Round(TxtNumSerietres)), "0000000")
End Sub

Private Sub TxtNumSerieUno_GotFocus()
  MarcarTexto TxtNumSerieUno
End Sub

Private Sub TxtNumSerieUno_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieUno_LostFocus()
  TextoValido TxtNumSerieUno, True, , 0
  If Val(TxtNumSerieUno) <= 0 Then TxtNumSerieUno = "001"
  TxtNumSerieUno = Format(Val(TxtNumSerieUno), "000")
End Sub

Private Sub TxtOtrosIng_GotFocus()
  MarcarTexto TxtOtrosIng
End Sub

Private Sub TxtOtrosIng_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtOtrosIng_LostFocus()
  TextoValido TxtOtrosIng, True, , 0
  'Calculo el salario - el aporte para colocar en la Base Imponible
  TxtBaseImp = Convertir_Numero(TxtSalario, 2) - Convertir_Numero(TxtAporte, 2)
End Sub

Private Sub TxtPorcenAporte_LostFocus()
Dim Aporte As Double
  Aporte = 0.0935
  TxtAporte = Convertir_Numero(TxtSalario, 2) * Aporte
End Sub

Private Sub TxtSalario_GotFocus()
  MarcarTexto TxtSalario
End Sub

Private Sub TxtSalario_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtSalario_LostFocus()
  TextoValido TxtSalario, True, , 0
End Sub

Private Sub TxtValorTot_GotFocus()
  MarcarTexto TxtValorTot
End Sub

Private Sub TxtValorTot_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtValorTot_LostFocus()
  TextoValido TxtValorTot, True, , 0
End Sub

Public Sub Listar_Air()
  'Enceramos el espacio de cada usuario para emprezar con una nueva retencion
   sSQL = "DELETE * " _
        & "FROM Asiento_Dependencia " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   ConectarAdoExecute sSQL
  'Borra Asiento Air
   sSQL = "DELETE * " _
        & "FROM Asiento_Air " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Tipo_Trans = 'D' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   ConectarAdoExecute sSQL
  
End Sub

Public Sub Carga_Porcen_Ret()
Dim FechaRet As Date
  Saldo = Convertir_Numero(TxtSalario, 2) * 12
 'Fecha Normal
  FechaInicial = "01/" & Format(NumMeses, "00") & "/" & Format(Year(Date), "0000")
  'FechaMitad = "15/" & Format(NumMeses, "00") & "/" & Format(CAño, "0000")
  FechaFinal = UltimoDiaMes(FechaInicial)
 'Convertir fecha segun la plataforma
  FechaIni = BuscarFecha(FechaInicial)
  'FechaMid = BuscarFecha(FechaMitad)
  FechaFin = BuscarFecha(FechaFinal)
  
 'Carga la Tabla de Porcentajes Retenidos
  sSQL = "SELECT * " _
       & "FROM Tipo_Renta " _
       & "WHERE Fecha_Inicio <= #" & FechaIni & "# " _
       & "AND Fecha_Final >= #" & FechaFin & "# " _
       & "AND Desde <= " & Saldo & "  " _
       & "AND Hasta >= " & Saldo & "  " _
       & "ORDER BY Excede "
  SelectDBCombo DCValorRet, AdoAux, sSQL, "Excede"
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          If DCValorRet > 0 Then
             'Asigno en una variable el Campo Desde para restar
             SaldoInic = .Fields("Desde")
          End If
          .MoveNext
       Loop
    End If
  End With
End Sub

Sub Carga_Meses()
  sSQL = "SELECT * " _
       & "FROM Tabla_Meses " _
       & "WHERE NoMes <> 0 " _
       & "ORDER BY NoMes "
  SelectDBCombo DCMes, AdoMes, sSQL, "Mes"
  DCMes.SetFocus
End Sub

Public Sub Encerar_Var()
  Ln_No = 0
  Saldo = 0
  SaldoInic = 0
  SalN = "S"
End Sub

Public Sub Limpiar_Controles()
  DCProveedor.Text = ""
  TxtNumeroC.Text = ""
  TxtSalario.Text = ""
  TxtAporte.Text = ""
  TxtOtrosIng.Text = ""
  TxtBaseImp.Text = ""
  DCValorRet.Text = ""
  LblNumIdent.Caption = ""
  LblTD.Caption = ""
  TxtNumSerieUno.Text = ""
  TxtNumSerieDos.Text = ""
  TxtNumSerietres.Text = ""
  TxtNumAutor.Text = ""
  TxtNumComp.Text = ""
  TxtValorTot.Text = ""
  CTP.Clear
  CTP.AddItem "CE"
  CTP.AddItem "CI"
  CTP.AddItem "CD"
  CTP.Text = "CE"
  DCMes.Text = ""
End Sub

Sub Insertar_AsientoAir()
  'Selecciona el numero mayor para continuar la secuencia en el
  'campo T_No y A_No
  Ln_No = Maximo_De("Asiento_Air", "A_No")
  'Espizq = SinEspaciosIzq(DCConceptoRet)
  'Espder = Trim(Mid(DCConceptoRet, Len(Espizq) + 3, Len(DCConceptoRet)))
  SetAdoAddNew "Asiento_Air"
  SetAdoFields "CodRet", "DEP"
  SetAdoFields "Detalle", "Dependencia"
  SetAdoFields "BaseImp", Convertir_Numero(TxtBaseImp, 2)
  SetAdoFields "Porcentaje", Convertir_Numero(DCValorRet) / 100
  SetAdoFields "ValRet", Convertir_Numero(TxtValorTot, 2)
  SetAdoFields "EstabRetencion", TxtNumSerieUno
  SetAdoFields "PtoEmiRetencion", TxtNumSerieDos
  SetAdoFields "SecRetencion", Convertir_Numero(TxtNumSerietres)
  SetAdoFields "AutRetencion", TxtNumAutor
  SetAdoFields "FechaEmiRet", Date
  SetAdoFields "Cta_Retencion", "Relacion Dependencia"
  SetAdoFields "EstabFactura", 0
  SetAdoFields "PuntoEmiFactura", 0
  SetAdoFields "Factura_No", 0
  SetAdoFields "IdProv", CodigoCliente
  SetAdoFields "A_No", Ln_No
  SetAdoFields "T_No", Trans_No
  SetAdoFields "Tipo_Trans", "D"
  SetAdoUpdate
  RatonNormal
End Sub

