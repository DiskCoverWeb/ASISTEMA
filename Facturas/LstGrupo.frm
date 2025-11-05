VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ListarGrupos 
   Caption         =   "LISTADO POR GRUPOS"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   11280
   WindowState     =   1  'Minimized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   28560
      _ExtentX        =   50377
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Modulo"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Generar_Facturas"
            Object.ToolTipText     =   "Generación de Facturas en Bloque"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Listado_x_Grupos"
            Object.ToolTipText     =   "Imprime un listado resumido de los grupo creados"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Generar_Eliminar_Rubros"
            Object.ToolTipText     =   "Generar o Eliminar por Lotes los Rubros a Facturar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Generar_Deuda_Pendiente"
            Object.ToolTipText     =   "Genera las Deudas Pendientes"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Impresora"
            Object.ToolTipText     =   "Imprimir Resultados"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Recalcular_Fechas"
            Object.ToolTipText     =   "Recalcula Fecha de Facturación"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Recibos"
            Object.ToolTipText     =   "Imprimir Recibos"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Enviar a Excel"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Email_Tutor"
            Object.ToolTipText     =   "Enviar al Tutor la Nomina por mail"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   5985
         TabIndex        =   0
         Top             =   0
         Width           =   14400
         Begin VB.CheckBox CheqRangos 
            Caption         =   "&Por Rangos:"
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
            TabIndex        =   1
            Top             =   210
            Width           =   1380
         End
         Begin VB.CheckBox CheqPendientes 
            Caption         =   "Listar Solo Pendientes"
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
            Left            =   5460
            TabIndex        =   4
            Top             =   210
            Value           =   1  'Checked
            Width           =   1380
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&?"
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
            Left            =   13965
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   210
            Width           =   330
         End
         Begin MSDataListLib.DataCombo DCGrupoI 
            Bindings        =   "LstGrupo.frx":0000
            DataSource      =   "AdoGrupo"
            Height          =   360
            Left            =   1575
            TabIndex        =   2
            Top             =   210
            Visible         =   0   'False
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DCTipoPago 
            Bindings        =   "LstGrupo.frx":0017
            DataSource      =   "AdoTipoPago"
            Height          =   315
            Left            =   8505
            TabIndex        =   6
            Top             =   210
            Width           =   5370
            _ExtentX        =   9472
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DCGrupoF 
            Bindings        =   "LstGrupo.frx":0031
            DataSource      =   "AdoGrupo"
            Height          =   360
            Left            =   3360
            TabIndex        =   3
            Top             =   210
            Visible         =   0   'False
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label34 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " TIPO DE PAGO"
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
            Left            =   7035
            TabIndex        =   5
            Top             =   210
            Width           =   1485
         End
      End
   End
   Begin VB.Frame FrmEmail 
      BackColor       =   &H00800000&
      Height          =   2955
      Left            =   105
      TabIndex        =   32
      Top             =   2730
      Visible         =   0   'False
      Width           =   20070
      Begin VB.ListBox LstClientes 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   105
         Style           =   1  'Checkbox
         TabIndex        =   40
         Top             =   1575
         Width           =   10935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Enviar mail >>"
         Height          =   540
         Left            =   15855
         TabIndex        =   42
         Top             =   840
         Width           =   1800
      End
      Begin VB.TextBox TxtMensaje 
         Height          =   855
         Left            =   8505
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   43
         Top             =   525
         Width           =   7260
      End
      Begin VB.CheckBox CheqConDeuda 
         BackColor       =   &H00800000&
         Caption         =   "&Enviar mail con deuda pendiente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   15855
         TabIndex        =   41
         Top             =   210
         Value           =   1  'Checked
         Width           =   1800
      End
      Begin VB.TextBox TxtAsunto 
         Height          =   330
         Left            =   1155
         MaxLength       =   60
         TabIndex        =   36
         Top             =   630
         Width           =   7260
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Adjuntar >>"
         Height          =   330
         Left            =   105
         TabIndex        =   37
         Top             =   1050
         Width           =   1065
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Escriba el mensaje"
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
         Left            =   8505
         TabIndex        =   39
         Top             =   210
         Width           =   7260
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Asunto"
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
         TabIndex        =   35
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label LblArchivo 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1155
         TabIndex        =   38
         Top             =   1050
         Width           =   7260
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Remitente"
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
         TabIndex        =   33
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1155
         TabIndex        =   34
         Top             =   210
         Width           =   7260
      End
   End
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "LstGrupo.frx":0048
      Height          =   3480
      Left            =   105
      TabIndex        =   31
      Top             =   2730
      Width           =   20385
      _ExtentX        =   35957
      _ExtentY        =   6138
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
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
   Begin VB.CommandButton Command2 
      Caption         =   "&S"
      Height          =   330
      Left            =   105
      TabIndex        =   52
      Top             =   6300
      Width           =   330
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   105
      TabIndex        =   7
      Top             =   735
      Width           =   20280
      Begin VB.TextBox TxtNota 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   11550
         MaxLength       =   100
         TabIndex        =   27
         Top             =   1050
         Width           =   8625
      End
      Begin VB.TextBox TxtObs 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   1575
         MaxLength       =   100
         TabIndex        =   25
         Top             =   1050
         Width           =   8625
      End
      Begin VB.CheckBox CheqFA 
         Caption         =   "Fecha FA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9765
         TabIndex        =   16
         Top             =   210
         Width           =   1170
      End
      Begin MSDataListLib.DataCombo DCCliente 
         Bindings        =   "LstGrupo.frx":005F
         DataSource      =   "AdoCliente"
         Height          =   360
         Left            =   2835
         TabIndex        =   15
         Top             =   525
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "Clientes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox CheqResumen 
         Caption         =   "Resumen Periodos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5985
         TabIndex        =   13
         Top             =   210
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.ComboBox CTipoConsulta 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   360
         Left            =   2835
         TabIndex        =   12
         Top             =   210
         Width           =   3060
      End
      Begin MSDataListLib.DataCombo DCProductos 
         Bindings        =   "LstGrupo.frx":0078
         DataSource      =   "AdoProductos"
         Height          =   360
         Left            =   14700
         TabIndex        =   23
         Top             =   630
         Visible         =   0   'False
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "Clientes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.OptionButton OpcActivos 
         Caption         =   "Activos"
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
         Left            =   12390
         TabIndex        =   21
         Top             =   630
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton OpcInactivos 
         Caption         =   "Inactivos"
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
         Left            =   13545
         TabIndex        =   22
         Top             =   630
         Width           =   1170
      End
      Begin VB.CheckBox CheqPorRubro 
         Caption         =   "Por Rubros:"
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
         Left            =   11025
         TabIndex        =   20
         Top             =   630
         Width           =   1380
      End
      Begin VB.CheckBox CheqDesc 
         Caption         =   "D&escuentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8085
         TabIndex        =   14
         Top             =   210
         Width           =   1380
      End
      Begin MSDataListLib.DataCombo DCLinea 
         Bindings        =   "LstGrupo.frx":0093
         DataSource      =   "AdoLinea"
         Height          =   360
         Left            =   14490
         TabIndex        =   19
         Top             =   210
         Visible         =   0   'False
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "CxC Clientes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox MBFechaF 
         Height          =   330
         Left            =   1365
         TabIndex        =   11
         Top             =   525
         Width           =   1275
         _ExtentX        =   2249
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
      Begin VB.CheckBox CheqVenc 
         Caption         =   "Vencimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1365
         TabIndex        =   9
         Top             =   210
         Value           =   1  'Checked
         Width           =   1380
      End
      Begin MSMask.MaskEdBox MBFechaI 
         Height          =   330
         Left            =   105
         TabIndex        =   10
         Top             =   525
         Width           =   1275
         _ExtentX        =   2249
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
      Begin MSMask.MaskEdBox MBFecha 
         Height          =   330
         Left            =   11025
         TabIndex        =   17
         Top             =   210
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
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
      Begin VB.Label LblFactura 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No. 000000-000000000"
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
         Left            =   17955
         TabIndex        =   53
         Top             =   210
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NOTA:"
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
         Left            =   10290
         TabIndex        =   26
         Top             =   1050
         Width           =   1275
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " OBSERVACION"
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
         TabIndex        =   24
         Top             =   1050
         Width           =   1485
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &Linea de Facturacion:"
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
         Left            =   12390
         TabIndex        =   18
         Top             =   210
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &Emision:"
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
         TabIndex        =   8
         Top             =   210
         Width           =   1275
      End
   End
   Begin VB.TextBox TxtAyuda 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2535
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Top             =   8610
      Visible         =   0   'False
      Width           =   7575
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   405
      Left            =   105
      TabIndex        =   30
      Top             =   2310
      Width           =   21750
      _ExtentX        =   38365
      _ExtentY        =   714
      _Version        =   393216
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "LISTADO POR GRUPOS"
      TabPicture(0)   =   "LstGrupo.frx":00AA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "PENSION MENSUAL DEL AÑO"
      TabPicture(1)   =   "LstGrupo.frx":00C6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "ALUMNOS CON DESCUENTO"
      TabPicture(2)   =   "LstGrupo.frx":00E2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "ENVIOS POR MAIL"
      TabPicture(3)   =   "LstGrupo.frx":00FE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "RESUMEN PENSIONES POR MES"
      TabPicture(4)   =   "LstGrupo.frx":011A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "CARTERA RECAUDADA"
      TabPicture(5)   =   "LstGrupo.frx":0136
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "ENVIAR DEUDA POR API Y EMAIL"
      TabPicture(6)   =   "LstGrupo.frx":0152
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).ControlCount=   0
   End
   Begin MSAdodcLib.Adodc AdoCiudad 
      Height          =   330
      Left            =   2835
      Top             =   3990
      Visible         =   0   'False
      Width           =   2010
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
      Caption         =   "Ciudad"
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
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   2835
      Top             =   3675
      Visible         =   0   'False
      Width           =   2010
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
      Caption         =   "Cliente"
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
   Begin MSAdodcLib.Adodc AdoNiveles 
      Height          =   330
      Left            =   2835
      Top             =   6195
      Visible         =   0   'False
      Width           =   2010
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
      Caption         =   "Niveles"
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
   Begin MSAdodcLib.Adodc AdoLinea 
      Height          =   330
      Left            =   2835
      Top             =   5880
      Visible         =   0   'False
      Width           =   2010
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
      Caption         =   "Linea"
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   2835
      Top             =   4305
      Visible         =   0   'False
      Width           =   2010
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
   Begin MSAdodcLib.Adodc AdoGrupo 
      Height          =   330
      Left            =   2835
      Top             =   4935
      Visible         =   0   'False
      Width           =   2010
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
      Caption         =   "Grupo"
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
   Begin MSAdodcLib.Adodc AdoAux2 
      Height          =   330
      Left            =   2835
      Top             =   4620
      Visible         =   0   'False
      Width           =   2010
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
      Caption         =   "Aux2"
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
   Begin MSAdodcLib.Adodc AdoProductos 
      Height          =   330
      Left            =   2835
      Top             =   5250
      Visible         =   0   'False
      Width           =   2010
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
      Caption         =   "Productos"
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
   Begin MSAdodcLib.Adodc AdoTipoPago 
      Height          =   330
      Left            =   2835
      Top             =   5565
      Visible         =   0   'False
      Width           =   2010
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
      Caption         =   "TipoPago"
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
   Begin MSAdodcLib.Adodc AdoQuery 
      Height          =   330
      Left            =   9870
      Top             =   6930
      Width           =   10305
      _ExtentX        =   18177
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
      Caption         =   "Listado de Facturas"
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
   Begin MSAdodcLib.Adodc AdoParte 
      Height          =   330
      Left            =   4935
      Top             =   3675
      Visible         =   0   'False
      Width           =   2010
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
      Caption         =   "Niveles"
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
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total  por Cobrar"
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
      Left            =   6510
      TabIndex        =   49
      Top             =   6930
      Width           =   1590
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SubTotal CxC"
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
      TabIndex        =   48
      Top             =   6930
      Width           =   1380
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Left            =   1470
      TabIndex        =   47
      Top             =   6930
      Width           =   1695
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Anticipos"
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
      Left            =   3255
      TabIndex        =   46
      Top             =   6930
      Width           =   1485
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Left            =   4725
      TabIndex        =   45
      Top             =   6930
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Left            =   8085
      TabIndex        =   44
      Top             =   6930
      Width           =   1695
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   20790
      Top             =   3780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LstGrupo.frx":016E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LstGrupo.frx":0488
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LstGrupo.frx":07A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LstGrupo.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LstGrupo.frx":0DD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LstGrupo.frx":10F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LstGrupo.frx":140A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LstGrupo.frx":1724
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LstGrupo.frx":1A3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LstGrupo.frx":2690
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Ctrl+M> Modificar|<Ctrl+F6> No Modifica|<Ctrl+Ins> Insertar|<Ctrl+B> Buscar|<Ctrl+Supr> Eliminar|<Ctrl+V> Cambio de Valores"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   105
      TabIndex        =   28
      Top             =   11655
      Width           =   16290
   End
End
Attribute VB_Name = "ListarGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''  Keys_Especiales Shift
''  If CtrlDown And KeyCode = vbKeyF1 Then GenerarDataTexto ListarGrupos, AdoNomina
''Lstclientes.Selected(I) = True

Dim Tiene_Cursos As Boolean
Dim PorGrupo As Boolean
Dim PorDireccion As Boolean
Dim ListaDeCampos As String
Dim Tutor As String
Dim TutorEmail As String
Dim MesesNominaGrupo As String

Public Sub Reporte_CxC_Mes_PDF(VerDocumento As Boolean, PorValor As Boolean)
Dim ValorCampo As String
Dim EmailCli As String
Dim PosCampo(1 To 15) As Single
Dim Cont As Byte
    DGQuery.Visible = False
    If CFechaLong(MBFechaI) <= CFechaLong(MBFechaF) And Len(MesesNominaGrupo) > 1 Then
       RatonReloj
       Cont = 1
       If Len(Tutor) = 1 Then Tutor = "Ninguno"
       With AdoQuery.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
            DireccionCli = .fields("Direccion")
            EmailCli = ""
            Insertar_Mail EmailCli, TutorEmail
             
            RutaDocumentoPDF = ""
            MensajeEncabData = "REPORTE CARTERA DE " & DireccionCli
            SQLMsg1 = ""
            SQLMsg2 = ""
            SQLMsg3 = ""
            NombreArchivo = "Nomina_CxC_Meses-" & Replace(Tutor, " ", "_")
            SetNombrePRN = Impresota_PDF
           'Geneeramos el documento
            tPrint.TipoImpresion = Es_PDF
            tPrint.NombreArchivo = NombreArchivo
            tPrint.TituloArchivo = "Reporte de CxC Por Meses " & Tutor
            tPrint.TipoLetra = TipoHelvetica
            tPrint.OrientacionPagina = 1
            tPrint.PaginaA4 = True
            tPrint.EsCampoCorto = True
            tPrint.VerDocumento = VerDocumento
            Set cPrint = New cImpresion
            cPrint.iniciaImpresion
            PosLinea = 1
            PosCampo(1) = 1.1    ' No
            PosCampo(2) = 2.9    ' Cedula
            PosCampo(3) = 4.5    ' Apellidos y Nombres
            PosCampo(4) = 8.5    ' Enero
            PosCampo(5) = 9.5    ' Febrero
            PosCampo(6) = 10.5   ' Marzo
            PosCampo(7) = 11.5   ' Abril
            PosCampo(8) = 12.5   ' Mayo
            PosCampo(9) = 13.5   ' Junio
            PosCampo(10) = 14.5  ' Julio
            PosCampo(11) = 15.5  ' Agosto
            PosCampo(12) = 16.5  ' Septiembre
            PosCampo(13) = 17.5  ' Octubre
            PosCampo(14) = 18.5  ' Noviembre
            PosCampo(15) = 19.5  ' Diciembre
            cPrint.printEncabezado 1.2, PosLinea, TipoHelvetica
           'Pagina No. 1
            
            TipoCta = Ninguno
            SerieFactura = Ninguno
            Factura_No = 0
            CodigoP = Ninguno
               
            cPrint.fondoDeLetra = Negro
            cPrint.tipoNegrilla = True
            cPrint.PorteDeLetra = 8
            cPrint.colorDeLetra = Negro
            'PosLinea = 3.1
            cPrint.printTexto PosCampo(1), PosLinea, "TUTOR: " & Tutor
            PosLinea = PosLinea + 0.4
            
            Cadena = "La informacion presente reposa en la base de dato de la Institucion, corte realizado desde " & FechaStrg(MBFechaI) & " al " & FechaStrg(FechaSistema) & ", " _
                   & "cualquier informacion adicional comuniquese a la institucion."
            PosLinea = cPrint.printTextoMultiple(1.1, PosLinea, Cadena, 18)
            PosLinea = PosLinea + 0.5
            PrimeraLinea = PosLinea
            cPrint.printLinea 2.3, PosLinea, cPrint.dAnchoPapel - 1.1, PosLinea
            PosLinea = PosLinea + 0.1
            cPrint.printTexto PosCampo(1), PosLinea, "No"
            cPrint.printTexto PosCampo(2), PosLinea, "Cedula"
            cPrint.printTexto PosCampo(3), PosLinea, "Apellidos_Nombres"
            For I = 4 To 15
                cPrint.printTexto PosCampo(I), PosLinea, MidStrg(.fields(I - 2).Name, 1, 3)
            Next I
            PosLinea = PosLinea + 0.4
            cPrint.printLinea 2.3, PosLinea, cPrint.dAnchoPapel - 1.1, PosLinea
            PosLinea = PosLinea + 0.1
            cPrint.colorDeLetra = Negro
            cPrint.tipoNegrilla = False
            cPrint.PorteDeLetra = 7
            Do While Not .EOF
               cPrint.printTexto PosCampo(1), PosLinea, Format(Cont, "00")
               cPrint.printFields PosCampo(2), PosLinea, .fields("CI_RUC")
               cPrint.printTexto PosCampo(3), PosLinea, ULCase(.fields("Cliente"))
               cPrint.printCuadro PosCampo(4) - 0.1, PosLinea, PosCampo(4) + 0.5, PosLinea + 0.4, Blanco_Claro, "BF"
               For I = 4 To 15
                   If .fields(I - 2) <> 0 Then ValorCampo = Format(.fields(I - 2), "#,#0.00") Else ValorCampo = " "
                   If Not PorValor And Len(ValorCampo) > 1 Then ValorCampo = "Pag."
                   cPrint.printTexto PosCampo(I), PosLinea, ValorCampo
               Next I
               Cont = Cont + 1
               PosLinea = PosLinea + 0.4
               cPrint.printLinea 2.3, PosLinea, cPrint.dAnchoPapel - 1.1, PosLinea
               PosLinea = PosLinea + 0.1
              .MoveNext
            Loop
            cPrint.printLinea 2.3, PrimeraLinea, 2.3, PosLinea - 0.1
            cPrint.printLinea 20.4, PrimeraLinea, 20.4, PosLinea - 0.1
            For I = 2 To 15
                cPrint.printLinea PosCampo(I) - 0.05, PrimeraLinea, PosCampo(I) - 0.05, PosLinea - 0.1
            Next I
            cPrint.printTexto PosCampo(1), PosLinea, "OBSERVACIONES:"
            PosLinea = PosLinea + 0.3
            For I = 1 To 4
                PosLinea = PosLinea + 0.5
                cPrint.printLinea 2.3, PosLinea, cPrint.dAnchoPapel - 1.1, PosLinea
                PosLinea = PosLinea + 0.1
            Next I
           'fin del documento
            cPrint.finalizaImpresion
        End If
       End With
   End If
   DGQuery.Visible = True
End Sub

Public Sub Tipo_Rango_Grupos()
  If CheqRangos.value <> 0 Then
     Codigo1 = DCGrupoI
     Codigo2 = DCGrupoF
  Else
     If PorGrupo Or PorDireccion Then
        Codigo1 = DCCliente.Text
        Codigo2 = DCCliente.Text
     Else
        Codigo1 = "Todos"
        Codigo2 = "Todos"
     End If
  End If
  If Codigo1 = "" Then Codigo1 = Ninguno
  If Codigo2 = "" Then Codigo2 = Ninguno
End Sub

Public Sub Listar_Grupo()
   
   If PorDireccion Then
      sSQL = "SELECT Direccion " _
           & "FROM Clientes " _
           & "WHERE FA <> " & Val(adFalse) & " "
      If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
      sSQL = sSQL _
           & "GROUP BY Direccion " _
           & "ORDER BY Direccion "
      SelectDB_Combo DCCliente, AdoCliente, sSQL, "Direccion", True
   Else
      sSQL = "SELECT Grupo " _
           & "FROM Clientes " _
           & "WHERE FA <> " & Val(adFalse) & " "
      If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
      sSQL = sSQL _
           & "GROUP BY Grupo " _
           & "ORDER BY Grupo "
      SelectDB_Combo DCCliente, AdoCliente, sSQL, "Grupo", True
   End If
   'MsgBox sSQL
End Sub

'''Public Sub ProcGrabarMult()
'''Dim Periodo_Facturacion As String
'''Dim Total_IVAFM As Currency
''' 'Seteamos los encabezados para las facturas
'''  Validar_Porc_IVA MBFechaI
'''  FA.Porc_IVA = Porc_IVA
'''  NoMes = Month(MBFechaI)
'''  Periodo_Facturacion = CStr(Year(MBFechaI))
'''  DGQuery.Visible = False
'''  'DGQuery1.Visible = False
'''  sSQL = "SELECT C.Grupo,C.Cliente,C.Codigo,CF.Periodo,CF.Num_Mes,SUM(CF.Valor) As TValor " _
'''       & "FROM Clientes As C, Clientes_Facturacion As CF " _
'''       & "WHERE CF.Item = '" & NumEmpresa & "' " _
'''       & "AND C.T = 'N' "
'''  If PorGrupo <> 0 Then
'''     sSQL = sSQL & "AND C.Grupo = '" & DCCliente.Text & "' "
'''  Else
'''     If CheqRangos.value <> 0 Then sSQL = sSQL & "AND C.Grupo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
'''  End If
'''  If CheqFA.value = 0 Then
'''     sSQL = sSQL _
'''          & "AND CF.Num_Mes = " & NoMes & " " _
'''          & "AND CF.Periodo = '" & Periodo_Facturacion & "' "
'''  Else
'''     sSQL = sSQL & "AND CF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
'''  End If
'''  sSQL = sSQL & "AND C.Codigo = CF.Codigo " _
'''       & "GROUP BY C.Grupo,C.Cliente,C.Codigo,CF.Periodo,CF.Num_Mes " _
'''       & "ORDER BY C.Grupo,C.Cliente,CF.Periodo,CF.Num_Mes "   'CF.Periodo
'''  Select_Adodc AdoQuery, sSQL
'''  Contador = 0
''' 'MsgBox AdoQuery.Recordset.RecordCount
'''  If AdoQuery.Recordset.RecordCount > 0 Then
'''     RatonReloj
'''     Moneda_US = False
'''     TextoProc = Ninguno
'''     TextoFormaPago = PagoCred
'''     If CheqFA.value = 0 Then FechaTexto = MBFechaI Else FechaTexto = MBFecha
'''     FA.T = Pendiente
'''     FA.Nuevo_Doc = True
'''     FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
'''     FA.Fecha = FechaTexto
'''     Factura_No = FA.Factura
'''    'Grabamos Facturacion Multiple
'''     If AdoQuery.Recordset.RecordCount > 0 Then
'''        Factura_Desde = Factura_No
'''        Factura_Hasta = Factura_No + AdoQuery.Recordset.RecordCount
'''     End If
'''     'MsgBox Factura_Desde & " - " & Factura_Hasta
'''     sSQL = "DELETE * " _
'''          & "FROM Detalle_Factura " _
'''          & "WHERE Factura BETWEEN " & Factura_Desde & " and " & Factura_Hasta & " " _
'''          & "AND Periodo = '" & Periodo_Contable & "' " _
'''          & "AND Item = '" & NumEmpresa & "' " _
'''          & "AND TC = '" & FA.TC & "' " _
'''          & "AND Serie = '" & FA.Serie & "' " _
'''          & "AND Autorizacion = '" & FA.Autorizacion & "' "
'''     Ejecutar_SQL_SP sSQL
'''
'''     sSQL = "DELETE * " _
'''          & "FROM Facturas " _
'''          & "WHERE Factura BETWEEN " & Factura_Desde & " and " & Factura_Hasta & " " _
'''          & "AND Periodo = '" & Periodo_Contable & "' " _
'''          & "AND Item = '" & NumEmpresa & "' " _
'''          & "AND TC = '" & FA.TC & "' " _
'''          & "AND Serie = '" & FA.Serie & "' " _
'''          & "AND Autorizacion = '" & FA.Autorizacion & "' "
'''     Ejecutar_SQL_SP sSQL
'''
'''    'Grabamos el numero de factura
'''     Do While Not AdoQuery.Recordset.EOF
'''        FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, True)
'''        FA.CodigoC = AdoQuery.Recordset.fields("Codigo")
'''        FA.Cliente = AdoQuery.Recordset.fields("Cliente")
'''        FA.Grupo = AdoQuery.Recordset.fields("Grupo")
'''        NoMes = AdoQuery.Recordset.fields("Num_Mes")
'''        MiMes = MesesLetras(NoMes)
'''        Periodo_Facturacion = AdoQuery.Recordset.fields("Periodo")
'''        FA.EmailC = Ninguno
'''        FA.Fecha = FechaTexto
'''        FA.Nota = "Facturas del mes de " & MesesLetras(FechaMes(FechaTexto))
'''       'NoMes = Month(FechaTexto)
'''        ListarGrupos.Caption = "(" & Format$(Contador / AdoQuery.Recordset.RecordCount, "00%") & ") - " & FA.Cliente
'''        sSQL = "DELETE * " _
'''             & "FROM Asiento_F " _
'''             & "WHERE Item = '" & NumEmpresa & "' " _
'''             & "AND CodigoU = '" & CodigoUsuario & "' "
'''        Ejecutar_SQL_SP sSQL
'''
'''        sSQL = "SELECT C.Cliente,CF.Codigo_Inv,CP.Producto,CF.Valor,CF.Descuento,CF.Descuento2,CP.IVA,C.Codigo,C.Grupo " _
'''             & "FROM Clientes_Facturacion As CF,Clientes As C,Catalogo_Productos CP " _
'''             & "WHERE CF.Item = '" & NumEmpresa & "' " _
'''             & "AND CP.Periodo = '" & Periodo_Contable & "' " _
'''             & "AND C.Codigo = '" & FA.CodigoC & "' " _
'''             & "AND CF.Num_Mes = " & NoMes & " " _
'''             & "AND CF.Periodo = '" & Periodo_Facturacion & "' " _
'''             & "AND CF.Codigo_Inv = CP.Codigo_Inv " _
'''             & "AND CF.Codigo = C.Codigo " _
'''             & "AND CF.Item = CP.Item " _
'''             & "ORDER BY CF.Codigo_Inv "
'''        Select_Adodc AdoParte, sSQL
'''
'''       'MsgBox AdoParte.Recordset.RecordCount
'''
'''        With AdoParte.Recordset
'''         If .RecordCount > 0 Then
'''             Do While Not .EOF
'''                If .fields("Valor") > 0 Then
'''                   'MsgBox MiMes & vbCrLf & NoMes & vbCrLf & Periodo_Facturacion
'''                    SetAdoAddNew "Asiento_F"
'''                    SetAdoFields "CODIGO", .fields("Codigo_Inv")
'''                    SetAdoFields "CODIGO_L", CodigoL
'''                    SetAdoFields "PRODUCTO", .fields("Producto")
'''                    SetAdoFields "CANT", 1
'''                    SetAdoFields "PRECIO", .fields("Valor")
'''                    SetAdoFields "Total_Desc", .fields("Descuento")
'''                    SetAdoFields "Total_Desc2", .fields("Descuento2")
'''                    SetAdoFields "TOTAL", .fields("Valor")
'''                    If .fields("IVA") Then
'''                        Total_IVAFM = Redondear(.fields("Valor") * Porc_IVA, 2)
'''                    Else
'''                        Total_IVAFM = 0
'''                    End If
'''                    SetAdoFields "Total_IVA", Total_IVAFM
'''                    SetAdoFields "Cta", Cta_Ventas
'''                    SetAdoFields "Item", NumEmpresa
'''                    SetAdoFields "Codigo_Cliente", FA.CodigoC
'''                   'SetAdoFields "RUTA", MidStrg("(" & FA.Grupo & ") " & FA.Cliente, 1, 50)
'''                    SetAdoFields "Mes", MiMes
'''                    SetAdoFields "TICKET", Periodo_Facturacion
'''                    SetAdoFields "CodigoU", CodigoUsuario
'''                    SetAdoFields "A_No", Contador
'''                    SetAdoUpdate
'''                End If
'''               .MoveNext
'''             Loop
'''             Factura_Hasta = FA.Factura
'''             FA.Tipo_PRN = "FM"
'''             Calculos_Totales_Factura FA
'''             FA.Nota = "FACTURA PENDIENTE DE PAGO"
'''             Grabar_Factura FA, False
'''            'SRI_Crear_Clave_Acceso_Facturas FA, Code39Clt1, False, False
'''
'''             TextCheqNo = TxtGrupo
'''             ListarGrupos.Caption = ListarGrupos.Caption & ", No. " & FA.Factura
'''             sSQL = "DELETE * " _
'''                  & "FROM Clientes_Facturacion " _
'''                  & "WHERE Item = '" & NumEmpresa & "' " _
'''                  & "AND Codigo = '" & FA.CodigoC & "' " _
'''                  & "AND Num_Mes = " & NoMes & " " _
'''                  & "AND Periodo = '" & Periodo_Facturacion & "' "
'''             Ejecutar_SQL_SP sSQL
'''             Control_Procesos Normal, "Grabar Factura No. " & FA.Serie & "-" & Format$(FA.Factura, "000000000")
'''         End If
'''        End With
'''        Contador = Contador + 1
'''        AdoQuery.Recordset.MoveNext
'''     Loop
'''
'''     SSTab2.Tab = 0
'''     RatonNormal
'''     Bandera = False
'''     Evaluar = True
'''     If TipoFactura = "NV" Then
'''        Cadena = "IMPRIMIR NOTAS DE VENTA" & vbCrLf & vbCrLf
'''     Else
'''        Cadena = "IMPRIMIR FACTURAS (FM)" & vbCrLf & vbCrLf
'''     End If
'''     Cadena = Cadena _
'''            & "DESDE: " & Factura_Desde & vbCrLf & vbCrLf _
'''            & "HASTA: " & Factura_Hasta & vbCrLf & vbCrLf _
'''            & "SON UN TOTAL DE: " & Format$(Factura_Hasta - Factura_Desde + 1, "#,##0") & vbCrLf & vbCrLf _
'''            & "EN EL MENU:" & vbCrLf & vbCrLf _
'''            & "ARCHIVOS" & vbCrLf _
'''            & Space(18) & "LISTAR ANULAR FACTURAS" & vbCrLf & vbCrLf _
'''            & "Opción: En Bloque"
'''     MsgBox Cadena
'''  Else
'''     MsgBox "No se puede grabar la Factura," & vbCrLf & "falta datos."
'''  End If
'''  DGQuery.Visible = True
'''  RatonNormal
'''End Sub

Public Sub Grabar_Facturas_en_Lote()
Dim Periodo_Facturacion As String
Dim Total_IVAFM As Currency

  DGQuery.Visible = False
  sSQL = "SELECT C.Grupo,C.Cliente,C.Codigo,CF.Periodo,CF.Num_Mes,SUM(CF.Valor) As TValor " _
       & "FROM Clientes As C, Clientes_Facturacion As CF " _
       & "WHERE CF.Item = '" & NumEmpresa & "' " _
       & "AND C.T = 'N' "
  If PorGrupo <> 0 Then
     sSQL = sSQL & "AND C.Grupo = '" & DCCliente.Text & "' "
  Else
     If CheqRangos.value <> 0 Then sSQL = sSQL & "AND C.Grupo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
  End If
  If CheqFA.value = 0 Then
     sSQL = sSQL _
          & "AND CF.Num_Mes = " & NoMes & " " _
          & "AND CF.Periodo = '" & Periodo_Facturacion & "' "
  Else
     sSQL = sSQL & "AND CF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  End If
  sSQL = sSQL & "AND C.Codigo = CF.Codigo " _
       & "GROUP BY C.Grupo,C.Cliente,C.Codigo,CF.Periodo,CF.Num_Mes " _
       & "ORDER BY C.Grupo,C.Cliente,CF.Periodo,CF.Num_Mes "   'CF.Periodo
  Select_Adodc AdoQuery, sSQL
  Contador = 0
 'MsgBox AdoQuery.Recordset.RecordCount
  If AdoQuery.Recordset.RecordCount > 0 Then
     RatonReloj
     
    'Seteamos los encabezados para las facturas
     Validar_Porc_IVA MBFechaI
     NoMes = Month(MBFechaI)
     Periodo_Facturacion = CStr(Year(MBFechaI))
     Factura_Desde = FA.Factura
     
    'Grabamos Facturacion Multiple
     Grabar_Facturas_x_Lotes_SP FA.Cod_CxC, DCGrupoI, DCGrupoF, MBFechaI, MBFechaF, MBFecha, NoMes, Periodo_Facturacion, FA.Tipo_Pago, TxtNota, TxtObs, PorGrupo, CheqRangos, CheqFA
     
     Factura_Hasta = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
     
     SSTab2.Tab = 0
     RatonNormal
     
     If TipoFactura = "NV" Then
        Cadena = "IMPRIMIR NOTAS DE VENTA" & vbCrLf & vbCrLf
     Else
        Cadena = "IMPRIMIR FACTURAS (FM)" & vbCrLf & vbCrLf
     End If
     Cadena = Cadena _
            & "DESDE: " & Factura_Desde & vbCrLf & vbCrLf _
            & "HASTA: " & Factura_Hasta & vbCrLf & vbCrLf _
            & "SON UN TOTAL DE: " & Format$(Factura_Hasta - Factura_Desde + 1, "#,##0") & vbCrLf & vbCrLf _
            & "EN EL MENU:" & vbCrLf & vbCrLf _
            & "ARCHIVOS" & vbCrLf _
            & Space(18) & "LISTAR ANULAR FACTURAS" & vbCrLf & vbCrLf _
            & "Opción: En Bloque"
     MsgBox Cadena
  Else
     MsgBox "No se puede grabar la Factura," & vbCrLf & "falta datos."
  End If
  DGQuery.Visible = True
  RatonNormal
End Sub


Private Sub CheqFA_Click()
  If CheqFA.value = 0 Then
     MBFecha.Visible = False
     Label2.Visible = False
     DCLinea.Visible = False
     LblFactura.Visible = False
  Else
     MBFecha.Visible = True
     Label2.Visible = True
     DCLinea.Visible = True
     LblFactura.Visible = True
     MBFecha.SetFocus
  End If
End Sub

Private Sub CheqPorRubro_Click()
 If CheqPorRubro.value = 0 Then
    DCProductos.Visible = False
 Else
    sSQL = "SELECT Codigo_Inv, Producto, Detalle, Codigo_Barra, Cta_Inventario " _
         & "FROM Catalogo_Productos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'P' " _
         & "AND LEN(Cta_Inventario) = 1 " _
         & "AND INV <> " & Val(adFalse) & " " _
         & "ORDER BY Producto "
    SelectDB_Combo DCProductos, AdoProductos, sSQL, "Producto"
    DCProductos.Visible = True
 End If
End Sub

Private Sub CheqRangos_Click()
 If CheqRangos.value = 0 Then
    DCGrupoI.Visible = False
    DCGrupoF.Visible = False
 Else
    sSQL = "SELECT Grupo " _
         & "FROM Clientes " _
         & "WHERE FA <> " & Val(adFalse) & " "
    If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
    sSQL = sSQL & "GROUP BY Grupo " _
         & "ORDER BY Grupo "
    SelectDB_Combo DCGrupoI, AdoGrupo, sSQL, "Grupo"
    SelectDB_Combo DCGrupoF, AdoGrupo, sSQL, "Grupo", True
 
    DCGrupoI.Visible = True
    DCGrupoF.Visible = True
 End If
End Sub

Private Sub Command1_Click()
  TxtAyuda.Visible = True
  TxtAyuda.SetFocus
End Sub

Public Sub Generar_Facturas_Grupos()
 Encerar_Factura FA
 FA.Tipo_Pago = SinEspaciosIzq(DCTipoPago)
 If Len(FA.Tipo_Pago) = 1 Then
    MsgBox "NO HA SELECCIONADO LA FORMA DE PAGO"
 Else
    FA.Cod_CxC = DCLinea.Text
    Lineas_De_CxC FA
    If CFechaLong(MBFechaI) > CFechaLong(FA.Vencimiento) Then
       MsgBox "No se puede General Facturas, porque la autorizacion ya esta caducada"
    Else
        If CTipoConsulta.Text = "Listar Todos" Then
          'FA.Factura = Numero_Factura(FA)
           FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
           Mensajes = "Esta Seguro de grabar desde " & vbCrLf & vbCrLf
           If TipoFactura = "NV" Then
              Mensajes = Mensajes & "La Nota de Venta No. " & FA.Serie & "-" & FA.Factura
           Else
              Mensajes = Mensajes & "La Factura No. " & FA.Serie & "-" & FA.Factura
           End If
           Mensajes = Mensajes & " en bloque "
           Titulo = "Formulario de Grabacion"
           If BoxMensaje = vbYes Then Grabar_Facturas_en_Lote  'ProcGrabarMult
        Else
           MsgBox "Debe seleccionar la opcion: 'Listar Todos' " & vbCrLf _
                & "Caso Contrario no podra facturar"
        End If
    End If
 End If
 CheqRangos.SetFocus
End Sub

Public Sub Listado_x_Grupos()
  sSQL = "SELECT Grupo,Direccion,COUNT(Grupo) As Alumnos " _
       & "FROM Clientes " _
       & "WHERE FA <> 0 "
  If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
  sSQL = sSQL _
       & "GROUP BY Grupo,Direccion " _
       & "ORDER BY Grupo,Direccion "
  Select_Adodc AdoNiveles, sSQL
  DGQuery.Caption = "RESUMEN DE GRUPO PARA FACTURAR"
  MensajeEncabData = DGQuery.Caption
  ImprimirAdodc AdoNiveles, 1, 10
  RatonNormal
End Sub

Public Sub Generar_Eliminar_Rubros()
Si_No = False
Codigo4 = Format$(Year(MBFechaI), "0000")
If PorGrupo Then
   FPensiones.Show 1
Else
   MsgBox "Debe estar con Visto: " & vbCrLf _
        & "Por Grupo" & vbCrLf _
        & "Caso Contrario no podra facturar"
End If
End Sub

Public Sub Generar_Deuda_Pendiente()
Si_No = True
If PorGrupo Then
   FPensiones.Show 1
Else
   MsgBox "Debe estar con Visto: " & vbCrLf _
        & "Por Grupo" & vbCrLf _
        & "Caso Contrario no podra facturar"
End If
End Sub

Public Sub Imprimir_Recibos_Cobros()
 'Control_Procesos  "I", "Reimpresion de Facturas desde la " & Factura_Desde & " a la " & Factura_Hasta
  sSQL = "SELECT SUM(Valor) As SaldoPend, Codigo " _
       & "FROM Clientes_Facturacion  " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' "
  If CheqRangos.value <> 0 Then sSQL = sSQL & "AND GrupoNo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
  sSQL = sSQL & "GROUP BY Codigo "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     Imprimir_Recibos_CxC_PreFA ListarGrupos, AdoAux, AdoAux2, MBFechaI, MBFechaF, Codigo1, Codigo2, FA
  Else
    MsgBox "No se puede imprimir el rando de Recibos"
  End If
End Sub

Public Sub Recalcular_Fechas()
    Mensajes = "Recalcular Meses de Cobros"
    Titulo = "Formulario de Recalculación"
    If BoxMensaje = vbYes Then
       RatonReloj
       sSQL = "SELECT Periodo " _
            & "FROM Clientes_Facturacion " _
            & "WHERE ISNUMERIC(Periodo) <> 0 " _
            & "AND Item = '" & NumEmpresa & "' " _
            & "GROUP BY Periodo " _
            & "ORDER BY Periodo "
       Select_Adodc AdoAux, sSQL
       With AdoAux.Recordset
        If .RecordCount > 0 Then
            Do While Not .EOF
               Anio = .fields("Periodo")
               For IMes = 1 To 12
                   Mifecha = BuscarFecha(UltimoDiaMes("01/" & Format$(IMes, "00") & "/" & Anio))
                   sSQL = "UPDATE Clientes_Facturacion " _
                        & "SET Fecha = #" & Mifecha & "# " _
                        & "WHERE Item = '" & NumEmpresa & "' " _
                        & "AND Periodo = '" & Anio & "' " _
                        & "AND Num_Mes = " & IMes & " "
                  'MsgBox sSQL
                   Ejecutar_SQL_SP sSQL
               Next IMes
              .MoveNext
            Loop
        End If
       End With
       RatonNormal
    End If
End Sub

Private Sub Command2_Click()
    Unload ListarGrupos
End Sub

Private Sub Command3_Click()
    NombreArchivo = SelectDialogFile(, "Todos los archivos|*.*")
    RutaGeneraFile = NombreArchivo
    If NombreArchivo <> "" Then
       LblArchivo.Caption = RutaGeneraFile
    Else
       LblArchivo.Caption = ""
    End If
End Sub

Private Sub Command5_Click()
Dim Si_Envia As Boolean
Dim Codigo_Banco As String
Dim CadDeuda As String
Dim NombreRepresentante As String
Dim Curso As String
Dim IdMail As Long

  DGQuery.Visible = False
  If CheqConDeuda.value Then
     sSQL = "UPDATE Reporte_CxC_Cuotas " _
          & "SET E = 0 " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
  Else
     sSQL = "UPDATE Clientes " _
          & "SET FactM = 0 " _
          & "WHERE FA <> " & Val(adFalse) & " "
  End If
  Ejecutar_SQL_SP sSQL
  
  For IdMail = 0 To LstClientes.ListCount - 1
    If LstClientes.Selected(IdMail) Then
       NombreCliente = TrimStrg(MidStrg(LstClientes.List(IdMail), 1, 79))
       If CheqConDeuda.value Then
          sSQL = "UPDATE Reporte_CxC_Cuotas " _
               & "SET E = 1 " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND CodigoU = '" & CodigoUsuario & "' " _
               & "AND Cliente = '" & NombreCliente & "' "
       Else
          sSQL = "UPDATE Clientes " _
               & "SET FactM = 1 " _
               & "WHERE FA <> " & Val(adFalse) & " " _
               & "AND Cliente LIKE '" & NombreCliente & "%' "
       End If
       Ejecutar_SQL_SP sSQL
    End If
  Next IdMail
  If CheqConDeuda.value Then
     sSQL = "SELECT " & ListaDeCampos & ", C.Representante, C.CI_RUC, C.Email, C.EmailR " _
          & "FROM Reporte_CxC_Cuotas As RCC, Clientes As C " _
          & "WHERE RCC.Item = '" & NumEmpresa & "' " _
          & "AND RCC.CodigoU = '" & CodigoUsuario & "' " _
          & "AND RCC.E <> 0 " _
          & "AND RCC.Codigo = C.Codigo " _
          & "ORDER BY RCC.GrupoNo, RCC.Cliente "
  Else
     sSQL = "SELECT Cliente, 0 As CxC_20XX, 0 As SubTotal, 0 As Anticipos, 0 As Total, Direccion As Detalle_Grupo, Grupo As GrupoNo, Representante, CI_RUC, Email, EmailR " _
          & "FROM Clientes " _
          & "WHERE FA <> " & Val(adFalse) & " " _
          & "AND FactM <> " & Val(adFalse) & " " _
          & "ORDER BY GrupoNo, Cliente "
  End If
  Select_Adodc AdoAux, sSQL, , , "CxC en Grupo"
  'DGQuery1.Visible = False
  TMail.ListaMail = 255
  TMail.ListaError = ""
  TMail.para = ""
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          ComunicadoEntidad = ""
          Mifecha = FechaStrg(FechaSistema)
          Total = .fields("Total")
          Curso = .fields("Detalle_Grupo")
          TBeneficiario.Cliente = .fields("Cliente")
          TBeneficiario.Representante = .fields("Representante")
          TBeneficiario.Grupo_No = .fields("GrupoNo")
          TBeneficiario.CI_RUC = .fields("CI_RUC")
          
          TMail.Mensaje = "Estimado/a "
          If Len(TBeneficiario.Representante) > 1 Then TMail.Mensaje = TMail.Mensaje & TBeneficiario.Representante Else TMail.Mensaje = TMail.Mensaje & TBeneficiario.Cliente
          If Len(TBeneficiario.Representante) > 1 And TBeneficiario.Representante <> TBeneficiario.Cliente Then TMail.Mensaje = TMail.Mensaje & ", Representante de: " & TBeneficiario.Cliente
          
          If CheqConDeuda.value <> 0 Then
             TMail.Asunto = "Envio automatizado de su cartera pendiente por USD " & Format$(Total, "#,#0.00")
             
             html_Detalle_adicional = "<thead><tr><strong><td>DETALLE</td><td>PENDIENTE</td><td>MONTO</td></strong></tr></thead><tbody>"
             For J = 2 To .fields.Count - 7
                 If .fields(J) > 0 Then
                     Select Case .fields(J).Name
                       Case "SubTotal", "Total", "Anticipos"
                            'No hace nada
                       Case Else
                            html_Detalle_adicional = html_Detalle_adicional _
                                                   & "<tr>" _
                                                   & "<td>Valor Pendiente</td>" _
                                                   & "<td>" & .fields(J).Name & "</td>" _
                                                   & "<td class='row text-right'>" & Format(.fields(J), "#,#0.00") & "</td>" _
                                                   & "</tr>"
                     End Select
                 End If
             Next J
             
             html_Detalle_adicional = html_Detalle_adicional _
                                    & "<tr>" _
                                    & "<td colspan='2'><N>TOTAL PENDIENTE POR CANCELAR</N></td>" _
                                    & "<td class='row text-right'><N>" & Format$(Total, "#,#0.00") & "</N></td>" _
                                    & "</tr></tbody>"

             TMail.Mensaje = TMail.Mensaje _
                           & ", Le escribimos para recordarle que su saldo actual es USD " & Format$(Total, "#,#0.00") & ". Pensando " _
                           & "en nuevas maneras de brindarle un mejor servicio, detallamos a continuación el estado de cuenta digital " _
                           & "con corte al " & Mifecha & ". En caso de tener inconformidad, comuniquese con atencion al Cliente." & vbCrLf & vbCrLf _
                           & "<N>Su codigo de Referencia es:</N> " & TBeneficiario.CI_RUC & "." & vbCrLf _
                           & "<N>Usted pertenece al grupo:</N> " & TBeneficiario.Grupo_No & " - " & Curso & "." _
                           & "<table BORDER CELLPADDING=10 CELLSPACING=0 class='content-table'>" & html_Detalle_adicional & "</table>" _
                           & "Necesita asistencia para realizar su pago? Estamos aquí para ayudarle, cualquier consulta comuniquese al teléfono: " & Telefono1 & "." & vbCrLf
          Else
             If Len(TxtAsunto) > 1 Then TMail.Asunto = TxtAsunto Else TMail.Asunto = "Envio automatizado para Usted."
             TMail.Mensaje = TMail.Mensaje & vbCrLf & "Ubicacion: " & TBeneficiario.Grupo_No & " - " & Curso & "." & vbCrLf
          End If
          If Len(TxtMensaje.Text) > 1 Then TMail.Mensaje = TMail.Mensaje & TxtMensaje.Text & "." & vbCrLf
          If CheqConDeuda.value <> 0 Then TMail.Mensaje = TMail.Mensaje & "<N>NOTA:</N> Si realizo su pago con transferencia bancaria, considere que el pago se verá reflejado en un maximo de hasta 72 horas. "
          TMail.Mensaje = TMail.Mensaje & vbCrLf & "<N>SERVIRLE ES NUESTRO OBJETIVO, DISFRUTARLO EL SUYO.</N>"
          
          TMail.TipoDeEnvio = "CE"
          Email_CE_Copia = True
          TMail.Destinatario = TBeneficiario.Representante
          If Len(LblArchivo.Caption) > 1 Then TMail.Adjunto = LblArchivo.Caption & ";" Else TMail.Adjunto = ""
          TMail.para = ""
          Insertar_Mail TMail.para, .fields("EmailR")
          Insertar_Mail TMail.para, .fields("Email")
          FEnviarCorreos.Show 1
          TMail.Mensaje = ""
          TMail.MensajeHTML = ""
         .MoveNext
       Loop
    End If
  End With
  DGQuery.Visible = True
  MsgBox "Proceso Terminado con exito."
  If Len(TMail.ListaError) > 1 Then
     MsgBox "Revice en su correo los errores "
     TMail.para = Lista_De_Correos(0).Correo_Electronico
     TMail.Asunto = "CORREOS CON ERRORES"
     TMail.Mensaje = TMail.ListaError
     FEnviarCorreos.Show 1
  End If
End Sub

Private Sub CTipoConsulta_GotFocus()
   PorGrupo = False
   PorDireccion = False
End Sub

Private Sub CTipoConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub CTipoConsulta_LostFocus()
   If Len(CTipoConsulta.Text) <= 6 Then CTipoConsulta.Text = "Listar por Grupo"
   Select Case CTipoConsulta.Text
     Case "Listar por Grupo"
          PorGrupo = True
          Listar_Grupo
          DCCliente.Visible = True
     Case "Listar por Direccion"
          PorDireccion = True
          Listar_Grupo
          DCCliente.Visible = True
     Case Else
          DCCliente.Visible = False
          SSTab2.Tab = 0
   End Select
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_LostFocus()
  If CargarFrom Then
    'MsgBox "Desktop Dest: Una Vez"
     Actualizar_Datos_Representantes_SP Mas_Grupos
     Insertar_Clientes_Auxiliar_SP
     CargarFrom = False
  End If

  Listar_Clientes_Grupo
  SSTab2.Tab = 0
  
End Sub

Private Sub DCGrupoI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCGrupoI_LostFocus()
  Codigo1 = DCGrupoI
End Sub

Private Sub DCGrupoF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCGrupoF_LostFocus()
  Codigo2 = DCGrupoF
End Sub

Private Sub DCLinea_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCLinea_LostFocus()
  LogoFactura = Ninguno
  AltoFactura = 0
  AnchoFactura = 0
  CodigoL = Ninguno
  EspacioFactura = 0
  Pos_Factura = 0
  Cta_Cobrar = Ninguno
  Cta_Ventas = Ninguno
  FA.Cod_CxC = DCLinea.Text
  Lineas_De_CxC FA
  LblFactura.Caption = " No. " & FA.Serie & "-" & Format(ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False), "000000000")
  If FA.Cta_CxP <> Ninguno Then DCLinea.Visible = True Else DCLinea.Visible = False
End Sub

Private Sub DCProductos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCProductos_LostFocus()
  Listar_Deuda_por_Api
End Sub

'''Private Sub DGParte_KeyDown(KeyCode As Integer, Shift As Integer)
'''  Keys_Especiales Shift
'''  If AdoParte.Recordset.RecordCount > 0 Then
'''     If CtrlDown And KeyCode = vbKeyDelete Then
'''        Codigo = AdoParte.Recordset.Fields("Codigo")
'''        CodigoP = AdoParte.Recordset.Fields("Codigo_Inv")
'''        Mensajes = "Esta seguro de Borrar Codigo " & CodigoP & " de " & Codigo
'''        Titulo = "Pregunta de Eliminacion"
'''        If BoxMensaje = vbYes Then
'''           sSQL = "DELETE * " _
'''                & "FROM Clientes_Facturacion " _
'''                & "WHERE Codigo_Inv = '" & CodigoP & "' " _
'''                & "AND Codigo = '" & Codigo & "' " _
'''                & "AND Num_Mes = 0 " _
'''                & "AND Item = '" & NumEmpresa & "' "
'''           Ejecutar_SQL_SP sSQL
'''           'TipoConsultaCuotas
'''           TipoConsultaCxC
'''        End If
'''     End If
'''     If CtrlDown And KeyCode = vbKeyV Then
'''        Codigo = AdoParte.Recordset.Fields("Codigo")
'''        CodigoP = AdoParte.Recordset.Fields("Codigo_Inv")
'''        Cadena = AdoParte.Recordset.Fields("Cliente") & vbCrLf & vbCrLf _
'''               & "PRODUCTO: " & AdoParte.Recordset.Fields("Producto") & vbCrLf & vbCrLf _
'''               & "NUEVO VALOR:"
'''        Valor = CCur(Val(InputBox(Cadena, "CAMBIO DE VALORES", Format$(AdoParte.Recordset.Fields("Valor"), "#,##0.00"))))
'''        If Valor >= 0 Then
'''           sSQL = "UPDATE Clientes_Facturacion " _
'''                & "SET Valor = " & Valor & " " _
'''                & "WHERE Codigo_Inv = '" & CodigoP & "' " _
'''                & "AND Codigo = '" & Codigo & "' " _
'''                & "AND Num_Mes = 0 " _
'''                & "AND Item = '" & NumEmpresa & "' "
'''           Ejecutar_SQL_SP sSQL
'''          'TipoConsultaCuotas
'''           TipoConsultaCxC
'''        End If
'''     End If
'''     If CtrlDown And KeyCode = vbKeyG Then
'''        CodigoP = AdoParte.Recordset.Fields("Codigo_Inv")
'''        Codigo1 = DCCliente.Text
'''        Cadena = "CAMBIO DE VALORES DEL GRUPO: " & Codigo1 & vbCrLf & vbCrLf _
'''               & "PRODUCTO: " & AdoParte.Recordset.Fields("Producto") & vbCrLf & vbCrLf _
'''               & "NUEVO VALOR DEL CODIGO: (" & CodigoP & "):"
'''        Valor = CCur(Val(InputBox(Cadena, "CAMBIO DE VALORES", Format$(AdoParte.Recordset.Fields("Valor"), "#,##0.00"))))
'''        DGParte.Visible = False
'''        If Valor >= 0 Then
'''           If SQL_Server Then
'''              sSQL = "UPDATE Clientes_Facturacion " _
'''                   & "SET Valor = " & Valor & " " _
'''                   & "FROM Clientes_Facturacion As CF,Clientes As C "
'''           Else
'''              sSQL = "UPDATE Clientes_Facturacion As CF,Clientes As C " _
'''                   & "SET CF.Valor = " & Valor & " "
'''           End If
'''           sSQL = sSQL & "WHERE CF.Codigo = C.Codigo " _
'''                & "AND C.Grupo = '" & Codigo1 & "' " _
'''                & "AND CF.Item = '" & NumEmpresa & "' " _
'''                & "AND CF.Num_Mes = 0 " _
'''                & "AND CF.Codigo_Inv = '" & CodigoP & "' "
'''           Ejecutar_SQL_SP sSQL
'''          'TipoConsultaCuotas
'''           TipoConsultaCxC
'''           DCCliente.SetFocus
'''        End If
'''        DGParte.Visible = True
'''     End If
'''     If CtrlDown And KeyCode = vbKeyB Then Buscar_Datos DGParte, AdoParte
'''  End If
'''End Sub

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
Dim IMes As Byte
  Keys_Especiales Shift
  
  If AdoQuery.Recordset.RecordCount > 0 Then
     If SSTab2.Tab = 0 Then
     'If CtrlDown And KeyCode = vbKeyB Then Buscar_Datos DGQuery, AdoQuery
     If CtrlDown And KeyCode = vbKeyInsert Then
        
           CodigoCliente = AdoQuery.Recordset.fields("Codigo")
           Codigo = AdoQuery.Recordset.fields("Codigo")
           Codigo1 = AdoQuery.Recordset.fields("Cliente")
           Codigo2 = AdoQuery.Recordset.fields("Grupo")
           Codigo4 = Format$(Year(MBFechaI), "0000")
           FAsignaFact.Show 1
        'TipoConsultaCuotas
     End If
     If CtrlDown And KeyCode = vbKeyD Then
        Codigo1 = AdoQuery.Recordset.fields("Grupo")
        Cadena = "CAMBIO DE VALORES DEL GRUPO: " & Codigo1 & vbCrLf & vbCrLf _
               & "UBICACION: " & AdoQuery.Recordset.fields("Direccion") & vbCrLf & vbCrLf _
               & "NUEVO VALOR DEL CODIGO: (" & CodigoP & "):"
        Codigo2 = UCaseStrg(InputBox(Cadena, "CAMBIO DE VALORES", AdoQuery.Recordset.fields("Direccion")))
        DGQuery.Visible = False
        If Len(Codigo2) > 1 Then
           sSQL = "UPDATE Clientes " _
                & "SET Direccion = '" & Codigo2 & "' " _
                & "WHERE Grupo = '" & Codigo1 & "' "
           If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
           Ejecutar_SQL_SP sSQL
           SSTab2.Tab = 0
           DCCliente.SetFocus
        End If
        DGQuery.Visible = True
     End If
     If CtrlDown And KeyCode = vbKeyG Then
        Codigo1 = AdoQuery.Recordset.fields("Grupo")
        Cadena = "CAMBIO DE VALORES DEL GRUPO: " & Codigo1 & vbCrLf & vbCrLf _
               & "GRUPO: " & AdoQuery.Recordset.fields("Grupo") & vbCrLf & vbCrLf _
               & "NUEVO VALOR DEL CODIGO: (" & CodigoP & "):"
        Codigo2 = UCaseStrg(InputBox(Cadena, "CAMBIO DE VALORES", AdoQuery.Recordset.fields("Grupo")))
        DGQuery.Visible = False
        If Len(Codigo2) > 1 Then
           sSQL = "UPDATE Clientes " _
                & "SET Grupo = '" & Codigo2 & "' " _
                & "WHERE Grupo = '" & Codigo1 & "' "
           If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
           Ejecutar_SQL_SP sSQL
           SSTab2.Tab = 0
           DCCliente.SetFocus
        End If
        DGQuery.Visible = True
     End If
     If CtrlDown And KeyCode = vbKeyN Then
        DGQuery.Visible = False
        Codigo1 = AdoQuery.Recordset.fields("Grupo")
        CodigoCliente = AdoQuery.Recordset.fields("Codigo")
        DCCliente.SetFocus
        Mensajes = "Desactivar este grupo"
        Titulo = "Formulario de Activacion"
        If BoxMensaje = vbYes Then
          'Facturacion Mensual
           If SQL_Server Then
              sSQL = "UPDATE Clientes_Facturacion " _
                   & "SET GrupoNo = C.Grupo " _
                   & "FROM Clientes_Facturacion As CF,Clientes As C "
           Else
              sSQL = "UPDATE Clientes_Facturacion As CF,Clientes As C " _
                   & "SET CF.GrupoNo = C.Grupo "
           End If
           sSQL = sSQL & "WHERE CF.Item = '" & NumEmpresa & "' "
           If Mas_Grupos Then sSQL = sSQL & "AND C.DirNumero = '" & NumEmpresa & "' "
           sSQL = sSQL & "AND CF.Codigo = C.Codigo "
           Ejecutar_SQL_SP sSQL
          'Educativo
           If SQL_Server Then
              sSQL = "UPDATE Trans_Notas " _
                   & "SET CodE = C.Grupo " _
                   & "FROM Trans_Notas As CF,Clientes As C "
           Else
              sSQL = "UPDATE Trans_Notas As CF,Clientes As C " _
                   & "SET CF.CodE = C.Grupo "
           End If
           sSQL = sSQL & "WHERE CF.Item = '" & NumEmpresa & "' " _
                & "AND CF.Periodo = '" & Periodo_Contable & "' " _
                & "AND CF.Codigo = C.Codigo "
           Ejecutar_SQL_SP sSQL
           If SQL_Server Then
              sSQL = "UPDATE Trans_Asistencia " _
                   & "SET CodE = C.Grupo " _
                   & "FROM Trans_Asistencia As CF,Clientes As C "
           Else
              sSQL = "UPDATE Trans_Asistencia As CF,Clientes As C " _
                   & "SET CF.CodE = C.Grupo "
           End If
           sSQL = sSQL & "WHERE CF.Item = '" & NumEmpresa & "' " _
                & "AND CF.Periodo = '" & Periodo_Contable & "' " _
                & "AND CF.Codigo = C.Codigo "
           Ejecutar_SQL_SP sSQL
           If SQL_Server Then
              sSQL = "UPDATE Trans_Notas_Auxiliares " _
                   & "SET CodE = C.Grupo " _
                   & "FROM Trans_Notas_Auxiliares As CF,Clientes As C "
           Else
              sSQL = "UPDATE Trans_Notas_Auxiliares As CF,Clientes As C " _
                   & "SET CF.CodE = C.Grupo "
           End If
           sSQL = sSQL & "WHERE CF.Item = '" & NumEmpresa & "' " _
                & "AND CF.Periodo = '" & Periodo_Contable & "' " _
                & "AND CF.Codigo = C.Codigo "
           Ejecutar_SQL_SP sSQL
           If SQL_Server Then
              sSQL = "UPDATE Trans_Notas_Grado " _
                   & "SET CodE = C.Grupo " _
                   & "FROM Trans_Notas_Grado As CF,Clientes As C "
           Else
              sSQL = "UPDATE Trans_Notas_Grado As CF,Clientes As C " _
                   & "SET CF.CodE = C.Grupo "
           End If
           sSQL = sSQL & "WHERE CF.Item = '" & NumEmpresa & "' " _
                & "AND CF.Periodo = '" & Periodo_Contable & "' " _
                & "AND CF.Codigo = C.Codigo "
           Ejecutar_SQL_SP sSQL
           If SQL_Server Then
              sSQL = "UPDATE Trans_Actas " _
                   & "SET CodE = C.Grupo " _
                   & "FROM Trans_Actas As CF,Clientes As C "
           Else
              sSQL = "UPDATE Trans_Actas As CF,Clientes As C " _
                   & "SET CF.CodE = C.Grupo "
           End If
           sSQL = sSQL & "WHERE CF.Item = '" & NumEmpresa & "' " _
                & "AND CF.Periodo = '" & Periodo_Contable & "' " _
                & "AND CF.Codigo = C.Codigo "
           Ejecutar_SQL_SP sSQL
           
          'Borrar Alumnos de Grupo de Notas y Facturacion Mensual
           sSQL = "DELETE * " _
                & "FROM Clientes_Facturacion " _
                & "WHERE GrupoNo = '" & Codigo1 & "' " _
                & "AND Item = '" & NumEmpresa & "' "
           Ejecutar_SQL_SP sSQL
           sSQL = "DELETE * " _
                & "FROM Trans_Notas " _
                & "WHERE CodE = '" & Codigo1 & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' "
           Ejecutar_SQL_SP sSQL
           sSQL = "DELETE * " _
                & "FROM Trans_Asistencia " _
                & "WHERE CodE = '" & Codigo1 & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' "
           Ejecutar_SQL_SP sSQL
           sSQL = "DELETE * " _
                & "FROM Trans_Notas_Auxiliares " _
                & "WHERE CodE = '" & Codigo1 & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' "
           Ejecutar_SQL_SP sSQL
           sSQL = "DELETE * " _
                & "FROM Trans_Notas_Grado " _
                & "WHERE CodE = '" & Codigo1 & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' "
           Ejecutar_SQL_SP sSQL
           sSQL = "DELETE * " _
                & "FROM Trans_Actas " _
                & "WHERE CodE = '" & Codigo1 & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' "
           Ejecutar_SQL_SP sSQL
           
          'Deshabilitamos Alumnos/Clientes
           sSQL = "UPDATE Clientes " _
                & "SET FA = " & Val(adFalse) & " " _
                & "WHERE Grupo = '" & Codigo1 & "' "
           If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
           Ejecutar_SQL_SP sSQL
           'TipoConsultaCxC
           Listar_Grupo
        End If
        DGQuery.Visible = True
        DCCliente.SetFocus
     End If
     
     If CtrlDown And KeyCode = vbKeyF10 Then
        Mensajes = "Seguro de Eliminar Todos los Rubros de Facturacion"
        Titulo = "Formulario de Eliminacion"
        If BoxMensaje = vbYes Then
           sSQL = "DELETE * " _
                & "FROM Clientes_Facturacion " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' "
           Ejecutar_SQL_SP sSQL
        End If
     End If
     If CtrlDown And KeyCode = vbKeyM Then
        DGQuery.AllowUpdate = True
        MsgBox "Proceso Aceptado, puede Modificar"
        DGQuery.SetFocus
     End If
     If CtrlDown And KeyCode = vbKeyF6 Then DGQuery.AllowUpdate = False
     If CtrlDown And KeyCode = vbKeyP Then
        SQLMsg1 = "[" & AdoQuery.Recordset.RecordCount & "] " & DGQuery.Caption
        DGQuery.Visible = False
        Cuadricula = True
        ImprimirRubrosFacturaGrupo AdoQuery.Recordset.fields("Grupo"), True
        'ImprimirAdodc AdoQuery, 1, 8, True
        DGQuery.Visible = True
     End If
     If CtrlDown And KeyCode = vbKeyR Then
        DGQuery.Visible = False
        Codigo1 = AdoQuery.Recordset.fields("Grupo")
        DCCliente.SetFocus
        Mensajes = "Retirar Beneficiarios sin deuda del Grupo: " & Codigo1
        Titulo = "Formulario de Retiro"
        If BoxMensaje = vbYes Then
           sSQL = "UPDATE Clientes " _
                & "SET X = 'R' " _
                & "WHERE Codigo <> '" & Ninguno & "' "
           Ejecutar_SQL_SP sSQL
           
           If SQL_Server Then
              sSQL = "UPDATE Clientes " _
                   & "SET X = '.' " _
                   & "FROM Clientes As C,Clientes_Facturacion As CF "
           Else
              sSQL = "UPDATE Clientes As C,Clientes_Facturacion As CF " _
                   & "SET X = '.' "
           End If
           sSQL = sSQL _
                & "WHERE C.Codigo = CF.Codigo "
           Ejecutar_SQL_SP sSQL
           
          'Deshabilitamos Alumnos/Clientes
           sSQL = "UPDATE Clientes " _
                & "SET FA = " & Val(adFalse) & " " _
                & "WHERE X = 'R' " _
                & "AND Grupo = '" & Codigo1 & "' "
           Ejecutar_SQL_SP sSQL
           'TipoConsultaCxC
           Listar_Grupo
        End If
        DGQuery.Visible = True
        DCCliente.SetFocus
     End If
     If KeyCode = vbKeyF2 Then
        DGQuery.SelStart = 0
        DGQuery.SelLength = 0
     End If
  End If
  End If
End Sub

Private Sub Form_Activate()
   MBFechaI = "01/01/" & Year(FechaSistema)
   MBFechaF = CLongFecha(CFechaLong(MBFechaI) + 364)
   FechaValida MBFecha
   FechaValida MBFechaI
   FechaValida MBFechaF

   CTipoConsulta.Clear
   CTipoConsulta.AddItem "Listar por Grupo"
   CTipoConsulta.AddItem "Listar por Direccion"
   CTipoConsulta.AddItem "Listar Todos"
   CTipoConsulta.Text = "Listar por Grupo"
 
   TxtAyuda.Text = "<F1>          Genera Archivos de Texto" & vbCrLf _
                 & "<Ctrl+B>      Buscar Datos " & vbCrLf _
                 & "<Ctrl+G>      Cambia en Grupo el Valor del Grupo" & vbCrLf _
                 & "<Ctrl+D>      Cambia en Grupo el Valor de la Direccion" & vbCrLf _
                 & "<Ctrl+R>      Retirar Beneficiarios sin deuda del Grupo" & vbCrLf _
                 & "<Ctrl+Insert> Insertar Rubros" & vbCrLf _
                 & "<Ctrl+F10>    Eliminar Totdos Rubros de Facturacion" & vbCrLf _
                 & "<Ctrl+F11>    Inserta Totdos Rubros de Facturacion" & vbCrLf _
                 & "<Esc>         Salid de Ayuda." & vbCrLf
   
''   sSQL = "SELECT Curso " _
''        & "FROM Catalogo_Cursos " _
''        & "WHERE Item = '" & NumEmpresa & "' " _
''        & "AND Periodo = '" & Periodo_Contable & "' " _
''        & "ORDER BY Curso "
''   Select_Adodc AdoRepresentante, sSQL
''   If AdoRepresentante.Recordset.RecordCount > 0 Then Tiene_Cursos = True Else Tiene_Cursos = False
      
   sSQL = "SELECT (Codigo & ' ' & Descripcion) As CTipoPago " _
        & "FROM Tabla_Referenciales_SRI " _
        & "WHERE Tipo_Referencia = 'FORMA DE PAGO' " _
        & "AND Codigo IN ('01','16','17','18','19','20','21') " _
        & "ORDER BY Codigo "
   SelectDB_Combo DCTipoPago, AdoTipoPago, sSQL, "CTipoPago"
      
   FA.TC = TipoFactura
   FA.Fecha = PrimerDiaMes(FechaSistema)
         
   Label13.Caption = Lista_De_Correos(0).Correo_Electronico
   
   SSTab2.Tab = 0
   SSTab2.width = MDI_X_Max - SSTab2.Left
   
   DGQuery.width = SSTab2.width
   'DGQuery.Height = ((MDI_Y_Max - SSTab2.Top) / 2) - 800
   DGQuery.Height = MDI_Y_Max - 3300
   
   FrmEmail.width = SSTab2.width
   FrmEmail.Height = MDI_Y_Max - 3300
   
   LstClientes.Height = FrmEmail.Height - LstClientes.Top - 200
   LstClientes.width = FrmEmail.width - LstClientes.Left - FrmEmail.Left - 100
   
   Label3.Top = DGQuery.Top + DGQuery.Height + 10
   Label4.Top = DGQuery.Top + DGQuery.Height + 10
   Label8.Top = DGQuery.Top + DGQuery.Height + 10
   Label9.Top = DGQuery.Top + DGQuery.Height + 10
   Label10.Top = DGQuery.Top + DGQuery.Height + 10
   Label11.Top = DGQuery.Top + DGQuery.Height + 10
   Label1.Top = DGQuery.Top + DGQuery.Height + AdoQuery.Height + 10
   Label1.width = MDI_X_Max - 100
   AdoQuery.Top = DGQuery.Top + DGQuery.Height + 10
   AdoQuery.width = MDI_X_Max - AdoQuery.Left - 50
   
   Label1.Caption = "<Ctrl+M> Modificar | <Ctrl+F6> No Modifica |<Ctrl+Ins> Insertar | <Ctrl+B> Buscar | <Ctrl+Supr> Eliminar | <Ctrl+V> Cambio de Valores | <Ctrl+N> Desactivar Grupo "
   
   TipoFactura = "FA"
   Listar_Grupo
   RatonNormal
   ListarGrupos.WindowState = 2
  'CTipoConsulta.SetFocus
   MBFechaI.SetFocus
End Sub

Private Sub Form_Deactivate()
    ListarGrupos.WindowState = 1
End Sub

Private Sub Form_Load()
   'CentrarForm ListarSuscripciones
   ConectarAdodc AdoAux
   ConectarAdodc AdoAux2
   ConectarAdodc AdoGrupo
   ConectarAdodc AdoParte
   ConectarAdodc AdoLinea
   ConectarAdodc AdoQuery
   ConectarAdodc AdoCliente
   ConectarAdodc AdoCiudad
   ConectarAdodc AdoNiveles
   ConectarAdodc AdoTipoPago
   ConectarAdodc AdoProductos
End Sub

Private Sub LstClientes_Click()
  If LstClientes.Selected(0) = True Then
     For I = 0 To LstClientes.ListCount - 1
         LstClientes.Selected(I) = True
     Next I
  End If
End Sub

Private Sub MBFecha_GotFocus()
    MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFecha_LostFocus()
   FechaValida MBFecha
   FA.TC = TipoFactura
   FA.Fecha = MBFecha
   sSQL = "SELECT * " _
        & "FROM Catalogo_Lineas " _
        & "WHERE TL <> " & Val(adFalse) & " " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Fact = '" & FA.TC & "' " _
        & "AND Fecha <= #" & BuscarFecha(FA.Fecha) & "# " _
        & "AND Vencimiento >= #" & BuscarFecha(FA.Fecha) & "# " _
        & "ORDER BY Codigo "
   SelectDB_Combo DCLinea, AdoLinea, sSQL, "Concepto"
   If AdoLinea.Recordset.RecordCount > 0 Then DCLinea.Visible = True Else DCLinea.Visible = False
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
''  If KeyCode = vbKeyF12 Then
''     sSQL = "SELECT * " _
''          & "FROM Facturas " _
''          & "WHERE Fecha > Fecha_V "
''     Select_Adodc AdoNiveles, sSQL
''     With AdoNiveles.Recordset
''      If .RecordCount > 0 Then
''          Do While Not .EOF
''            .MoveNext
''          Loop
''      End If
''     End With
''  End If
End Sub

Private Sub MBFechaF_LostFocus()
   FechaValida MBFechaF
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_LostFocus()
   FechaValida MBFechaI
   FA.TC = TipoFactura
   FA.Fecha = MBFechaI
   sSQL = "SELECT * " _
        & "FROM Catalogo_Lineas " _
        & "WHERE TL <> " & Val(adFalse) & " " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Fact = '" & FA.TC & "' " _
        & "AND Fecha <= #" & BuscarFecha(FA.Fecha) & "# " _
        & "AND Vencimiento >= #" & BuscarFecha(FA.Fecha) & "# " _
        & "ORDER BY Codigo "
   SelectDB_Combo DCLinea, AdoLinea, sSQL, "Concepto"
   If AdoLinea.Recordset.RecordCount > 0 Then DCLinea.Visible = True Else DCLinea.Visible = False
   MBFechaF.Text = CLongFecha(CFechaLong(MBFechaI.Text) + 365)
End Sub

Public Sub Listar_Clientes_Grupo()
Dim sSaldo_Pendiente As String
  RatonReloj
 
  sSQL = "SELECT T, Cliente, Grupo, Direccion, Codigo, CI_RUC, Email, Email2, Fecha_N, Representante, TD_R, CI_RUC_R, DireccionT, Telefono_R, TelefonoT, EmailR, Saldo_Pendiente " _
       & "FROM Clientes " _
       & "WHERE Cliente <> '.' "
  If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
  If CheqRangos.value <> 0 Then
     sSQL = sSQL & "AND Grupo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
  Else
    'Tipo de Consulta
     If PorGrupo Then
        DGQuery.Caption = "LISTADO DE CLIENTES (GRUPO No. " & DCCliente.Text & ")"
        sSQL = sSQL & "AND Grupo = '" & DCCliente & "' "
     ElseIf PorDireccion Then
        DGQuery.Caption = "LISTADO DE CLIENTES (DIRECCION: " & DCCliente.Text & ")"
        sSQL = sSQL & "AND Direccion = '" & DCCliente & "' "
     Else
        DGQuery.Caption = "LISTADO DE CLIENTES"
        DCCliente.Text = "Todos"
     End If
  End If
  sSQL = sSQL _
       & "AND FA <> " & Val(adFalse) & " " _
       & "ORDER BY Grupo, Cliente "
  Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , True
  RatonNormal
End Sub

Public Sub Listar_Clientes_CxC_Meses()
Dim GrupoDireccion As String
  
  DGQuery.Visible = False
  GrupoDireccion = DCCliente.Text
  MesesNominaGrupo = ""
  If CTipoConsulta.Text = "Listar Todos" Then
     MsgBox "Este reporte se procesa solo para un Grupo determinado"
  Else
      RatonReloj
      Tutor = Ninguno
      TutorEmail = Ninguno
      MesesNominaGrupo = ""
     'MsgBox "Desktop Test"
      If CFechaLong(MBFechaI) <= CFechaLong(MBFechaF) Then Procesar_Saldos_CxC_Meses_SP GrupoDireccion, MBFechaI, MBFechaF, Tutor, TutorEmail, MesesNominaGrupo
  End If
  
  sSQL = "SELECT C.T, C.Cliente, "
  If Len(MesesNominaGrupo) > 1 Then sSQL = sSQL & MesesNominaGrupo & ", "
  sSQL = sSQL _
       & "C.Grupo, C.Direccion, C.Codigo, C.CI_RUC, C.Email, C.Email2, C.EmailR, C.Representante, C.TD_R, C.CI_RUC_R, C.Saldo_Pendiente " _
       & "FROM Clientes As C, Clientes_Auxiliar As CA " _
       & "WHERE C.Cliente <> '.' "
  If Mas_Grupos Then sSQL = sSQL & "AND C.DirNumero = '" & NumEmpresa & "' "
  If CheqRangos.value <> 0 Then
     sSQL = sSQL & "AND C.Grupo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
  Else
    'Tipo de Consulta
     If PorGrupo Then
        DGQuery.Caption = "LISTADO DE CLIENTES (GRUPO No. " & GrupoDireccion & ")"
        sSQL = sSQL & "AND C.Grupo = '" & GrupoDireccion & "' "
     ElseIf PorDireccion Then
        DGQuery.Caption = "LISTADO DE CLIENTES (DIRECCION: " & GrupoDireccion & ")"
        sSQL = sSQL & "AND C.Direccion = '" & GrupoDireccion & "' "
     Else
        DGQuery.Caption = "LISTADO DE CLIENTES"
        DCCliente.Text = "Todos"
        sSQL = sSQL & "AND C.Grupo = 'Todos' "
     End If
     If Len(TutorEmail) > 1 And Len(Tutor) > 1 Then DGQuery.Caption = DGQuery.Caption & ", TUTOR: " & Tutor & ", EMAIL TUTOR: " & TutorEmail
  End If
  sSQL = sSQL _
       & "AND C.Codigo = CA.Codigo " _
       & "AND C.FA <> " & Val(adFalse) & " " _
       & "ORDER BY C.Grupo, C.Cliente "
    Clipboard.Clear
    Clipboard.SetText sSQL
'  MsgBox "Desktop Test: " & sSQL
  Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , True
  
  DGQuery.Visible = True
  RatonNormal
End Sub

Public Sub Listar_Clientes_Email()
Dim DeudaCliente As String
  RatonReloj
  DGQuery.Visible = False
  LstClientes.Clear
  LstClientes.AddItem "TODOS" & String(85, " ") & "SALDO PENDIENTE"
  If PorGrupo Or PorDireccion Then
     With AdoQuery.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             sSaldo_Pendiente = Format$(.fields("Saldo_Pendiente"), "#,##0.00")
             DeudaCliente = .fields("Cliente") & String(80 - Len(.fields("Cliente")), " ") & String(15 - Len(sSaldo_Pendiente), " ") & sSaldo_Pendiente
             If Len(.fields("EmailR")) > 1 Then
                DeudaCliente = DeudaCliente & " -Email: " & .fields("EmailR")
             Else
                DeudaCliente = DeudaCliente & " -Email: " & .fields("Email")
             End If
             LstClientes.AddItem DeudaCliente
            .MoveNext
          Loop
          For I = 1 To LstClientes.ListCount - 1
              LstClientes.Selected(I) = False
          Next I
          LstClientes.Text = LstClientes.List(0)
      End If
     End With
  End If
  DGQuery.Visible = True
  RatonNormal
End Sub
    
'''Public Sub TipoConsultaCxC()
'''  RatonReloj
'''  SSTab2.Tab = 0
'''  Listar_Clientes_Grupo
''' 'Empezamos a ingrezar los valores de los alumnos
'''  If AnioI <> AnioF Then DGQuery1.Caption = DGQuery1.Caption & " AL " & AnioF
'''  DGParte.Caption = "LISTADO DE RUBROS A FACTURAR"
'''  ListarGrupos.Caption = "LISTADO DE SUSCRIPCIONES"
'''  RatonNormal
'''End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
    DGQuery.Visible = False
    FrmEmail.Visible = False
    Cuadricula = True
    Opcion = SSTab2.Tab
    FechaValida MBFechaI
    FechaValida MBFechaF
    FechaIni = BuscarFecha(MBFechaI)
    FechaFin = BuscarFecha(MBFechaF)
    Tipo_Rango_Grupos
    Select Case SSTab2.Tab
      Case 0: 'Listar Grupos
              DGQuery.Visible = True
              Listar_Clientes_Grupo
      Case 1: 'Pensiones mensuales del Periodo
              Reporte_CxC_Cuotas_SP Codigo1, Codigo2, MBFechaI, MBFechaF, SubTotal, Diferencia, TotalIngreso, ListaDeCampos, CheqResumen.value, CheqVenc.value
              RatonReloj
              sSQL = "SELECT " & ListaDeCampos & " " _
                   & "FROM Reporte_CxC_Cuotas " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND CodigoU = '" & CodigoUsuario & "' " _
                   & "ORDER BY GrupoNo,Cliente "
              Select_Adodc_Grid DGQuery, AdoQuery, sSQL, 2, True
              DGQuery.Visible = True
              Label9.Caption = Format$(SubTotal, "#,##0.00")
              Label10.Caption = Format$(Diferencia, "#,##0.00")
              Label4.Caption = Format$(TotalIngreso, "#,##0.00")
              RatonNormal
      Case 2: 'Listado de Becados o descuentos
              DGQuery.Visible = True
              SQLMsg1 = "ALUMNOS CON DESCUENTOS"
              SQLMsg2 = ""
              sSQL = "SELECT C.Cliente As Estudiantes,C.Grupo,CF.Mes,CF.Valor,CF.Descuento,CF.Descuento2,(CF.Valor-(CF.Descuento+CF.Descuento2)) As Total_Pagar,(((CF.Descuento+CF.Descuento2)/CF.Valor)*100) As Porc " _
                   & "FROM Clientes As C, Clientes_Facturacion As CF " _
                   & "WHERE CF.Item = '" & NumEmpresa & "' " _
                   & "AND CF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
                   & "AND (CF.Descuento+CF.Descuento2) <> 0 "
              If CheqRangos.value <> 0 Then
                 sSQL = sSQL & "AND C.Grupo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
              Else
                 If PorGrupo Then
                    sSQL = sSQL & "AND C.Grupo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
                 ElseIf PorDireccion Then
                    sSQL = sSQL & "AND C.Direccion BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
                 End If
              End If
              sSQL = sSQL _
                   & "AND CF.Codigo = C.Codigo " _
                   & "ORDER BY C.Grupo,C.Cliente,CF.Num_Mes "
              Select_Adodc_Grid DGQuery, AdoQuery, sSQL
      Case 3: 'Envio por mails
              Reporte_CxC_Cuotas_SP Codigo1, Codigo2, MBFechaI, MBFechaF, SubTotal, Diferencia, TotalIngreso, ListaDeCampos, CheqResumen.value, CheqVenc.value
              Listar_Clientes_Grupo
              Listar_Clientes_Email
              ListaDeCampos = Replace(ListaDeCampos, "Cliente,", "RCC.Cliente,")
              ListaDeCampos = Replace(ListaDeCampos, "GrupoNo,", "RCC.GrupoNo,")
              FrmEmail.Visible = True
              LstClientes.SetFocus
      Case 4: 'Resumen pensiones por mes
              sSQL = "SELECT CF.Periodo,COUNT(CP.Producto) AS Cant,CF.GrupoNo,CP.Producto,SUM(CF.Valor-(CF.Descuento+CF.Descuento2)) As Total " _
                   & "FROM Clientes_Facturacion As CF,Catalogo_Productos As CP " _
                   & "WHERE CP.Periodo = '" & Periodo_Contable & "' " _
                   & "AND CP.Item = '" & NumEmpresa & "' "
              If Month(MBFechaI) = Month(MBFechaF) Then
                 sSQL = sSQL & "AND CF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
              Else
                 sSQL = sSQL & "AND CF.Fecha <= #" & FechaFin & "# "
              End If
              sSQL = sSQL _
                   & "AND CF.GrupoNo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' " _
                   & "AND CF.Codigo_Inv = CP.Codigo_Inv " _
                   & "AND CF.Item = CP.Item " _
                   & "GROUP BY CF.Periodo,CF.GrupoNo,CP.Producto " _
                   & "UNION "
              sSQL = sSQL & "SELECT 'x' As Periodo,COUNT(CP.Producto) AS Cant,' ==> ' As GrupoNo,'Total por Cobrar' As Producto,SUM(CF.Valor-CF.Descuento) As Total " _
                   & "FROM Clientes_Facturacion As CF,Catalogo_Productos As CP " _
                   & "WHERE CP.Periodo = '" & Periodo_Contable & "' " _
                   & "AND CP.Item = '" & NumEmpresa & "' "
              If Month(MBFechaI) = Month(MBFechaF) Then
                 sSQL = sSQL & "AND CF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
              Else
                 sSQL = sSQL & "AND CF.Fecha <= #" & FechaFin & "# "
              End If
              sSQL = sSQL _
                   & "AND CF.GrupoNo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' " _
                   & "AND CF.Codigo_Inv = CP.Codigo_Inv " _
                   & "AND CF.Item = CP.Item " _
                   & "ORDER BY CF.Periodo,CF.GrupoNo,CP.Producto "
              Select_Adodc_Grid DGQuery, AdoQuery, sSQL, 2
              Clipboard.Clear
              Clipboard.SetText sSQL
              
              DGQuery.Visible = True
      Case 5: 'Listado de CxC Recaudadas
              Listar_Clientes_CxC_Meses
      Case 6: 'Listado Buses y Rubros
              DGQuery.Visible = True
              Listar_Deuda_por_Api
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim PorValor As Boolean

  ListarGrupos.Caption = "FACTURACION MULTIPLE"
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  Tipo_Rango_Grupos
  'MsgBox Button.key
   Select Case Button.key
     Case "Generar_Facturas": Generar_Facturas_Grupos
     Case "Listado_x_Grupos": Listado_x_Grupos
     Case "Generar_Eliminar_Rubros": Generar_Eliminar_Rubros
     Case "Generar_Deuda_Pendiente": Generar_Deuda_Pendiente
     Case "Recalcular_Fechas": Recalcular_Fechas
     Case "Impresora"
          DGQuery.Visible = False
          Cuadricula = True
          MensajeEncabData = SSTab2.Caption
          SQLMsg1 = DGQuery.Caption
          SQLMsg2 = ""
          SQLMsg3 = ""
          Select Case Opcion
            Case 0: ImprimirAdo AdoQuery, True, 1, 9
            Case 1: Total = 0
                    If CheqDesc.value Then SQLMsg1 = SQLMsg1 & ", Con Descuentos"
                    With AdoQuery.Recordset
                     If .RecordCount > 0 Then
                        .MoveFirst
                         Do While Not .EOF
                            Total = Total + .fields("Total")
                           .MoveNext
                         Loop
                    End If
                    End With
                    Imprimir_CxC_Grupos AdoQuery, 7, True
            Case 2: ImprimirAdo AdoQuery, True, 1, 9
            Case 3: ImprimirAdo AdoQuery, True, 1, 9, True
          End Select
          DGQuery.Visible = True
     Case "Recibos"
          DGQuery.Visible = False
          MensajeEncabData = SSTab2.Caption
          SQLMsg1 = DGQuery.Caption
          SQLMsg2 = ""
          SQLMsg3 = ""
         'MsgBox Opcion
          Select Case Opcion
            Case 0: Imprimir_Recibos_Cobros
            Case 1: Codigo1 = DCCliente.Text
                    Codigo2 = DCCliente.Text
                    If Codigo1 = "" Then Codigo1 = Ninguno
                    If Codigo2 = "" Then Codigo2 = Ninguno
                   'Control_Procesos  "I", "Reimpresion de Facturas desde la " & Factura_Desde & " a la " & Factura_Hasta
                    sSQL = "SELECT SUM(Valor) As SaldoPend,Codigo " _
                         & "FROM Clientes_Facturacion  " _
                         & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
                    If CheqRangos.value <> 0 Then sSQL = sSQL & "AND GrupoNo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
                    sSQL = sSQL & "GROUP BY Codigo "
                    Select_Adodc AdoAux, sSQL
                    If AdoAux.Recordset.RecordCount > 0 Then
                       Imprimir_Recibos_CxC_PreFA ListarGrupos, AdoAux, AdoAux2, MBFechaI, MBFechaF, Codigo1, Codigo2, FA
                    Else
                       MsgBox "No se puede imprimir el rando de Recibos"
                    End If
            Case 2:
            Case 3:
            Case 4:
            Case 5:
            Case 6:
          End Select
          DGQuery.Visible = True
     Case "Excel"
          DGQuery.Visible = False
          GenerarDataTexto ListarGrupos, AdoQuery
          DGQuery.Visible = True
     Case "Email_Tutor"
          Mensajes = "Reporte por Valor (SI), Por Referencia (NO):"
          Titulo = "FORMULARIO DE CONFIRMACION"
          If BoxMensaje = vbYes Then PorValor = True Else PorValor = False
          Reporte_CxC_Mes_PDF False, PorValor
          
          Mifecha = FechaSistema
          TMail.de = Lista_De_Correos(4).Correo_Electronico
          TMail.ListaMail = 255
          TMail.TipoDeEnvio = "NN"
          TMail.MensajeHTML = Leer_Archivo_Texto(RutaSistema & "\JAVASCRIPT\f_mail_basico.html")
          TMail.Asunto = "Listado de Cartera de " & DireccionCli & ", Hora (" & Time & ")"
          html_Informacion_adicional = "Reporte de cartera enviado al Tutor " & Tutor & ", Fecha de Envio: " & FechaSistema
          html_Detalle_adicional = ""
          FA.Fecha = FechaSistema
          FA.Recibo_No = Format(FA.Fecha, "yyyymmdd") & Format(FA.Factura, "000000000")
          TMail.Adjunto = RutaSysBases & "\TEMP\" & NombreArchivo & ".pdf"
          TMail.para = ""
          Insertar_Mail TMail.para, TutorEmail
          FEnviarCorreos.Show 1
          TMail.para = ""
          TMail.ListaMail = 0
     Case "Salir": Unload ListarGrupos
   End Select
End Sub

Private Sub TxtAyuda_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
     TxtAyuda.Visible = False
     DCCliente.SetFocus
  End If
End Sub

Public Sub Listar_Deuda_por_Api()
Dim FechaTope As String
Dim SiActualizo As Boolean
Dim ExisteUno As Boolean
    
    If Len(URLToken) > 1 And Len(Token) > 1 Then
        FechaTope = BuscarFecha(FechaSistema)
        If CheqVenc.value <> 0 Then FechaTope = BuscarFecha(MBFechaF.Text)
    
        sSQL = "UPDATE Clientes " _
             & "SET Saldo_Pendiente = 0, Dias_Mora = 0 " _
             & "WHERE Codigo <> '.' "
        Ejecutar_SQL_SP sSQL
            
        sSQL = "UPDATE Clientes " _
             & "SET Saldo_Pendiente = (SELECT ROUND(SUM(CF.Valor-CF.Descuento-CF.Descuento2),2,0) " _
             & "                       FROM Clientes_Facturacion As CF " _
             & "                       WHERE CF.Item = '" & NumEmpresa & "' " _
             & "                       AND CF.Fecha <= '" & FechaTope & "' " _
             & "                       AND CF.Codigo = Clientes.Codigo) " _
             & "WHERE Codigo <> '.' "
        If CheqRangos.value Then sSQL = sSQL & "AND Grupo BETWEEN '" & DCGrupoI.Text & "' and '" & DCGrupoF.Text & "' "
        Ejecutar_SQL_SP sSQL
        
        sSQL = "UPDATE Clientes " _
             & "SET Fecha_Cad = (SELECT MIN(CF.Fecha) " _
             & "                 FROM Clientes_Facturacion As CF " _
             & "                 WHERE CF.Item = '" & NumEmpresa & "' " _
             & "                 AND CF.Fecha <= '" & FechaTope & "' " _
             & "                 AND CF.Codigo = Clientes.Codigo) " _
             & "WHERE Codigo <> '.' "
        If CheqRangos.value Then sSQL = sSQL & "AND Grupo BETWEEN '" & DCGrupoI.Text & "' and '" & DCGrupoF.Text & "' "
        Ejecutar_SQL_SP sSQL
        
        sSQL = "UPDATE Clientes " _
             & "SET Saldo_Pendiente = 0 " _
             & "WHERE Saldo_Pendiente IS NULL "
        Ejecutar_SQL_SP sSQL
        
        sSQL = "UPDATE Clientes " _
             & "SET Fecha_Cad = '" & FechaTope & "' " _
             & "WHERE Fecha_Cad IS NULL "
        Ejecutar_SQL_SP sSQL
        
        sSQL = "UPDATE Clientes " _
             & "SET Dias_Mora = DATEDIFF(day,Fecha_Cad,'" & FechaTope & "') " _
             & "WHERE Codigo <> '.' "
        Ejecutar_SQL_SP sSQL
        Total = 0
        sSQL = "SELECT Grupo, Cliente As Estudiante, CI_RUC As Cedula, Saldo_Pendiente, Dias_Mora, EmailR, Codigo " _
             & "FROM Clientes " _
             & "WHERE FA <> 0 "
        If CheqRangos.value Then sSQL = sSQL & "AND Grupo BETWEEN '" & DCGrupoI.Text & "' and '" & DCGrupoF.Text & "' "
        sSQL = sSQL & "ORDER BY Grupo, Cliente "
        Select_Adodc_Grid DGQuery, AdoQuery, sSQL
        DGQuery.Visible = False
        With AdoQuery.Recordset
         If .RecordCount > 0 Then
             Do While Not .EOF
                Total = Total + .fields("Saldo_Pendiente")
               .MoveNext
             Loop
         End If
        End With
        Label4.Caption = Format(Total, "#,##0.00")
        DGQuery.Visible = True
        ExisteUno = True
        Mensajes = "Actualizar Deuda Pendiente de Clientes"
        Titulo = "Formulario de Deuda Pendiente"
        If BoxMensaje = vbYes Then
           RatonReloj
           DGQuery.Visible = False
           TextoImprimio = ""
           MiTiempo1 = Time
           With AdoQuery.Recordset
            If .RecordCount > 0 Then
                Progreso_Barra.Mensaje_Box = "Procesando actualizacion de Deuda Pendiente..."
                Progreso_Iniciar
                Progreso_Iniciar_Errores
                Progreso_Barra.Incremento = 0
                Progreso_Barra.Valor_Maximo = .RecordCount + 10
               .MoveFirst
                Do While Not .EOF
                   Progreso_Barra.Mensaje_Box = "[" & Format$(Time - MiTiempo1, "HH:MM:SS") & "] Actualizando del " & .fields("Grupo") & ". Al Estudiante: " & ULCase(.fields("Estudiante"))
                   Progreso_Esperar
                   SiActualizo = post_URL_JSon(.fields("Cedula"), .fields("Saldo_Pendiente"), .fields("Dias_Mora"))
                   If Not SiActualizo Then
                      If ExisteUno Then
                         Cadena = "GRUPO      " & vbTab & "CEDULA        " & vbTab & "ESTUDIANTE"
                         Insertar_Texto_Temporal_SP Cadena
                      End If
                      Cadena = .fields("Grupo") & String(11 - Len(.fields("Grupo")), " ") & vbTab & .fields("Cedula") & String(14 - Len(.fields("Cedula")), " ") & vbTab & .fields("Estudiante") & " no se pudo actualizar"
                      TextoImprimio = TextoImprimio & Cadena & vbCrLf
                      Insertar_Texto_Temporal_SP Cadena
                      ExisteUno = False
                   End If
                  .MoveNext
                Loop
            End If
           End With
           Progreso_Final
           Progreso_Barra.Mensaje_Box = "[" & Format$(Time - MiTiempo1, "HH:MM:SS") & "] Proceso Terminado"
           Progreso_Esperar
           DGQuery.Visible = True
           SSTab2.Tab = 0
           RatonNormal
           MsgBox "Proceso Terminado con exito"
           If Len(TextoImprimio) > 2 Then FInfoError.Show
        End If
    Else
        sSQL = "SELECT Grupo, Cliente As Estudiante, CI_RUC As Cedula, Saldo_Pendiente, Dias_Mora, EmailR, Codigo " _
             & "FROM Clientes " _
             & "WHERE FA <> 0 " _
             & "AND Grupo = 'Todos' "
        Select_Adodc_Grid DGQuery, AdoQuery, sSQL
        MsgBox "Esta Entidad no tiene configurado la conexion con Sistema Educativo Externo"
    End If
End Sub

Private Sub TxtNota_GotFocus()
    MarcarTexto TxtNota
End Sub

Private Sub TxtNota_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNota_LostFocus()
    TextoValido TxtNota, , True
End Sub

Private Sub TxtObs_GotFocus()
    MarcarTexto TxtObs
End Sub

Private Sub TxtObs_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtObs_LostFocus()
    TextoValido TxtObs, , True
End Sub
