VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form MayorizarErrores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "."
   ClientHeight    =   420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   7455
   Begin MSComctlLib.ProgressBar ProcBar 
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   3885
      Top             =   105
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
   Begin MSAdodcLib.Adodc AdoTrans 
      Height          =   330
      Left            =   105
      Top             =   105
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
      Caption         =   "Trans"
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
   Begin MSAdodcLib.Adodc AdoSubCtas 
      Height          =   330
      Left            =   1890
      Top             =   105
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
      Caption         =   "SubCtas"
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
Attribute VB_Name = "MayorizarErrores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Dim Primero As Boolean
Dim Cod_Cta As String
Dim Num_Comp As Long
  TextoImprimio = ""
  MayorizarErrores.Caption = "Espere un momento.... Iniciando la mayorizacion."
  FechaIni = BuscarFecha(FechaPeriodo)
  FechaFin = BuscarFecha(FechaFinal)
 'Seteamos procesos para la mayorizacion.
  If Periodo_Contable = "" Then Periodo_Contable = Ninguno
  
  MayorizarErrores.Caption = "Mayorizando."
  RatonReloj
  SQL1 = "UPDATE Catalogo_Cuentas " _
       & "SET Periodo = '" & Periodo_Contable & "' " _
       & "WHERE Periodo IS NULL "
  Ejecutar_SQL_SP SQL1

  MayorizarErrores.Caption = "Reindexando Comprobantes"
  RatonReloj
  SQL1 = "UPDATE Catalogo_Cuentas " _
       & "SET Procesado = 0 " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP SQL1
  
  RatonReloj
  SQL1 = "UPDATE Transacciones " _
       & "SET Procesado = 0 " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP SQL1
  
  RatonReloj
  SQL1 = "UPDATE Trans_SubCtas " _
       & "SET Procesado = 0 " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP SQL1
    
  RatonReloj
  MayorizarErrores.Caption = "Actualizando Usuarios..."
  sSQL = "UPDATE Comprobantes " _
       & "SET Si_Existe = " & Val(adFalse) & " " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' "
  Ejecutar_SQL_SP sSQL
  
  RatonReloj
  If SQL_Server Then
     sSQL = "UPDATE Comprobantes " _
          & "SET Si_Existe = " & Val(adTrue) & " " _
          & "FROM Comprobantes As Co,Accesos As C "
  Else
     sSQL = "UPDATE Comprobantes As Co,Accesos As C " _
          & "SET Si_Existe = " & Val(adTrue) & " "
  End If
  sSQL = sSQL & "WHERE Co.Periodo = '" & Periodo_Contable & "' " _
       & "AND Co.Item = '" & NumEmpresa & "' " _
       & "AND Co.CodigoU = C.Codigo "
  Ejecutar_SQL_SP sSQL
  
  RatonReloj
  sSQL = "UPDATE Comprobantes " _
       & "SET CodigoU = '.' " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Si_Existe = " & Val(adFalse) & " "
  Ejecutar_SQL_SP sSQL
  
  RatonReloj
  MayorizarErrores.Caption = "Actualizando Encabezados..."
  sSQL = "UPDATE Comprobantes " _
       & "SET Si_Existe = " & Val(adFalse) & " " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' "
  Ejecutar_SQL_SP sSQL
  
  RatonReloj
  If SQL_Server Then
     sSQL = "UPDATE Comprobantes " _
          & "SET Si_Existe = " & Val(adTrue) & " " _
          & "FROM Comprobantes As Co,Clientes As C "
  Else
     sSQL = "UPDATE Comprobantes As Co,Clientes As C " _
          & "SET Si_Existe = " & Val(adTrue) & " "
  End If
  sSQL = sSQL & "WHERE Co.Periodo = '" & Periodo_Contable & "' " _
       & "AND Co.Item = '" & NumEmpresa & "' " _
       & "AND Co.Codigo_B = C.Codigo "
  Ejecutar_SQL_SP sSQL
  
  RatonReloj
  sSQL = "UPDATE Comprobantes " _
       & "SET Codigo_B = '.' " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Si_Existe = " & Val(adFalse) & " "
  Ejecutar_SQL_SP sSQL
    
  Codigo1 = ""
  Si_No = True
  MayorizarErrores.Caption = "Mayorizando la Cuenta: Verificando Encabezado de Comprobantes "
  sSQL = "SELECT TP,Numero,Fecha " _
       & "FROM Comprobantes " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TP IN ('CD','CE','CI') " _
       & "ORDER BY TP,Numero,Fecha "
  Select_Adodc AdoTrans, sSQL
  RatonReloj
  With AdoTrans.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          sSQL = "SELECT TP,Numero,Fecha " _
               & "FROM Transacciones " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND TP = '" & .Fields("TP") & "' " _
               & "AND Numero = " & .Fields("Numero") & " "
          Select_Adodc AdoCtas, sSQL
          If AdoCtas.Recordset.RecordCount <= 0 Then
             If Si_No Then Codigo1 = Codigo1 & vbCrLf & "COMPROBANTES ELIMINADOS NO EXISTE TRANSACCIONES:" & vbCrLf
             Si_No = False
             Codigo1 = Codigo1 & "[" & .Fields("Fecha") & "] " & .Fields("TP") & " - " & .Fields("Numero") & vbCrLf
             sSQL = "DELETE * " _
                  & "FROM Comprobantes " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND TP = '" & .Fields("TP") & "' " _
                  & "AND Numero = " & .Fields("Numero") & " "
             Ejecutar_SQL_SP sSQL
             
             sSQL = "DELETE * " _
                  & "FROM Trans_Air " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND TP = '" & .Fields("TP") & "' " _
                  & "AND Numero = " & .Fields("Numero") & " "
             Ejecutar_SQL_SP sSQL
             
             sSQL = "DELETE * " _
                  & "FROM Trans_Compras " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND TP = '" & .Fields("TP") & "' " _
                  & "AND Numero = " & .Fields("Numero") & " "
             Ejecutar_SQL_SP sSQL
                          
             sSQL = "DELETE * " _
                  & "FROM Trans_Kardex " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND TP = '" & .Fields("TP") & "' " _
                  & "AND Numero = " & .Fields("Numero") & " "
             Ejecutar_SQL_SP sSQL
             
             sSQL = "DELETE * " _
                  & "FROM Trans_Prestamos " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND TP = '" & .Fields("TP") & "' " _
                  & "AND Numero = " & .Fields("Numero") & " "
             Ejecutar_SQL_SP sSQL
             
             sSQL = "DELETE * " _
                  & "FROM Trans_SubCtas " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND TP = '" & .Fields("TP") & "' " _
                  & "AND Numero = " & .Fields("Numero") & " "
             Ejecutar_SQL_SP sSQL
          End If
         .MoveNext
       Loop
   End If
  End With
  RatonNormal
  Codigo = ""
  sSQL = "SELECT TP,Numero,Fecha,SUM(Debe),SUM(Haber) " _
       & "FROM Transacciones " _
       & "WHERE Fecha > #" & BuscarFecha(FechaSistema) & "# " _
       & "AND T <> '" & Anulado & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "GROUP BY TP,Numero,Fecha " _
       & "ORDER BY TP,Numero,Fecha "
  Select_Adodc AdoTrans, sSQL
  With AdoTrans.Recordset
   If .RecordCount > 0 Then
       Codigo = "COMPROBANTES FUERA DE FECHA:" & vbCrLf
       Do While Not .EOF
          Codigo = Codigo & "[" & .Fields("Fecha") & "] " & .Fields("TP") & " - " & .Fields("Numero") & vbCrLf
         .MoveNext
       Loop
   End If
  End With
  If Round(SumaDebe - SumaHaber, 2) <> 0 Then
     Codigo = Codigo & vbCrLf & "LAS TRANSACCIONES NO CUADRAN POR:" & vbCrLf
     sSQL = "SELECT TP,Numero,Fecha,SUM(Debe) As TotDeb,SUM(Haber) As TotHab " _
          & "FROM Transacciones " _
          & "WHERE T <> '" & Anulado & "' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "GROUP BY TP,Numero,Fecha " _
          & "HAVING SUM(Debe) <> SUM(Haber) "
     Select_Adodc AdoTrans, sSQL
     With AdoTrans.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             Codigo = Codigo & "[" & .Fields("Fecha") & "] " & .Fields("TP") & " - " & .Fields("Numero") & "     Por  " & Format(Abs(.Fields("TotDeb") - .Fields("TotHab")), "#,##0.00") & vbCrLf
            .MoveNext
          Loop
      End If
     End With
  End If
 'MsgBox "Fin de la mayorizacion"
  
  MayorizarErrores.Caption = "Mayorizando la Cuenta: Comprobando Inventario "
  sSQL = "SELECT TP,Numero,Fecha " _
       & "FROM Trans_Kardex " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND LEN(TP) >= 2 " _
       & "AND Numero <> 0 " _
       & "GROUP BY TP,Numero,Fecha " _
       & "ORDER BY TP,Numero,Fecha "
  Select_Adodc AdoTrans, sSQL
  Si_No = True
  With AdoTrans.Recordset
   If .RecordCount > 0 Then
      'SetProgBar ProcBar, .RecordCount
       Do While Not .EOF
          sSQL = "SELECT TP,Numero,Fecha " _
               & "FROM Comprobantes " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND TP = '" & .Fields("TP") & "' " _
               & "AND Numero = " & .Fields("Numero") & " " _
               & "AND Fecha = #" & BuscarFecha(.Fields("Fecha")) & "# "
          Select_Adodc AdoCtas, sSQL
          If AdoCtas.Recordset.RecordCount <= 0 Then
             MayorizarErrores.Caption = "Mayorizando la Cuenta: Comprobando Inventario "
             SetAdoAddNew "Comprobantes"
             SetAdoFields "T", Normal
             SetAdoFields "TP", .Fields("TP")
             SetAdoFields "Numero", .Fields("Numero")
             SetAdoFields "Fecha", .Fields("Fecha")
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "Periodo", Periodo_Contable
             SetAdoFields "Codigo_B", Ninguno
             SetAdoFields "Concepto", "(...Comprobante incompleto, Revisar Inventario...)"
             SetAdoUpdate
             If Si_No Then Codigo = Codigo & vbCrLf & "COMPROBANTES MAL PROCESADOS:" & vbCrLf
             Si_No = False
             Codigo = Codigo & "[" & .Fields("Fecha") & "] " & .Fields("TP") & " - " & .Fields("Numero") & vbCrLf
          End If
         'IncProgBar ProcBar
         .MoveNext
       Loop
   End If
  End With
  MayorizarErrores.Caption = "Mayorizando la Cuenta: Comprobando SubModulos "
  sSQL = "SELECT TP,Numero,Fecha " _
       & "FROM Trans_SubCtas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "GROUP BY TP,Numero,Fecha " _
       & "ORDER BY TP,Numero,Fecha "
  Select_Adodc AdoTrans, sSQL
  With AdoTrans.Recordset
   If .RecordCount > 0 Then
      ' SetProgBar ProcBar, .RecordCount
       Do While Not .EOF
          sSQL = "SELECT TP,Numero,Fecha " _
               & "FROM Comprobantes " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND TP = '" & .Fields("TP") & "' " _
               & "AND Numero = " & .Fields("Numero") & " " _
               & "AND Fecha = #" & BuscarFecha(.Fields("Fecha")) & "# "
          Select_Adodc AdoCtas, sSQL
          If AdoCtas.Recordset.RecordCount <= 0 Then
             MayorizarErrores.Caption = "Mayorizando la Cuenta: Comprobando SubModulos "
             SetAdoAddNew "Comprobantes"
             SetAdoFields "T", Normal
             SetAdoFields "TP", .Fields("TP")
             SetAdoFields "Numero", .Fields("Numero")
             SetAdoFields "Fecha", .Fields("Fecha")
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "Periodo", Periodo_Contable
             SetAdoFields "Codigo_B", Ninguno
             SetAdoFields "Concepto", "(...Comprobante incompleto, Revisar Transacciones de SubModulos...)"
             SetAdoUpdate
             If Si_No Then Codigo = Codigo & vbCrLf & "COMPROBANTES INCOMPLETOS:" & vbCrLf
             Si_No = False
             Codigo = Codigo & "[" & .Fields("Fecha") & "] " & .Fields("TP") & " - " & .Fields("Numero") & vbCrLf
          End If
         ' IncProgBar ProcBar
         .MoveNext
       Loop
   End If
  End With
  MayorizarErrores.Caption = "Mayorizando la Cuenta: Comprobando Transacciones "
  sSQL = "SELECT TP,Numero,Fecha " _
       & "FROM Transacciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "GROUP BY TP,Numero,Fecha " _
       & "ORDER BY TP,Numero,Fecha "
  Select_Adodc AdoTrans, sSQL
  With AdoTrans.Recordset
   If .RecordCount > 0 Then
       'SetProgBar ProcBar, .RecordCount
       Do While Not .EOF
          sSQL = "SELECT TP,Numero,Fecha " _
               & "FROM Comprobantes " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND TP = '" & .Fields("TP") & "' " _
               & "AND Numero = " & .Fields("Numero") & " " _
               & "AND Fecha = #" & BuscarFecha(.Fields("Fecha")) & "# "
          Select_Adodc AdoCtas, sSQL
          If AdoCtas.Recordset.RecordCount <= 0 Then
             MayorizarErrores.Caption = "Mayorizando la Cuenta: Comprobando Transacciones "
             SetAdoAddNew "Comprobantes"
             SetAdoFields "T", Normal
             SetAdoFields "TP", .Fields("TP")
             SetAdoFields "Numero", .Fields("Numero")
             SetAdoFields "Fecha", .Fields("Fecha")
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "Periodo", Periodo_Contable
             SetAdoFields "Codigo_B", Ninguno
             SetAdoFields "Concepto", "(...Comprobante incompleto, Revisar Transacciones...)"
             SetAdoUpdate
             If Si_No Then Codigo = Codigo & vbCrLf & "COMPROBANTES INCOMPLETOS:" & vbCrLf
             Si_No = False
             Codigo = Codigo & "[" & .Fields("Fecha") & "] " & .Fields("TP") & " - " & .Fields("Numero") & vbCrLf
          End If
          'IncProgBar ProcBar
         .MoveNext
       Loop
   End If
  End With
  MayorizarErrores.Caption = "Mayorizando la Cuenta: Transacciones Duplicadas"
  sSQL = "SELECT TP,Numero,Fecha,Cta,ID,Debe,Haber,Count(ID) As No_Reg " _
       & "FROM Transacciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "GROUP BY TP,Numero,Fecha,Cta,ID,Debe,Haber " _
       & "HAVING Count(ID) > 1 " _
       & "ORDER BY TP,Numero,Fecha,Cta,ID "
  Select_Adodc AdoTrans, sSQL
  With AdoTrans.Recordset
   If .RecordCount > 0 Then
       SetProgBar ProcBar, .RecordCount
       FechaTexto = .Fields("Fecha")
       TipoDoc = .Fields("TP")
       Numero = .Fields("Numero")
       Cta = ""
       Do While Not .EOF
          If TipoDoc <> .Fields("TP") Or Numero <> .Fields("Numero") Then
             If Si_No Then Codigo = Codigo & vbCrLf & "TRANSACCIONES DUPLICADAS:" & vbCrLf
             Si_No = False
             Codigo = Codigo & " [" & FechaTexto & "] " & TipoDoc & " - " & Numero & " - " & Cta & vbCrLf
             TipoDoc = .Fields("TP")
             Numero = .Fields("Numero")
             FechaTexto = .Fields("Fecha")
             Cta = ""
          End If
          Cta = Cta & .Fields("Cta") & "|"
          IncProgBar ProcBar
         .MoveNext
       Loop
       Codigo = Codigo & Format(ProcBar.value + 1, "000000") & " [" & FechaTexto & "] " & TipoDoc & " - " & Numero & " - " & Cta & vbCrLf
   End If
  End With
  
 'Actualizamos los codigo de los Clientes Abonos
  If SQL_Server Then
     sSQL = "UPDATE Detalle_Factura " _
          & "SET CodigoC = F.CodigoC " _
          & "FROM Trans_Abonos As DF,Facturas As F "
  Else
     sSQL = "UPDATE Trans_Abonos As DF,Facturas As F " _
          & "SET DF.CodigoC = F.CodigoC "
  End If
  sSQL = sSQL _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND DF.Autorizacion = F.Autorizacion " _
       & "AND DF.Serie = F.Serie " _
       & "AND DF.Factura = F.Factura " _
       & "AND DF.Item = F.Item " _
       & "AND DF.Periodo = F.Periodo " _
       & "AND DF.TP = F.TC " _
       & "AND DF.CodigoC <> F.CodigoC "
  Ejecutar_SQL_SP sSQL
  
 'Actualizamos los codigo de los Clientes Detalle Factura
  If SQL_Server Then
     sSQL = "UPDATE Detalle_Factura " _
          & "SET CodigoC = F.CodigoC " _
          & "FROM Detalle_Factura As DF,Facturas As F "
  Else
     sSQL = "UPDATE Detalle_Factura As DF,Facturas As F " _
          & "SET DF.CodigoC = F.CodigoC "
  End If
  sSQL = sSQL _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND DF.Autorizacion = F.Autorizacion " _
       & "AND DF.Serie = F.Serie " _
       & "AND DF.Factura = F.Factura " _
       & "AND DF.Item = F.Item " _
       & "AND DF.Periodo = F.Periodo " _
       & "AND DF.TC = F.TC " _
       & "AND DF.CodigoC <> F.CodigoC "
  Ejecutar_SQL_SP sSQL
  
  'MsgBox Codigo
  MayorizarErrores.Caption = "Mayorizando la Cuenta: Transacciones de Submodulos Duplicadas"
  sSQL = "SELECT TP,Numero,Fecha,Cta,ID,Debitos,Creditos,Count(ID) As No_Reg " _
       & "FROM Trans_SubCtas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "GROUP BY TP,Numero,Fecha,Cta,ID,Debitos,Creditos " _
       & "HAVING Count(ID) > 1 " _
       & "ORDER BY TP,Numero,Fecha,Cta,ID "
  Select_Adodc AdoTrans, sSQL
  With AdoTrans.Recordset
   If .RecordCount > 0 Then
       SetProgBar ProcBar, .RecordCount
       TipoDoc = .Fields("TP")
       Numero = .Fields("Numero")
       FechaTexto = .Fields("Fecha")
       Cta = ""
       Do While Not .EOF
          If TipoDoc <> .Fields("TP") Or Numero <> .Fields("Numero") Then
             If Si_No Then Codigo = Codigo & vbCrLf & "TRANSACCIONES DUPLICADAS:" & vbCrLf
             Si_No = False
             Codigo = Codigo & " [" & FechaTexto & "] " & TipoDoc & " - " & Numero & " - " & Cta & vbCrLf
             TipoDoc = .Fields("TP")
             Numero = .Fields("Numero")
             FechaTexto = .Fields("Fecha")
             Cta = ""
          End If
          Cta = Cta & .Fields("Cta") & "|"
          IncProgBar ProcBar
         .MoveNext
       Loop
       Codigo = Codigo & Format(ProcBar.value + 1, "000000") & " [" & FechaTexto & "] " & TipoDoc & " - " & Numero & " - " & Cta & vbCrLf
   End If
  End With
  TextoImprimio = ""
  If Codigo1 <> "" Then Codigo = Codigo & vbCrLf & Codigo1
  If Codigo <> "" Then
     Cadena = "Warning:" & vbCrLf & Codigo
     TextoImprimio = Cadena
  End If
  'MsgBox ".........." & vbCrLf & Cadena & ".........."
  Unload MayorizarErrores
  If TextoImprimio <> "" Then FInfoError.Show
End Sub

Private Sub Form_Load()
  CentrarForm MayorizarErrores
  ConectarAdodc AdoCtas
  ConectarAdodc AdoTrans
  ConectarAdodc AdoSubCtas
 'SetProgBar ProcBar, 100
End Sub

Private Sub ProcBar_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
' Forma de mayorizar
End Sub
