Attribute VB_Name = "SubSQLServer"
Option Explicit

Public Sub Iniciar_Stored_Procedure(cMensajeProceso As String, _
                                    cMiSQL As ADODB.Connection, _
                                    cMiCmd As ADODB.Command, _
                                    cMiReg As ADODB.Recordset)
    RatonReloj
''  Progreso_Barra.Mensaje_Box = cMensajeProceso
''  Progreso_Iniciar
    
   'Seteamos la conexion
    Set cMiSQL = New ADODB.Connection
    Set cMiCmd = New ADODB.Command
    Set cMiReg = New ADODB.Recordset
    
    cMiSQL.CursorLocation = adUseClient
    If SQL_Server Then cMiSQL.ConnectionString = AdoStrCnn Else cMiSQL.ConnectionString = AdoStrCnnMySQL
    cMiSQL.open
    cMiSQL.CommandTimeout = 0
    
   'Ejecutar Procedimiento Almacenado y guarda resultados en recordset
    cMiCmd.ActiveConnection = cMiSQL
    cMiCmd.CommandType = adCmdStoredProc
End Sub

Public Sub Procesar_Stored_Procedure(cMiCmd As ADODB.Command, _
                                     cMiReg As ADODB.Recordset)
Dim IdP As Integer
Dim ListP As String
On Error GoTo Errorhandler
    RatonReloj
    ListP = cMiCmd.CommandText & vbCrLf
    For IdP = 0 To cMiCmd.Parameters.Count - 1
        ListP = ListP & cMiCmd.Parameters.Item(IdP).Name & " = '" & cMiCmd.Parameters.Item(IdP) & "'" & vbCrLf
    Next IdP
   'MsgBox Len(ListP) & vbCrLf & vbCrLf & ListP
   'Generar_File_SQL "Store_Procedure", ListP
    
'''    Progreso_Esperar True
    cMiCmd.CommandTimeout = 0
    cMiCmd.Prepared = True
    Set cMiReg = New ADODB.Recordset
    Set cMiReg = cMiCmd.Execute
Exit Sub
Errorhandler:
    RatonNormal
    MsgBox "Error:(" & Err & ") en la conexion de Internet"
    Exit Sub
End Sub

Public Sub Finalizar_Stored_Procedure(cMiSQL As ADODB.Connection, _
                                      cMiCmd As ADODB.Command, _
                                      cMiReg As ADODB.Recordset)
    Set cMiCmd.ActiveConnection = Nothing
    cMiSQL.Close
    Set cMiReg = Nothing
    Set cMiSQL = Nothing
    Set cMiCmd = Nothing
'''    Progreso_Final
    RatonNormal
End Sub

Public Sub Lista_Mensaje_SP_MySQL(MensajeAyuda As String)
Dim cnMySQL As ADODB.Connection
Dim rsMySQL As ADODB.Recordset
Dim cmdMySQL As ADODB.Command

    RatonReloj
    If Ping_IP("db.diskcoversystem.com") Then
       'Conexion a MySQL del servidor en las nubes
        Set cmdMySQL = New ADODB.Command
        Set cnMySQL = New ADODB.Connection
        cnMySQL.ConnectionString = AdoStrCnnMySQL
        cnMySQL.open
        Set cmdMySQL.ActiveConnection = cnMySQL
        cmdMySQL.CommandType = adCmdText
    
       'Parametros de entrada y de salida
        cmdMySQL.CommandText = "Call sp_lista_mensaje(@MensajeAyuda);"
    
       'Enviamos los parametro de solo entrada al SP
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("MensajeAyuda", adVarChar, adParamInput, 1024, MensajeAyuda)
    
       'Ejecucion del SP en MySQL
        Set rsMySQL = cmdMySQL.Execute
    
       'Recolectamos los resultados de los parametros de salida
        Set rsMySQL = cnMySQL.Execute("SELECT @MensajeAyuda;")
    
       'Pasamos a variables globales lso resultados del SP
        If Not rsMySQL.EOF Then MensajeAyuda = rsMySQL.fields(0)
    
       'Cerramos la conexion con MySQL
        rsMySQL.Close
        cnMySQL.Close
        
       'Liberando de la memoria es control de conexion
        Set cmdMySQL.ActiveConnection = Nothing
        Set rsMySQL = Nothing
        Set cnMySQL = Nothing
        Set cmdMySQL = Nothing
    Else
        MensajeAyuda = "AHORA YA ESTAMOS EN LAS NUBES, VISITANOS:" & vbCrLf _
                     & "https://www.diskcoversystem.com"
    End If
    RatonNormal
End Sub

Public Sub Acceso_IP_PCs_SP_MySQL(vActivo As Boolean)
Dim cnMySQL As ADODB.Connection
Dim rsMySQL As ADODB.Recordset
Dim cmdMySQL As ADODB.Command
Dim Mifecha1 As String
Dim MiHora1 As String

   'Conexion a MySQL del servidor en las nubes
    If Ping_IP("db.diskcoversystem.com") Then
         RatonReloj
         Mifecha1 = CStr(Format(date, "yyyymmdd"))
         MiHora1 = CStr(Format(Time, "hh:mm:ss"))
         
         Set cmdMySQL = New ADODB.Command
         Set cnMySQL = New ADODB.Connection
         cnMySQL.ConnectionString = AdoStrCnnMySQL
         cnMySQL.open
          
         Set cmdMySQL.ActiveConnection = cnMySQL
         cmdMySQL.CommandType = adCmdText
          
        'Parametros de entrada y de salida
         cmdMySQL.CommandText = "Call sp_acceso_ip_pcs(?,?,?,?,?,?,?,@pActivo);"
                               
        'Enviamos los parametro de solo entrada al SP
         cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("CodigoU", adVarChar, adParamInput, 10, CodigoUsuario)
         cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("IP_Local", adVarChar, adParamInput, 15, IP_PC.IP_PC)
         cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("IP_WAN", adVarChar, adParamInput, 15, IP_PC.WAN_PC)
         cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("PC_Nombre", adVarChar, adParamInput, 15, IP_PC.Nombre_PC)
         cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("PC_MAC", adVarChar, adParamInput, 17, IP_PC.MAC_PC)
         cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("Fecha", adVarChar, adParamInput, 10, Mifecha1)
         cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("Hora", adVarChar, adParamInput, 8, MiHora1)

        'Ejecucion del SP en MySQL
         Set rsMySQL = cmdMySQL.Execute
          
        'Recolectamos los resultados de los parametros de salida
         Set rsMySQL = cnMySQL.Execute("SELECT @pActivo;")
          
        'Pasamos a variables globales lso resultados del SP
         If Not rsMySQL.EOF Then vActivo = rsMySQL.fields(0)
         
        ' MsgBox CodigoUsuario & vbCrLf & IP_PC.IP_PC & vbCrLf & IP_PC.WAN_PC & vbCrLf & IP_PC.Nombre_PC & vbCrLf & IP_PC.MAC_PC & vbCrLf & Mifecha1 & vbCrLf & MiHora1 & vbCrLf & vActivo
         
        'Cerramos la conexion con MySQL
         rsMySQL.Close
         cnMySQL.Close
         
        'Set cmdMySQL.ActiveConnection = Nothing
         Set rsMySQL = Nothing
         Set cnMySQL = Nothing
         Set cmdMySQL = Nothing
         RatonNormal
    End If
End Sub

Public Sub Datos_Iniciales_Entidad_SP_MySQL()
Dim cnMySQL As ADODB.Connection
Dim rsMySQL As ADODB.Recordset
Dim cmdMySQL As ADODB.Command
Dim ParametrosDeSalida As String

    ServidorMySQL = False
    PCActivo = True
    EstadoUsuario = True
    IDEntidad = 0
    DescripcionEstado = "OK"
    EstadoEmpresa = Ninguno
    Fecha_CE = FechaSistema
    Fecha_CO = FechaSistema
    Fecha_VPN = FechaSistema
    Fecha_DB = FechaSistema
    Fecha_P12 = FechaSistema
    SerieFE = Ninguno
    MicroEmpresa = Ninguno
    AgenteRetencion = Ninguno
    Cartera = 0
    Cant_FA = 0
    TipoPlan = 0
    If Ping_IP("db.diskcoversystem.com") Then
        RatonReloj
       'Set cnMySQL = CreateObject("ADODB.Connection")
       'Conexion a MySQL del servidor en las nubes
        Set cmdMySQL = New ADODB.Command
        Set cnMySQL = New ADODB.Connection
        cnMySQL.ConnectionString = AdoStrCnnMySQL
        cnMySQL.open
        cnMySQL.CursorLocation = adUseClient
''        If cnMySQL.State = 0 Then
''           MsgBox "error coneccion"
''        Else
''            MsgBox "OK coneccion"
''        End If

        
        Set cmdMySQL.ActiveConnection = cnMySQL
        cmdMySQL.CommandType = adCmdText

       'Parametros de entrada y de salida
        ParametrosDeSalida = "@FechaCO, @FechaCE, @FechaDB, @FechaP12, @AgenteRetencion, @MicroEmpresa, @EstadoEmpresa, " _
                           & "@DescripcionEstado, @NombreEntidad, @Representante, @MensajeEmpresa, @ComunicadoEntidad, @SerieFA, " _
                           & "@TotCartera, @CantFA, @TipoPlan, @pActivo, @EstadoUsuario, @TokenEmpresa, @URLEmpresa"
                           
        cmdMySQL.CommandText = "Call sp_mysql_datos_iniciales_entidad(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?," & ParametrosDeSalida & ");"
       'Enviamos los parametro de solo entrada al SP
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("ItemEmpresa", adVarChar, adParamInput, 3, NumEmpresa)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("RUCEmpresa", adVarChar, adParamInput, 13, RUC)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("CodigoUsuario", adVarChar, adParamInput, 10, CodigoUsuario)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("NombreUsuario", adVarChar, adParamInput, 60, NombreUsuario)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("IDEUsuario", adVarChar, adParamInput, 15, IDEUsuario)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("PWRUsuario", adVarChar, adParamInput, 10, PWRUsuario)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("NombreEmpresa", adVarChar, adParamInput, 100, UCaseStrg(Empresa))
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("RazonSocialEmpresa", adVarChar, adParamInput, 120, RazonSocial)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("NombreCiudad", adVarChar, adParamInput, 35, NombreCiudad)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("ContadorEmpresa", adVarChar, adParamInput, 60, NombreContador)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("ContadorRUC", adVarChar, adParamInput, 13, RUC_Contador)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("GerenteEmpresa", adVarChar, adParamInput, 60, NombreGerente)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("NLogoTipo", adVarChar, adParamInput, 10, NLogoTipo)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("NMarcaAgua", adVarChar, adParamInput, 10, NMarcaAgua)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("EmailUsuario", adVarChar, adParamInput, 60, EmailUsuario)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("NivelesDeAccesos", adVarChar, adParamInput, 32768, CadenaParcial)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("IP_Local", adVarChar, adParamInput, 15, IP_PC.IP_PC)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("IP_WAN", adVarChar, adParamInput, 15, IP_PC.WAN_PC)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("PC_Nombre", adVarChar, adParamInput, 15, IP_PC.Nombre_PC)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("PC_MAC", adVarChar, adParamInput, 17, IP_PC.MAC_PC)
        
       'Ejecucion del SP en MySQL
        Set rsMySQL = cmdMySQL.Execute
       'Recolectamos los resultados de los parametros de salida
        Set rsMySQL = cnMySQL.Execute("SELECT " & ParametrosDeSalida & ";")
        Fecha_CO = Format(rsMySQL.fields(0), FormatoFechas)
        Fecha_CE = Format(rsMySQL.fields(1), FormatoFechas)
        Fecha_DB = Format(rsMySQL.fields(2), FormatoFechas)
        Fecha_P12 = Format(rsMySQL.fields(3), FormatoFechas)
        AgenteRetencion = rsMySQL.fields(4)
        MicroEmpresa = rsMySQL.fields(5)
        EstadoEmpresa = rsMySQL.fields(6)
        DescripcionEstado = rsMySQL.fields(7)
        NombreEntidad = rsMySQL.fields(8)
        RepresentanteLegal = rsMySQL.fields(9)
        MensajeEmpresa = rsMySQL.fields(10)
        ComunicadoEntidad = rsMySQL.fields(11)
        SerieFE = rsMySQL.fields(12)
        Cartera = rsMySQL.fields(13)
        Cant_FA = rsMySQL.fields(14)
        TipoPlan = rsMySQL.fields(15)
        PCActivo = rsMySQL.fields(16)
        EstadoUsuario = rsMySQL.fields(17)
        Token = rsMySQL.fields(18)
        URLToken = rsMySQL.fields(19)
        ServidorMySQL = True
        ParametrosDeSalida = ""
        
        For I = 0 To 19
            ParametrosDeSalida = ParametrosDeSalida & rsMySQL.fields(I).Name & " = " & rsMySQL.fields(I) & vbCrLf
        Next I
       'Cerramos la conexion con MySQL
        rsMySQL.Close
        cnMySQL.Close
         
       'Liberando de la memoria es control de conexion
        Set cmdMySQL.ActiveConnection = Nothing
        Set rsMySQL = Nothing
        Set cnMySQL = Nothing
        Set cmdMySQL = Nothing
        RatonNormal
       'MsgBox ParametrosDeSalida & ".-.-.-.-.-.-.-.-"
    End If
End Sub

Public Sub Estado_Empresa_SP_MySQL()
Dim cnMySQL As ADODB.Connection
Dim rsMySQL As ADODB.Recordset
Dim cmdMySQL As ADODB.Command
Dim ParametrosDeSalida As String

   'Conexion a MySQL del servidor en las nubes
   'Control_Procesos Normal, "Conexion MySQL Estado Empresa"
    RatonReloj
    ServidorMySQL = False
    PCActivo = True
    EstadoUsuario = True
    IDEntidad = 0
    DescripcionEstado = "OK"
    EstadoEmpresa = Ninguno
    Fecha_CE = FechaSistema
    Fecha_CO = FechaSistema
    Fecha_VPN = FechaSistema
    Fecha_DB = FechaSistema
    Fecha_P12 = FechaSistema
    SerieFE = Ninguno
    MicroEmpresa = Ninguno
    AgenteRetencion = Ninguno
    Cartera = 0
    Cant_FA = 0
    TipoPlan = 0
     If Ping_IP("db.diskcoversystem.com") Then
       'Conexion a MySQL del servidor en las nubes
        Set cmdMySQL = New ADODB.Command
        Set cnMySQL = New ADODB.Connection
        cnMySQL.ConnectionString = AdoStrCnnMySQL
        cnMySQL.open
        Set cmdMySQL.ActiveConnection = cnMySQL
        cmdMySQL.CommandType = adCmdText
       
       'Parametros de entrada y de salida
        ParametrosDeSalida = "@FechaCO, @FechaCE, @FechaVPN, @FechaDB, @FechaP12, @AgenteRetencion, @MicroEmpresa, @EstadoEmpresa, " _
                           & "@DescripcionEstado, @NombreEntidad, @Representante, @MensajeEmpresa, @ComunicadoEntidad, @TotCartera, " _
                           & "@CantFA, @TipoPlan, @SerieFA, @pActivo, @EstadoUsuario"
        cmdMySQL.CommandText = "Call sp_mysql_datos_estado_empresa(?, ?, ?, ?, ?, ?, ?," & ParametrosDeSalida & ");"
      
       'Enviamos los parametro de solo entrada al SP
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("ItemEmpresa", adVarChar, adParamInput, 3, NumEmpresa)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("RUCEmpresa", adVarChar, adParamInput, 13, RUC)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("CodigoUsuario", adVarChar, adParamInput, 10, CodigoUsuario)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("IP_Local", adVarChar, adParamInput, 15, IP_PC.IP_PC)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("IP_WAN", adVarChar, adParamInput, 15, IP_PC.WAN_PC)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("PC_Nombre", adVarChar, adParamInput, 15, IP_PC.Nombre_PC)
        cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("PC_MAC", adVarChar, adParamInput, 17, IP_PC.MAC_PC)

       'Ejecucion del SP en MySQL
        Set rsMySQL = cmdMySQL.Execute
       
       'Recolectamos los resultados de los parametros de salida
        Set rsMySQL = cnMySQL.Execute("SELECT " & ParametrosDeSalida & ";")
       
       'Pasamos a variables globales lso resultados del SP
        If Not rsMySQL.EOF Then
           Fecha_CO = Format(rsMySQL.fields(0), FormatoFechas)
           Fecha_CE = Format(rsMySQL.fields(1), FormatoFechas)
           Fecha_VPN = Format(rsMySQL.fields(2), FormatoFechas)
           Fecha_DB = Format(rsMySQL.fields(3), FormatoFechas)
           Fecha_P12 = Format(rsMySQL.fields(4), FormatoFechas)
         
           AgenteRetencion = rsMySQL.fields(5)
           MicroEmpresa = rsMySQL.fields(6)
           EstadoEmpresa = rsMySQL.fields(7)
           DescripcionEstado = rsMySQL.fields(8)
           NombreEntidad = rsMySQL.fields(9)
           RepresentanteLegal = rsMySQL.fields(10)
           MensajeEmpresa = rsMySQL.fields(11)
           ComunicadoEntidad = rsMySQL.fields(12)
           Cartera = rsMySQL.fields(13)
           Cant_FA = rsMySQL.fields(14)
           TipoPlan = rsMySQL.fields(15)
           SerieFE = rsMySQL.fields(16)
           PCActivo = rsMySQL.fields(17)
           EstadoUsuario = rsMySQL.fields(18)
           ServidorMySQL = True
        End If
        
       'Cerramos la conexion con MySQL
        rsMySQL.Close
        cnMySQL.Close
     End If
     
    'Set cmdMySQL.ActiveConnection = Nothing
     Set rsMySQL = Nothing
     Set cnMySQL = Nothing
     Set cmdMySQL = Nothing
    If Len(AgenteRetencion) > 1 Then AgenteRetencion = " Resolución: " & AgenteRetencion
    RatonNormal
End Sub

Public Sub Tipo_Contribuyente_SP_MySQL(RUCContribuyente As String, vMicroEmpresa As String, vAgenteRetencion As String)
Dim cnMySQL As ADODB.Connection
Dim rsMySQL As ADODB.Recordset
Dim cmdMySQL As ADODB.Command

  vMicroEmpresa = Ninguno
  vAgenteRetencion = Ninguno
  If Ping_IP("db.diskcoversystem.com") Then
    If Len(RUCContribuyente) = 13 Then
      'Conexion a MySQL del servidor en las nubes
      'Control_Procesos Normal, "Conexion MySQL Tipo Contribuyente"
       RatonReloj
      'Conexion a MySQL del servidor en las nubes
       Set cmdMySQL = New ADODB.Command
       Set cnMySQL = New ADODB.Connection
       cnMySQL.ConnectionString = AdoStrCnnMySQL
       cnMySQL.open
        
       Set cmdMySQL.ActiveConnection = cnMySQL
       cmdMySQL.CommandType = adCmdText
        
      'Parametros de entrada y de salida
       cmdMySQL.CommandText = "Call sp_tipo_contribuyente(?, @AgenteRetencion, @MicroEmpresa);"
                             
      'Enviamos los parametro de solo entrada al SP
       cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("RUC", adVarChar, adParamInput, 13, RUCContribuyente)
        
      'Ejecucion del SP en MySQL
       Set rsMySQL = cmdMySQL.Execute
        
      'Recolectamos los resultados de los parametros de salida
       Set rsMySQL = cnMySQL.Execute("SELECT @AgenteRetencion, @MicroEmpresa;")
        
      'Pasamos a variables globales lso resultados del SP
       If Not rsMySQL.EOF Then
          vAgenteRetencion = rsMySQL.fields(0)
          vMicroEmpresa = rsMySQL.fields(1)
       End If
      'Cerramos la conexion con MySQL
       rsMySQL.Close
       cnMySQL.Close
       
       'Set cmdMySQL.ActiveConnection = Nothing
       Set rsMySQL = Nothing
       Set cnMySQL = Nothing
       Set cmdMySQL = Nothing
       RatonNormal
    End If
  End If
End Sub

Public Sub Leer_URL_Token_SP_MySQL(RUCEmpresa As String, vURL As String, vToken As String)
Dim cnMySQL As ADODB.Connection
Dim rsMySQL As ADODB.Recordset
Dim cmdMySQL As ADODB.Command

  vToken = Ninguno
  vURL = Ninguno
  If Ping_IP("db.diskcoversystem.com") And Len(RUCEmpresa) = 13 Then
    'Conexion a MySQL del servidor en las nubes
    'Control_Procesos Normal, "Conexion MySQL Tipo Contribuyente"
     RatonReloj
    'Conexion a MySQL del servidor en las nubes
     Set cmdMySQL = New ADODB.Command
     Set cnMySQL = New ADODB.Connection
     cnMySQL.ConnectionString = AdoStrCnnMySQL
     cnMySQL.open
        
     Set cmdMySQL.ActiveConnection = cnMySQL
     cmdMySQL.CommandType = adCmdText
        
    'Parametros de entrada y de salida
       cmdMySQL.CommandText = "Call sp_leer_url_token(?, @URLEmpresa, @TokenEmpresa);"
                             
      'Enviamos los parametro de solo entrada al SP
       cmdMySQL.Parameters.Append cmdMySQL.CreateParameter("RUCEmpresa", adVarChar, adParamInput, 13, RUCEmpresa)
        
      'Ejecucion del SP en MySQL
       Set rsMySQL = cmdMySQL.Execute
        
      'Recolectamos los resultados de los parametros de salida
       Set rsMySQL = cnMySQL.Execute("SELECT @URLEmpresa, @TokenEmpresa;")
        
      'Pasamos a variables globales lso resultados del SP
       If Not rsMySQL.EOF Then
          vURL = rsMySQL.fields(0)
          vToken = rsMySQL.fields(1)
       End If
      'Cerramos la conexion con MySQL
       rsMySQL.Close
       cnMySQL.Close
       
       'Set cmdMySQL.ActiveConnection = Nothing
       Set rsMySQL = Nothing
       Set cnMySQL = Nothing
       Set cmdMySQL = Nothing
       RatonNormal
  End If
End Sub

'"Exec NombredelProcedimientoAlmacenado " & parametro1 & ",'" & parametro2 & "'," & parametro3
Public Sub Ejecutar_SP(StoredProcedure As String, Parameters As String)
Dim AdoReg As ADODB.Recordset
Dim SQLQuery As String
  RatonReloj
  SQLQuery = "EXEC " & StoredProcedure & " " & Parameters & " "
  Set AdoReg = New ADODB.Recordset
 'SQLQuery = CompilarSQL(SQLQuery)
 'MsgBox SQLQuery
 'adLockReadOnly
  If SQL_Server Then
     AdoReg.open SQLQuery, AdoStrCnn, , adOpenKeyset, adLockReadOnly
  Else
     AdoReg.open SQLQuery, AdoStrCnnMySQL, , adOpenKeyset, adLockReadOnly
  End If
 'AdoReg.Close
  RatonNormal
End Sub

Public Sub Crear_SP_FN(StoredProcedure As String, Optional NombreFile As String)
Dim AdoReg As ADODB.Recordset
Dim SQLQuery As String
  RatonReloj
  SQLQuery = "EXECUTE(" & StoredProcedure & ");"
  Generar_File_SQL NombreFile, SQLQuery
  
  Set AdoReg = New ADODB.Recordset
 'MsgBox SQLQuery
 'adLockReadOnly
  If SQL_Server Then
     AdoReg.open SQLQuery, AdoStrCnn, , adOpenKeyset, adLockReadOnly
  Else
     AdoReg.open SQLQuery, AdoStrCnnMySQL, , adOpenKeyset, adLockReadOnly
  End If
 'AdoReg.Close
  RatonNormal
End Sub

Public Sub Ejecutar_SQL_SP(SQL As String, Optional NoCompilar As Boolean, Optional NombreFile As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
Dim TByte As Long
    'MsgBox SQL
    If Not NoCompilar Then SQL = CompilarSQL(SQL)
    Generar_File_SQL NombreFile, SQL
    Iniciar_Stored_Procedure "Ejecucion SP con parametros", MiSQL, MiCmd, MiReg
   'MsgBox Len(SQL) & vbCrLf & SQL
    TByte = Len(SQL) + 10
    MiCmd.CommandText = "sp_Ejecutar_SQL"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@sSQL", adVarChar, adParamInput, TByte, SQL)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Ejecutar_SQL_AdoDB(SQLQuery As String, Optional NoCompilar As Boolean, Optional NombreFile As String)
Dim AdoCon1 As ADODB.Connection
  RatonReloj
  If Len(SQLQuery) > 1 Then
     Set AdoCon1 = New ADODB.Connection
     If Not NoCompilar Then SQLQuery = CompilarSQL(SQLQuery)
     NombreFile = "Archivo_" & Format(Time, "hh-mm-ss") & ".sql"
     Generar_File_SQL NombreFile, SQLQuery
    'MsgBox SQLQuery & vbCrLf & String(70, "_") & vbCrLf & AdoStrCnn
     If SQL_Server Then AdoCon1.open AdoStrCnn Else AdoCon1.open AdoStrCnnMySQL
     AdoCon1.Execute SQLQuery, RegAfectados, adCmdText
     AdoCon1.Close
  End If
  RatonNormal
 'If RegSN Then MsgBox "Registros Afectados: " & Format$(RegAfectados, "#,##0")
End Sub

Public Sub Eliminar_Empresa_SP(Item As String, NombreEmpresa As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
    Iniciar_Stored_Procedure "Eliminando (" & Item & "): " & NombreEmpresa, MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Eliminar_Empresa"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, Item)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Eliminar_Indices_SP()
Dim AdoReg As ADODB.Recordset
Dim SQLQuery As String
  RatonReloj
  'FEsperar.Show
  'Imagen_Esperar "Eliminacion de Indices"
  SQLQuery = "EXEC sp_Eliminar_Indices; "
  Set AdoReg = New ADODB.Recordset
  If SQL_Server Then
     AdoReg.open SQLQuery, AdoStrCnn, , adOpenKeyset, adLockReadOnly
  Else
     AdoReg.open SQLQuery, AdoStrCnnMySQL, , adOpenKeyset, adLockReadOnly
  End If
 'AdoReg.open SQLQuery, AdoStrCnn, , adOpenKeyset, adLockReadOnly
  RatonNormal
  'Unload FEsperar
End Sub

Public Sub Insertar_Ctas_Cierre_SP(InsCta As String, Valor As Currency)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
    If Len(InsCta) > 1 And Valor <> 0 Then
       Iniciar_Stored_Procedure "Insertar Ctas Cierre", MiSQL, MiCmd, MiReg
       MiCmd.CommandText = "sp_Insertar_Ctas_Cierre"
       MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
       MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
       MiCmd.Parameters.Append MiCmd.CreateParameter("@NumModulo", adVarChar, adParamInput, 2, NumModulo)
       MiCmd.Parameters.Append MiCmd.CreateParameter("@Usuario", adVarChar, adParamInput, 10, CodigoUsuario)
       MiCmd.Parameters.Append MiCmd.CreateParameter("@Codigo", adVarChar, adParamInput, 18, InsCta)
       MiCmd.Parameters.Append MiCmd.CreateParameter("@Valor", adCurrency, adParamInput, 16, Valor)
       MiCmd.Parameters.Append MiCmd.CreateParameter("@TransNo", adInteger, adParamInput, 14, Trans_No)
       Procesar_Stored_Procedure MiCmd, MiReg
       Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
    End If
End Sub

Public Sub Actualizar_Base_Datos_SP()

Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

    Iniciar_Stored_Procedure "Actualizar Base Datos", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Actualizar_Base_Datos"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@strIPServidor", adVarChar, adParamInput, 50, strIPServidor)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@ruta_file", adVarChar, adParamInput, 256, RutaSistema & "\BASES\UPDATE_DB\")
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Actualizar_SP_FN_SP()

Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

    Iniciar_Stored_Procedure "Actualizar SP FN", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Actualizar_SP_FN"
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Mayorizar_Cuentas_SP()
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

    Iniciar_Stored_Procedure "Mayorizar Cuentas", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Mayorizar_Cuentas"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Actualizar_Datos_Representantes_SP(Optional MasGrupos As Boolean)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
Dim BuscarCodigo1 As String

    Iniciar_Stored_Procedure "Cargando Datos para proceso con los Bancos", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Actualizar_Datos_Representantes"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@MasGrupos", adBoolean, adParamInput, 1, MasGrupos)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Actualizar_Tipo_Clientes_SP()
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
Dim BuscarCodigo1 As String

    Iniciar_Stored_Procedure "Actualizar Tipo Clientes", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Actualizar_Tipo_Clientes"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Actualiza_Transacciones_Kardex_SP()
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
    
    Iniciar_Stored_Procedure "Actualiza Transacciones Kardex", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Actualiza_Transacciones_Kardex"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@EsCoop", adBoolean, adParamInput, 1, OpcCoop)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@ConSucursal", adBoolean, adParamInput, 1, ConSucursal)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@CodigoUsuario", adVarChar, adParamInput, 10, CodigoUsuario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Modulo", adVarChar, adParamInput, 2, NumModulo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Dec_PVP", adInteger, adParamInput, 14, Dec_PVP)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Dec_Costo", adInteger, adParamInput, 14, Dec_Costo)
    
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Presenta_Errores_Contabilidad_SP()
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

    Iniciar_Stored_Procedure "Mayorizar Cuentas", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Presenta_Errores_Contabilidad"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Usuario", adVarChar, adParamInput, 10, CodigoUsuario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NumModulo", adVarChar, adParamInput, 2, NumModulo)
    'MiCmd.Parameters.Append MiCmd.CreateParameter("@ExisteErrores", adBoolean, adParamOutput, 1, ExisteErrores)
    Procesar_Stored_Procedure MiCmd, MiReg
    'ExisteErrores = MiCmd.Parameters("@ExisteErrores").value
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
    'If ExisteErrores Then
    FInfoError.Show
End Sub

Public Sub Presenta_Errores_Facturacion_SP(FechaDesde As MaskEdBox, FechaHasta As MaskEdBox)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

Dim FechaIniSP As String
Dim FechaFinSP As String
    
    FechaIniSP = BuscarFecha(FechaDesde.Text)
    FechaFinSP = BuscarFecha(FechaHasta.Text)
    Iniciar_Stored_Procedure "Errores de Facturacion", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Presenta_Errores_Facturacion"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaDesde", adVarChar, adParamInput, 10, FechaIniSP)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaHasta", adVarChar, adParamInput, 10, FechaFinSP)
    'MiCmd.Parameters.Append MiCmd.CreateParameter("@DecCosto", adUnsignedTinyInt, adParamInput, 1, Dec_Costo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Usuario", adVarChar, adParamInput, 10, CodigoUsuario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NumModulo", adVarChar, adParamInput, 2, NumModulo)
    'MiCmd.Parameters.Append MiCmd.CreateParameter("@ExisteErrores", adBoolean, adParamOutput, 1, ExisteErrores)
    Procesar_Stored_Procedure MiCmd, MiReg
    'ExisteErrores = MiCmd.Parameters("@ExisteErrores").value
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Grabar_Facturas_x_Lotes_SP(CodigoCxC As String, _
                                      GrupoINo As String, _
                                      GrupoFNo As String, _
                                      FechaDesde As MaskEdBox, _
                                      FechaHasta As MaskEdBox, _
                                      FechaFacturar As MaskEdBox, _
                                      NoMes As Integer, _
                                      AnioFA As String, _
                                      Tipo_Pago As String, _
                                      Nota As String, _
                                      Observacion As String, _
                                      PorGrupo As Boolean, _
                                      CheqRangos As Boolean, _
                                      CheqFA As Boolean)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

Dim FechaIniSP As String
Dim FechaFinSP As String
Dim FechaFacSP As String
    
    FechaIniSP = BuscarFecha(FechaDesde.Text)
    FechaFinSP = BuscarFecha(FechaHasta.Text)
    FechaFacSP = BuscarFecha(FechaFacturar.Text)
    Iniciar_Stored_Procedure "Generacion de Facturas en Bloque", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Grabar_Facturas_x_Lotes"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NumModulo", adVarChar, adParamInput, 2, NumModulo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Usuario", adVarChar, adParamInput, 10, CodigoUsuario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@GrupoIni", adVarChar, adParamInput, 10, GrupoINo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@GrupoFin", adVarChar, adParamInput, 10, GrupoFNo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaDesde", adVarChar, adParamInput, 10, FechaIniSP)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaHasta", adVarChar, adParamInput, 10, FechaFinSP)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaFacturar", adVarChar, adParamInput, 10, FechaFacSP)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@CodigoCxC", adVarChar, adParamInput, 10, CodigoCxC)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NoMes", adInteger, adParamInput, 10, NoMes)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@AnioFA", adVarChar, adParamInput, 10, AnioFA)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Tipo_Pago", adVarChar, adParamInput, 2, Tipo_Pago)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Nota", adVarChar, adParamInput, 100, Nota)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Observacion", adVarChar, adParamInput, 100, Observacion)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@PorGrupo", adBoolean, adParamInput, 1, PorGrupo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@CheqRangos", adBoolean, adParamInput, 1, CheqRangos)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@CheqFA", adBoolean, adParamInput, 1, CheqFA)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Digito_Verificador_SP(NumeroRUC As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

Dim RUCCI As String
Dim CodigoRUCCI As String
Dim DigitoVerificador As String
Dim TipoBeneficiario As String
Dim RUCNatural As Boolean
    
   'Determinamos que tipo de RUC/CI es
    Tipo_RUC_CI.Tipo_Beneficiario = "P"
    Tipo_RUC_CI.Codigo_RUC_CI = NumEmpresa & "0000001"
    Tipo_RUC_CI.Digito_Verificador = "-"
    Tipo_RUC_CI.RUC_CI = NumeroRUC
    Tipo_RUC_CI.RUC_Natural = False
    TipoSRI.Existe = False
  
    Iniciar_Stored_Procedure "Digito Verificador", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Digito_Verificador"
    
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NumeroRUC", adVarChar, adParamInput, 15, NumeroRUC)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@RUCCI", adVarChar, adParamOutput, 15, RUCCI)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@CodigoRUCCI", adVarChar, adParamOutput, 10, CodigoRUCCI)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@DigitoVerificador", adVarChar, adParamOutput, 1, DigitoVerificador)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@TipoBeneficiario", adVarChar, adParamOutput, 1, TipoBeneficiario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@RUCNatural", adBoolean, adParamOutput, 1, RUCNatural)
    Procesar_Stored_Procedure MiCmd, MiReg
   'Recolectamos los resultados del digito verificador
    Tipo_RUC_CI.Digito_Verificador = MiCmd.Parameters("@DigitoVerificador").value
    Tipo_RUC_CI.Tipo_Beneficiario = MiCmd.Parameters("@TipoBeneficiario").value
    Tipo_RUC_CI.Codigo_RUC_CI = MiCmd.Parameters("@CodigoRUCCI").value
    Tipo_RUC_CI.RUC_Natural = MiCmd.Parameters("@RUCNatural").value
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
    If Tipo_RUC_CI.Tipo_Beneficiario <> "R" Then TipoSRI.Existe = False
    'MsgBox Tipo_RUC_CI.Tipo_Beneficiario
End Sub

Public Sub Actualizar_Abonos_Facturas_SP(TFA As Tipo_Facturas, _
                                         Optional SaldoReal As Boolean, _
                                         Optional PorFecha As Boolean)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

''''facturacion2@fidal-amlat.org    Fidal2022
''''facturacion1@asproduc.com   Fidal2022

   'Determinamos formatos de los parametros de entrada
   'MsgBox "=> " & TFA.Fecha_Corte & vbCrLf & TFA.Fecha_Desde & vbCrLf & TFA.Fecha_Hasta
    FechaCorte = TFA.Fecha_Corte
    FechaIni = TFA.Fecha_Desde
    FechaFin = TFA.Fecha_Hasta
    
    If IsDate(FechaCorte) And Len(FechaCorte) = 10 Then FechaCorte = BuscarFecha(FechaCorte) Else FechaCorte = BuscarFecha(FechaSistema)
    If IsDate(FechaIni) And Len(FechaIni) = 10 Then FechaIni = BuscarFecha(FechaIni) Else FechaIni = BuscarFecha(FechaSistema)
    If IsDate(FechaFin) And Len(FechaFin) = 10 Then FechaFin = BuscarFecha(FechaFin) Else FechaFin = BuscarFecha(FechaSistema)
    If FechaCorte = BuscarFecha(FechaSistema) Then SaldoReal = True
    
   'MsgBox "==> " & TFA.Fecha_Corte & vbCrLf & TFA.Fecha_Desde & vbCrLf & TFA.Fecha_Hasta & vbCrLf & SaldoReal & vbCrLf & PorFecha
    
    Iniciar_Stored_Procedure "Actualizar Abonos Facturas", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Actualizar_Abonos_Facturas"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NumModulo", adVarChar, adParamInput, 2, NumModulo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Usuario", adVarChar, adParamInput, 10, CodigoUsuario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@TC", adVarChar, adParamInput, 2, TFA.TC)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Serie", adVarChar, adParamInput, 6, TFA.Serie)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Factura", adInteger, adParamInput, 14, TFA.Factura)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaCorte", adVarChar, adParamInput, 10, FechaCorte)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaDesde", adVarChar, adParamInput, 10, FechaIni)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaHasta", adVarChar, adParamInput, 10, FechaFin)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@SaldoReal", adBoolean, adParamInput, 1, SaldoReal)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@PorFecha", adBoolean, adParamInput, 1, PorFecha)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@ExisteErrores", adBoolean, adParamOutput, 1, ExisteErrores)
    Procesar_Stored_Procedure MiCmd, MiReg
    ExisteErrores = MiCmd.Parameters("@ExisteErrores").value
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Actualizar_Saldos_Facturas_SP(dTC As String, dSerie As String, dFactura As Long)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
    
    Iniciar_Stored_Procedure "Actualizar Saldos Facturas", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Actualizar_Saldos_Facturas"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@TC", adVarChar, adParamInput, 2, dTC)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Serie", adVarChar, adParamInput, 6, dSerie)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Factura", adInteger, adParamInput, 14, dFactura)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Mayorizar_Inventario_SP(Optional FechaCorte As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
Dim FechaCorteKardex As String

    If Len(FechaCorte) >= 8 Then FechaCorteKardex = BuscarFecha(FechaCorte) Else FechaCorteKardex = BuscarFecha(FechaSistema)
    Iniciar_Stored_Procedure "Mayorizar Inventario", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Mayorizar_Inventario"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Usuario", adVarChar, adParamInput, 10, CodigoUsuario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Modulo", adVarChar, adParamInput, 2, NumModulo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaCorte", adVarChar, adParamInput, 10, FechaCorteKardex)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@TipoKardex", adVarChar, adParamOutput, 6, TipoKardex)
    Procesar_Stored_Procedure MiCmd, MiReg
    TipoKardex = MiCmd.Parameters("@TipoKardex").value
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Eliminar_Nulos_SP(NombreTabla As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

    Iniciar_Stored_Procedure "Eliminacion nulos de " & NombreTabla, MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Eliminar_Nulos"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NombreTabla", adVarChar, adParamInput, 50, NombreTabla)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Copiar_Tabla_SP(NombreTabla As String, ItemOld As String, ItemNew As String, PeriodoOld As String, PeriodoNew As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

    Iniciar_Stored_Procedure "Copiar Tabla: " & NombreTabla, MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Copiar_Tabla"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NombreTabla", adVarChar, adParamInput, 50, NombreTabla)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@ItemOld", adVarChar, adParamInput, 3, ItemOld)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@ItemNew", adVarChar, adParamInput, 3, ItemNew)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@PeriodoOld", adVarChar, adParamInput, 10, PeriodoOld)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@PeriodoNew", adVarChar, adParamInput, 10, PeriodoNew)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Update_Default_SP(NombreTabla As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
    Iniciar_Stored_Procedure "Actualizar datos de " & NombreTabla, MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Update_Default"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@nombreTabla", adVarChar, adParamInput, 50, NombreTabla)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Productos_Cierre_Caja_SP(FechaDesde As MaskEdBox, FechaHasta As MaskEdBox)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
Dim FechaIniSP As String
Dim FechaFinSP As String

    FechaIniSP = BuscarFecha(FechaDesde.Text)
    FechaFinSP = BuscarFecha(FechaHasta.Text)
    Iniciar_Stored_Procedure "Cierre Diario de Caja", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Productos_Cierre_Caja"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaDesde", adVarChar, adParamInput, 10, FechaIniSP)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaHasta", adVarChar, adParamInput, 10, FechaFinSP)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Insertar_Productos_Cierre_Caja_SP(FechaDesde As MaskEdBox, FechaHasta As MaskEdBox)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
Dim FechaIniSP As String
Dim FechaFinSP As String

    FechaIniSP = BuscarFecha(FechaDesde.Text)
    FechaFinSP = BuscarFecha(FechaHasta.Text)
    Iniciar_Stored_Procedure "Insertar Productos Cierre Caja", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Insertar_Productos_Cierre_Caja"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaDesde", adVarChar, adParamInput, 10, FechaIniSP)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaHasta", adVarChar, adParamInput, 10, FechaFinSP)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Reporte_CxC_Cuotas_SP(GrupoINo As String, _
                                 GrupoFNo As String, _
                                 MBFechaInicial As String, _
                                 MBFechaCorte As String, _
                                 SubTotal As Currency, _
                                 TotalAnticipo As Currency, _
                                 TotalCxC As Currency, _
                                 ListaDeCampos As String, _
                                 Resumido As Boolean, _
                                 Vencimiento As Boolean)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
Dim EjercicioFiscal As String
    
    FechaIni = BuscarFecha(MBFechaInicial)
    FechaFin = BuscarFecha(MBFechaCorte)
    EjercicioFiscal = Year(MBFechaCorte)
    GrupoINo = TrimStrg(MidStrg(GrupoINo, 1, 10))
    GrupoFNo = TrimStrg(MidStrg(GrupoFNo, 1, 10))
    If Vencimiento Then FechaIni = FechaFin
    
    Iniciar_Stored_Procedure "Reporte CxC Cuotas Pre-facturacion", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Reporte_CxC_Cuotas"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@CodigoUsuario", adVarChar, adParamInput, 10, CodigoUsuario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@EjercicioFiscal", adVarChar, adParamInput, 4, EjercicioFiscal)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaInicio", adVarChar, adParamInput, 10, FechaIni)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaCorte", adVarChar, adParamInput, 10, FechaFin)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@GrupoNo", adVarChar, adParamInput, 10, GrupoINo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@GrupoNo", adVarChar, adParamInput, 10, GrupoFNo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Resumido", adBoolean, adParamInput, 1, Resumido)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@SubTotal", adCurrency, adParamOutput, 16, SubTotal)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@TotalAnticipo", adCurrency, adParamOutput, 16, TotalAnticipo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@TotalCxC", adCurrency, adParamOutput, 16, TotalCxC)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@ListaCampos", adVarChar, adParamOutput, 2048, ListaDeCampos)
    Procesar_Stored_Procedure MiCmd, MiReg
   'Recojemos los datos salientes de los campos que retorna valor el store procedure del SQL Server
    SubTotal = MiCmd.Parameters("@SubTotal").value
    TotalAnticipo = MiCmd.Parameters("@TotalAnticipo").value
    TotalCxC = MiCmd.Parameters("@TotalCxC").value
    ListaDeCampos = MiCmd.Parameters("@ListaCampos").value
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Reporte_Resumen_Existencias_SP(MBFechaInicial As String, _
                                          MBFechaFinal As String, _
                                          CodigoBodega As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

    If CFechaLong(MBFechaInicial) <= CFechaLong(MBFechaFinal) Then
       FechaIni = BuscarFecha(MBFechaInicial)
       FechaFin = BuscarFecha(MBFechaFinal)
       Iniciar_Stored_Procedure "Reporte Resumen Existencias", MiSQL, MiCmd, MiReg
       MiCmd.CommandText = "sp_Reporte_Resumen_Existencias"
       MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
       MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
       MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaInicial", adVarChar, adParamInput, 10, FechaIni)
       MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaFinal", adVarChar, adParamInput, 10, FechaFin)
       MiCmd.Parameters.Append MiCmd.CreateParameter("@CodBod", adVarChar, adParamInput, 2, CodigoBodega)
       Procesar_Stored_Procedure MiCmd, MiReg
       Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
    End If
End Sub

Public Sub Reporte_Cartera_Clientes_SP(MBFechaInicial As String, _
                                       MBFechaFinal As String, _
                                       CodigoCliente As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

    If CFechaLong(MBFechaInicial) <= CFechaLong(MBFechaFinal) Then
       Iniciar_Stored_Procedure "Reporte Cartera Clientes", MiSQL, MiCmd, MiReg
       MiCmd.CommandText = "sp_Reporte_Cartera_Clientes"
       MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
       MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
       MiCmd.Parameters.Append MiCmd.CreateParameter("@CodigoUsuario", adVarChar, adParamInput, 10, CodigoUsuario)
       MiCmd.Parameters.Append MiCmd.CreateParameter("@CodigoCliente", adVarChar, adParamInput, 10, CodigoCliente)
       MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaInicio", adVarChar, adParamInput, 10, BuscarFecha(MBFechaInicial))
       MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaCorte", adVarChar, adParamInput, 10, BuscarFecha(MBFechaFinal))
       Procesar_Stored_Procedure MiCmd, MiReg
       Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
    End If
End Sub

'DECLARE @ListaCampos VARCHAR(MAX), @SumatoriaCampos VARCHAR(MAX)
'EXEC sp_Reporte_Rol_Pagos_Colectivo '106', '.', '0702164179', '20220501', '20220531', 'Todos', 1,@ListaCampos OUTPUT, @SumatoriaCampos OUTPUT
Public Sub Reporte_Rol_Pagos_Colectivo_SP(FechaIniRol As String, _
                                          FechaFinRol As String, _
                                          GrupoRol As String, _
                                          OrdenAlfabetico As Boolean, _
                                          ListaDeCampos As String, _
                                          SumaDeCampos As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
    Iniciar_Stored_Procedure "Generacion Reporte Rol Pagos Colectivo", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Reporte_Rol_Pagos_Colectivo"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@CodigoUsuario", adVarChar, adParamInput, 10, CodigoUsuario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaIniRol", adVarChar, adParamInput, 10, FechaIniRol)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaFinRol", adVarChar, adParamInput, 10, FechaFinRol)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@GrupoRol", adVarChar, adParamInput, 15, GrupoRol)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@OrdenAlfabetico", adBoolean, adParamInput, 1, OrdenAlfabetico)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@ListaCampos", adVarChar, adParamOutput, 5120, ListaDeCampos)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@SumatoriaCampos", adVarChar, adParamOutput, 5120, SumaDeCampos)
    Procesar_Stored_Procedure MiCmd, MiReg
   'Recojemos los datos salientes de los campos que retorna valor el store procedure del SQL Server
    ListaDeCampos = MiCmd.Parameters("@ListaCampos").value
    SumaDeCampos = MiCmd.Parameters("@SumatoriaCampos").value
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

'''Public Sub Procesar_Rol_Pagos_Mensual_SP(FechaIniRol As String, _
'''                                         FechaFinRol As String, _
'''                                         GrupoRol As String, _
'''                                         CheqCxP As Boolean, _
'''                                         DCCxP As String)
'''Dim MiSQL As ADODB.Connection
'''Dim MiCmd As ADODB.Command
'''Dim MiReg As ADODB.Recordset
'''
'''    Iniciar_Stored_Procedure "Generacion Rol Pagos Colectivo Mensual", MiSQL, MiCmd, MiReg
'''    MiCmd.CommandText = "sp_Procesar_Rol_Pagos_Mensual"
'''    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
'''    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
'''    MiCmd.Parameters.Append MiCmd.CreateParameter("@NumModulo", adVarChar, adParamInput, 2, NumModulo)
'''    MiCmd.Parameters.Append MiCmd.CreateParameter("@CodigoUsuario", adVarChar, adParamInput, 10, CodigoUsuario)
'''    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaIniRol", adVarChar, adParamInput, 10, FechaIniRol)
'''    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaFinRol", adVarChar, adParamInput, 10, FechaFinRol)
'''    MiCmd.Parameters.Append MiCmd.CreateParameter("@GrupoRol", adVarChar, adParamInput, 15, GrupoRol)
'''    MiCmd.Parameters.Append MiCmd.CreateParameter("@DCCxP", adVarChar, adParamInput, 18, DCCxP)
'''    MiCmd.Parameters.Append MiCmd.CreateParameter("@CheqCxP", adBoolean, adParamInput, 1, CheqCxP)
'''    Procesar_Stored_Procedure MiCmd, MiReg
'''    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
'''End Sub

Public Sub Procesar_Rol_Pagos_del_Mes_SP(FechaIniRol As String, _
                                         FechaFinRol As String, _
                                         GrupoRol As String, _
                                         DCCxP As String, _
                                         No_Cheque As Long)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

    Iniciar_Stored_Procedure "Generacion Rol Pagos del Mes", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Procesar_Rol_Pagos_del_Mes"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NumModulo", adVarChar, adParamInput, 2, NumModulo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@CodigoUsuario", adVarChar, adParamInput, 10, CodigoUsuario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaIniRol", adVarChar, adParamInput, 10, FechaIniRol)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaFinRol", adVarChar, adParamInput, 10, FechaFinRol)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@GrupoRol", adVarChar, adParamInput, 15, GrupoRol)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@DCCxP", adVarChar, adParamInput, 18, DCCxP)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@No_Cheque", adInteger, adParamInput, 14, No_Cheque)
    
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Procesar_Rol_Pagos_Asientos_SP(FechaIniRol As String, _
                                          FechaFinRol As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

    Iniciar_Stored_Procedure "Generacion Rol Pagos del Mes", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Procesar_Rol_Pagos_Asientos"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NumModulo", adVarChar, adParamInput, 2, NumModulo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@CodigoUsuario", adVarChar, adParamInput, 10, CodigoUsuario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaIniRol", adVarChar, adParamInput, 10, FechaIniRol)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaFinRol", adVarChar, adParamInput, 10, FechaFinRol)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Procesar_Balance_Analitico_Mensual_SP(TipoBalance As String, ConSubModulos As Boolean, MBFechaI As String, MBFechaF As String, ListaDeCampos As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
    
    FechaIni = BuscarFecha(MBFechaI)
    FechaFin = BuscarFecha(MBFechaF)
    Iniciar_Stored_Procedure "Procesar Balance Analitico Mensual", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Procesar_Balance_Analitico_Mensual"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@TipoBalance", adVarChar, adParamInput, 2, TipoBalance)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@ConSubModulos", adBoolean, adParamInput, 1, ConSubModulos)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@CodigoUsuario", adVarChar, adParamInput, 10, CodigoUsuario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaDesde", adVarChar, adParamInput, 10, FechaIni)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaHasta", adVarChar, adParamInput, 10, FechaFin)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@ListaMeses", adVarChar, adParamOutput, 4096, ListaDeCampos)
    Procesar_Stored_Procedure MiCmd, MiReg
   'Recojemos los datos salientes de los campos que retorna valor el store procedure del SQL Server
    ListaDeCampos = MiCmd.Parameters("@ListaMeses").value
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Procesar_Balance_SP(EsBalanceMes As Boolean, MBFechaI As String, MBFechaF As String, TipoBalanceCC As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
    
    FechaIni = BuscarFecha(MBFechaI)
    FechaFin = BuscarFecha(MBFechaF)
    'MsgBox EsBalanceMes & " ..."
    Iniciar_Stored_Procedure "Procesar Balance de Comprobacion", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Procesar_Balance"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaDesde", adVarChar, adParamInput, 10, FechaIni)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaHasta", adVarChar, adParamInput, 10, FechaFin)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@EsCoop", adBoolean, adParamInput, 1, OpcCoop)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@ConSucursal", adBoolean, adParamInput, 1, ConSucursal)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@EsBalanceMes", adBoolean, adParamInput, 1, EsBalanceMes)
    'MiCmd.Parameters.Append MiCmd.CreateParameter("@CentroCostos", adVarChar, adParamInput, 10, TipoBalanceCC)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Duplicar_Tabla_SP(NombreTablaOrigen As String, NombreTablaDestino As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
    Iniciar_Stored_Procedure "Duplicacion de la tabla: " & NombreTablaOrigen, MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Duplicar_Tabla"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NombreTablaOrigen", adVarChar, adParamInput, 50, NombreTablaOrigen)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NombreTablaDestino", adVarChar, adParamInput, 50, NombreTablaDestino)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Procesar_Balance_Ext_SP()
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
Dim Errores As Boolean

     Errores = False
     Iniciar_Stored_Procedure "Procesando Balance de Comprobacion Externos", MiSQL, MiCmd, MiReg
     MiCmd.CommandText = "sp_Procesar_Balance_Ext"
     MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
     MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
     Procesar_Stored_Procedure MiCmd, MiReg
     Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Insertar_Texto_Temporal_SP(Texto As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

     Iniciar_Stored_Procedure "Insertar Texto Temporal", MiSQL, MiCmd, MiReg
     MiCmd.CommandText = "sp_Insertar_Texto_Temporal"
     MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
     MiCmd.Parameters.Append MiCmd.CreateParameter("@Usuario", adVarChar, adParamInput, 10, CodigoUsuario)
     MiCmd.Parameters.Append MiCmd.CreateParameter("@NumModulo", adVarChar, adParamInput, 2, NumModulo)
     MiCmd.Parameters.Append MiCmd.CreateParameter("@Texto", adVarChar, adParamInput, 100, MidStrg(Texto, 1, 100))
     Procesar_Stored_Procedure MiCmd, MiReg
     Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg

End Sub

Public Sub Actualizar_Datos_ATS_SP(Items As String, MBFechaI As String, MBFechaF As String, Numero As Long, ATFisico As Boolean)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
Dim AdoCtasDB As ADODB.Recordset
Dim Tiempo_Espera As Byte
Dim FechaIni As String
Dim FechaFin As String

    FechaIni = BuscarFecha(MBFechaI)
    FechaFin = BuscarFecha(MBFechaF)
    
    Iniciar_Stored_Procedure "Procesando Datos del ATS", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Actualizar_Datos_ATS"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, Items)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaDesde", adVarChar, adParamInput, 10, FechaIni)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaHasta", adVarChar, adParamInput, 10, FechaFin)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Numero", adInteger, adParamInput, 14, Numero)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@ATFisico", adBoolean, adParamInput, 1, ATFisico)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
    MsgBox "...."
End Sub

Public Sub Iniciar_Datos_Default_SP()
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
    
    Iniciar_Stored_Procedure "Procesando Iniciar Datos Default", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Iniciar_Datos_Default"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@RUCEmpresa", adVarChar, adParamInput, 13, RUC)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@CodigoUsuario", adVarChar, adParamInput, 10, CodigoUsuario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaC", adVarChar, adParamInput, 10, BuscarFecha(Fecha_CE))
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NumModulo", adVarChar, adParamInput, 2, NumModulo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Cotizacion", adCurrency, adParamInput, 6, Dolar)
   'Parametros de Salida
    MiCmd.Parameters.Append MiCmd.CreateParameter("@No_ATS", adVarChar, adParamOutput, 2048, No_ATS)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@ListSucursales", adVarChar, adParamOutput, 2048, ListSucursales)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NombreProvincia", adVarChar, adParamOutput, 35, NombreProvincia)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@ConSucursal", adBoolean, adParamOutput, 1, ConSucursal)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@SiUnidadEducativa", adBoolean, adParamOutput, 1, SiUnidadEducativa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@PorcIVA", adSingle, adParamOutput, 2, Porc_IVA)
   'Recibimos datos del resultado del SP
    Procesar_Stored_Procedure MiCmd, MiReg
    No_ATS = MiCmd.Parameters("@No_ATS").value
    ListSucursales = MiCmd.Parameters("@ListSucursales").value
    NombreProvincia = MiCmd.Parameters("@NombreProvincia").value
    ConSucursal = MiCmd.Parameters("@ConSucursal").value
    SiUnidadEducativa = MiCmd.Parameters("@SiUnidadEducativa").value
    Porc_IVA = MiCmd.Parameters("@PorcIVA").value
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
    
'''    MsgBox No_ATS & vbCrLf _
'''         & ListSucursales & vbCrLf _
'''         & ConSucursal & vbCrLf _
'''         & SiUnidadEducativa & vbCrLf _
'''         & NombreProvincia & vbCrLf _
'''         & Porc_IVA
End Sub

Public Sub Eliminar_Duplicados_SP(NombreTabla As String, CamposDuplicados As String, CampoPivote1 As String, CampoPivote2 As String, Optional Item_000 As Boolean)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
Dim AdoCtasDB As ADODB.Recordset
    
    Iniciar_Stored_Procedure "Eliminar duplicados de: " & NombreTabla, MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Eliminar_Duplicados"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NombreTabla", adVarChar, adParamInput, 60, NombreTabla)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@CamposDuplicados", adVarChar, adParamInput, 1024, CamposDuplicados)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@CampoPivote1", adVarChar, adParamInput, 60, CampoPivote1)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@CampoPivote2", adVarChar, adParamInput, 60, CampoPivote2)
    If Item_000 Then
       MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, "000")
    Else
       MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    End If
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Subir_Archivo_Plano_SP(NombreTabla As String, RutaArchivo As String, SeparadorCampo As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

    Iniciar_Stored_Procedure "Subir Archivo de: " & NombreTabla, MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Subir_Archivo_Plano"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NombreTabla", adVarChar, adParamInput, 255, NombreTabla)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@RutaArchivo", adVarChar, adParamInput, 255, RutaArchivo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@SeparadorCampo", adVarChar, adParamInput, 1, SeparadorCampo)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub


Public Sub Eliminar_Periodo_SP(Periodo As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

    Iniciar_Stored_Procedure NumEmpresa & " - Eliminando Periodo " & Periodo, MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Eliminar_Periodo"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Eliminar_Asientos_SP(B_Asiento As Boolean)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

    Iniciar_Stored_Procedure NumEmpresa & " - Eliminar Asientos " & Periodo, MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Eliminar_Asientos"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Usuario", adVarChar, adParamInput, 10, CodigoUsuario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@TransNo", adInteger, adParamInput, 14, Trans_No)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@B_Asiento", adBoolean, adParamInput, 1, B_Asiento)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Leer_Codigo_Inv_SP(BuscarCodigo As String, FechaInventario As String, CodBodega As String, CodMarca As String, CodigoDeInv As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
Dim FechaKardex As String

    FechaKardex = BuscarFecha(FechaInventario)
    CodigoDeInv = Ninguno
    Iniciar_Stored_Procedure "", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Leer_Codigo_Inv"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@BuscarCodigo", adVarChar, adParamInput, 130, BuscarCodigo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaInventario", adVarChar, adParamInput, 10, FechaKardex)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@CodBodega", adVarChar, adParamInput, 18, CodBodega)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@CodMarca", adVarChar, adParamInput, 25, CodMarca)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@CodigoDeInv", adVarChar, adParamOutput, 25, CodigoDeInv)
    Procesar_Stored_Procedure MiCmd, MiReg
   'Recojemos los datos salientes de los campos que retorna valor el store procedure del SQL Server
    CodigoDeInv = MiCmd.Parameters("@CodigodeInv").value
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Leer_Datos_Cliente_SP(BuscarCodigo As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
Dim BuscarCodigo1 As String

   'MsgBox BuscarCodigo
    Iniciar_Stored_Procedure "", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Leer_Datos_Cliente"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@BuscarCodigo", adVarChar, adParamInput, 180, BuscarCodigo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Codigo_Encontrado", adVarChar, adParamOutput, 10, BuscarCodigo1)
    Procesar_Stored_Procedure MiCmd, MiReg
   'Recojemos los datos salientes de los campos que retorna valor el store procedure del SQL Server
    BuscarCodigo = MiCmd.Parameters("@Codigo_Encontrado").value
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Reporte_CxCxP_x_Meses_SP(CtaSubMod As String, MBFechaF As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
Dim FechaFin As String

    FechaFin = BuscarFecha(MBFechaF)
    
    Iniciar_Stored_Procedure "Reporte CxCxP x Meses", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Reporte_CxCxP_x_Meses"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Cta", adVarChar, adParamInput, 18, CtaSubMod)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@CodigoUsuario", adVarChar, adParamInput, 10, CodigoUsuario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@FechaCorte", adVarChar, adParamInput, 10, FechaFin)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
End Sub

Public Sub Listar_Comprobante_SP(C1 As Comprobantes)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

Dim AdoRegistros As ADODB.Recordset
Dim ExisteComp As Boolean
Dim sSQLAux As String

   'Determinamos espacios de memoria para grabar
    RatonReloj
    If Trans_No <= 0 Then Trans_No = 1
    If Ln_No <= 0 Then Ln_No = 1
    If LnSC_No <= 0 Then LnSC_No = 1
    If Ret_No <= 0 Then Ret_No = 1
    
    ExisteComp = False

 'Encabezado del Comprobante
  sSQL = "SELECT C.Fecha, C.Codigo_B, C.Concepto, C.Cotizacion, C.Monto_Total, C.Efectivo, Cl.CI_RUC, Cl.Cliente, Cl.Email, Cl.TD, " _
       & "Cl.Direccion, Cl.Telefono, Cl.Grupo, Cl.RISE, Cl.Especial " _
       & "FROM Comprobantes As C, Clientes As Cl " _
       & "WHERE C.Item = '" & C1.Item & "' " _
       & "AND C.Periodo = '" & Periodo_Contable & "' " _
       & "AND C.TP = '" & C1.TP & "' " _
       & "AND C.Numero = " & C1.Numero & " " _
       & "AND C.Codigo_B = Cl.Codigo "
  Select_AdoDB AdoRegistros, sSQL
  With AdoRegistros
   If .RecordCount > 0 Then
       C1.CodigoB = .fields("Codigo_B")
       C1.Beneficiario = .fields("Cliente")
       C1.Email = .fields("Email")
       C1.Concepto = .fields("Concepto")
       C1.Cotizacion = .fields("Cotizacion")
       C1.Monto_Total = .fields("Monto_Total")
       C1.Efectivo = .fields("Efectivo")
       C1.RUC_CI = .fields("CI_RUC")
       C1.TD = .fields("TD")
       C1.Fecha = .fields("Fecha")
       C1.Direccion = .fields("Direccion")
       C1.Telefono = .fields("Telefono")
       C1.Grupo = .fields("Grupo")
       If .fields("RISE") Then C1.TipoContribuyente = C1.TipoContribuyente & " RISE"
       If .fields("Especial") Then C1.TipoContribuyente = C1.TipoContribuyente & " Contribuyente especial"
      'TipoSRI = consulta_RUC_SRI( C1.RUC_CI)

       If Len(C1.RUC_CI) = 13 Then Tipo_Contribuyente_SP_MySQL C1.RUC_CI, TipoSRI.MicroEmpresa, TipoSRI.AgenteRetencion
       Select Case C1.TD
         Case "C": TipoSRI.Estado = "CEDULA"
         Case "P": TipoSRI.Estado = "PASAPORTE"
         Case "R": TipoSRI.Estado = "RUC ACTIVO"
       End Select
       C1.AgenteRetencion = TipoSRI.AgenteRetencion
       C1.MicroEmpresa = TipoSRI.MicroEmpresa
       C1.Estado = TipoSRI.Estado
       ExisteComp = True
   End If
  End With
  AdoRegistros.Close
  
 'Si existe el comprobante lo presentamos
  If ExisteComp Then
     Iniciar_Stored_Procedure "Listar Comprobante", MiSQL, MiCmd, MiReg
     MiCmd.CommandText = "sp_Listar_Comprobante"
     MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
     MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
     MiCmd.Parameters.Append MiCmd.CreateParameter("@CodigoUsuario", adVarChar, adParamInput, 10, CodigoUsuario)
     MiCmd.Parameters.Append MiCmd.CreateParameter("@TransNo", adInteger, adParamInput, 14, Trans_No)
     MiCmd.Parameters.Append MiCmd.CreateParameter("@TP", adVarChar, adParamInput, 2, C1.TP)
     MiCmd.Parameters.Append MiCmd.CreateParameter("@Numero", adInteger, adParamInput, 14, C1.Numero)
     MiCmd.Parameters.Append MiCmd.CreateParameter("@RetNueva", adBoolean, adParamOutput, 1, C1.RetNueva)
     MiCmd.Parameters.Append MiCmd.CreateParameter("@SerieR", adVarChar, adParamOutput, 6, C1.Serie_R)
     MiCmd.Parameters.Append MiCmd.CreateParameter("@Retencion", adInteger, adParamOutput, 14, C1.Retencion)
     MiCmd.Parameters.Append MiCmd.CreateParameter("@AutorizacionR", adVarChar, adParamOutput, 49, C1.Autorizacion_R)
     MiCmd.Parameters.Append MiCmd.CreateParameter("@CtasModificar", adVarChar, adParamOutput, 5120, C1.Ctas_Modificar)
     MiCmd.Parameters.Append MiCmd.CreateParameter("@CodigoInvModificar", adVarChar, adParamOutput, 5120, C1.CodigoInvModificar)
     MiCmd.Parameters.Append MiCmd.CreateParameter("@LnNo", adSmallInt, adParamOutput, 2, Ln_No)
     MiCmd.Parameters.Append MiCmd.CreateParameter("@LnSCNo", adSmallInt, adParamOutput, 2, LnSC_No)
     Procesar_Stored_Procedure MiCmd, MiReg
    'Recojemos los datos salientes de los campos que retorna valor el store procedure del SQL Server
     C1.RetNueva = MiCmd.Parameters("@RetNueva").value
     C1.Serie_R = MiCmd.Parameters("@SerieR").value
     C1.Retencion = MiCmd.Parameters("@Retencion").value
     C1.Autorizacion_R = MiCmd.Parameters("@AutorizacionR").value
     C1.Ctas_Modificar = MiCmd.Parameters("@CtasModificar").value
     C1.CodigoInvModificar = MiCmd.Parameters("@CodigoInvModificar").value
     Ln_No = MiCmd.Parameters("@LnNo").value
     LnSC_No = MiCmd.Parameters("@LnSCNo").value
     Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
  End If
  RatonNormal
'''  MsgBox C1.RetNueva & vbCrLf _
'''       & C1.Serie_R & vbCrLf _
'''       & C1.Retencion & vbCrLf _
'''       & C1.Autorizacion_R & vbCrLf _
'''       & C1.Ctas_Modificar & vbCrLf _
'''       & C1.CodigoInvModificar & vbCrLf _
'''       & Ln_No & vbCrLf _
'''       & LnSC_No
End Sub

Public Sub Subir_Archivo_CSV_SP(PathCSV As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

Dim FileCSV As String
Dim PathCSVT As String
Dim LineFile As String
Dim TipoFile As String
Dim NumFile As Long
    RatonReloj
    TipoFile = ""
    If Len(PathCSV) > 1 Then
       NumFile = FreeFile
       Open PathCSV For Input As #NumFile
         If Not EOF(NumFile) Then Line Input #NumFile, LineFile
         If InStr(LineFile, ";emision") > 0 Then TipoFile = "05"
         If InStr(LineFile, ";CI_RUC_Codigo") > 0 Then TipoFile = "15"
         If InStr(LineFile, ";COD_MES") > 0 Then TipoFile = "27"
         If InStr(LineFile, ";CI_RUC_P_SUBMOD") > 0 Then TipoFile = "99"
       Close #NumFile
       If TipoFile <> "" Then
          FileCSV = Right$(PathCSV, Len(PathCSV) - InStrRev(PathCSV, "\"))
          PathCSVT = MidStrg(PathCSV, 1, Len(PathCSV) - Len(FileCSV))
          Iniciar_Stored_Procedure "sp Subir Archivo CSV", MiSQL, MiCmd, MiReg
          MiCmd.CommandText = "sp_Subir_Archivo_CSV"
          MiCmd.Parameters.Append MiCmd.CreateParameter("@strIPServidor", adVarChar, adParamInput, 100, strIPServidor)
          MiCmd.Parameters.Append MiCmd.CreateParameter("@PathFileCSV", adVarChar, adParamInput, 255, PathCSVT)
          MiCmd.Parameters.Append MiCmd.CreateParameter("@FileCSV", adVarChar, adParamInput, 100, FileCSV)
          MiCmd.Parameters.Append MiCmd.CreateParameter("@Usuario", adVarChar, adParamInput, 10, CodigoUsuario)
          MiCmd.Parameters.Append MiCmd.CreateParameter("@TipoFile", adVarChar, adParamInput, 2, TipoFile)
          Procesar_Stored_Procedure MiCmd, MiReg
          Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
       End If
    End If
    RatonNormal
End Sub

Public Sub Subir_Archivo_TXT_SP(PathTXT As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset

Dim FileTXT As String
Dim PathTXTT As String
Dim LineFile As String
Dim TipoFile As String
Dim NumFile As Long
    RatonReloj
    TipoFile = ""
    If Len(PathTXT) > 1 Then
       NumFile = FreeFile
       Open PathTXT For Input As #NumFile
         If Not EOF(NumFile) Then Line Input #NumFile, LineFile
         If InStr(LineFile, "CLAVE_ACCESO") > 0 Then TipoFile = "SRI"
       Close #NumFile
       
       If TipoFile <> "" Then
          FileTXT = Right$(PathTXT, Len(PathTXT) - InStrRev(PathTXT, "\"))
          PathTXTT = MidStrg(PathTXT, 1, Len(PathTXT) - Len(FileTXT))
          Iniciar_Stored_Procedure "sp Subir Archivo TXT", MiSQL, MiCmd, MiReg
          MiCmd.CommandText = "sp_Subir_Archivo_TXT"
          MiCmd.Parameters.Append MiCmd.CreateParameter("@strIPServidor", adVarChar, adParamInput, 100, strIPServidor)
          MiCmd.Parameters.Append MiCmd.CreateParameter("@PathFileTXT", adVarChar, adParamInput, 255, PathTXTT)
          MiCmd.Parameters.Append MiCmd.CreateParameter("@FileTXT", adVarChar, adParamInput, 100, FileTXT)
          MiCmd.Parameters.Append MiCmd.CreateParameter("@Usuario", adVarChar, adParamInput, 10, CodigoUsuario)
          MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
          MiCmd.Parameters.Append MiCmd.CreateParameter("@TipoFile", adVarChar, adParamInput, 5, TipoFile)
          Procesar_Stored_Procedure MiCmd, MiReg
          Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
       End If
    End If
    RatonNormal
End Sub

Public Sub Importar_Contabilidad_SP(vTP As String)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
    RatonReloj
    Iniciar_Stored_Procedure "sp Importar Contabilidad", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Importar_Contabilidad"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Usuario", adVarChar, adParamInput, 10, CodigoUsuario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NumModulo", adVarChar, adParamInput, 2, NumModulo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@TP", adVarChar, adParamInput, 2, vTP)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
    RatonNormal
End Sub

Public Sub Importar_Contabilidad_SubModulos_SP()
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
    RatonReloj
    Iniciar_Stored_Procedure "sp Importar Contabilidad con SubModulo", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Importar_Contabilidad_SubModulos"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Usuario", adVarChar, adParamInput, 10, CodigoUsuario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NumModulo", adVarChar, adParamInput, 2, NumModulo)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
    RatonNormal
End Sub

Public Sub Importar_Abonos_Facturas_SP()
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
    RatonReloj
    Iniciar_Stored_Procedure "sp Importar Abonos Facturas", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Importar_Abonos_Facturas"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Usuario", adVarChar, adParamInput, 10, CodigoUsuario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NumModulo", adVarChar, adParamInput, 2, NumModulo)
    Procesar_Stored_Procedure MiCmd, MiReg
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
    RatonNormal
End Sub

Public Sub Importar_Compras_Diarias_SP(TipoComp As String, Numero As Long)
Dim MiSQL As ADODB.Connection
Dim MiCmd As ADODB.Command
Dim MiReg As ADODB.Recordset
Dim NumComp As Long

    RatonReloj
    Select Case TipoComp
      Case "CE": NumComp = ReadSetDataNum("Egresos", True, False)
      Case "CI": NumComp = ReadSetDataNum("Ingresos", True, False)
      Case Else: NumComp = ReadSetDataNum("Diario", True, False)
                 TipoComp = "CD"
    End Select
    Iniciar_Stored_Procedure "sp Importar Compras Diarias", MiSQL, MiCmd, MiReg
    MiCmd.CommandText = "sp_Importar_Compras_Diarias"
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Item", adVarChar, adParamInput, 3, NumEmpresa)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Periodo", adVarChar, adParamInput, 10, Periodo_Contable)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Usuario", adVarChar, adParamInput, 10, CodigoUsuario)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@NumModulo", adVarChar, adParamInput, 2, NumModulo)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@TP", adVarChar, adParamInput, 2, TipoComp)
    MiCmd.Parameters.Append MiCmd.CreateParameter("@Numero", adInteger, adParamOutput, 14, NumComp)
    Procesar_Stored_Procedure MiCmd, MiReg
   'Recojemos los datos salientes de los campos que retorna valor el store procedure del SQL Server
    Numero = MiCmd.Parameters("@Numero").value
    Finalizar_Stored_Procedure MiSQL, MiCmd, MiReg
    RatonNormal
End Sub

