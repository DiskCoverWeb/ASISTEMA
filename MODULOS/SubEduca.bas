Attribute VB_Name = "SubEducativo"
Option Explicit

Public Sub Encabezado_Institucion(X0 As Single, x1 As Single)
Dim Y0 As Single
Dim y1 As Single
   PosLinea = 0.5
   PorteLetra = Printer.FontSize
   LetraAnterior = Printer.FontName
   Printer.FontName = TipoTimes
   Printer.FontItalic = False
   Printer.FontSize = 8
   If X0 < 0.5 Then X0 = 0.5
   Y0 = 0.5
   y1 = Y0 + 1.1
   If x1 > LimiteAncho Then x1 = LimiteAncho - 0.1
   PrinterPaint LogoTipo, X0, Y0, 3, 1.5
   If Pagina > 0 Then
      Printer.FontSize = 9
      PrinterTexto x1 - 1, Y0, "Pág. No. " & CStr(Pagina)
      Pagina = Pagina + 1
   End If
   Printer.FontBold = True
   Printer.ForeColor = Gris
   Printer.FontSize = 12: Printer.FontItalic = False
   PrinterTexto CentrarTextoEncab(UCaseStrg(Institucion1), X0, x1) + 0.03, Y0, UCaseStrg(Institucion1)
   Y0 = Y0 + 0.03
   Printer.ForeColor = Negro
   PrinterTexto CentrarTextoEncab(UCaseStrg(Institucion1), X0, x1), Y0, UCaseStrg(Institucion1)
   Printer.FontSize = 18
   Y0 = Y0 + 0.45
   PrinterTexto CentrarTextoEncab(UCaseStrg(Institucion2), X0, x1) + 0.025, Y0, UCaseStrg(Institucion2)
   Y0 = Y0 + 0.025
   Printer.ForeColor = Negro
   PrinterTexto CentrarTextoEncab(UCaseStrg(Institucion2), X0, x1), Y0, UCaseStrg(Institucion2)
   Y0 = Y0 + 1

   If MensajeEncabData <> "" Then
      Printer.FontSize = 12
      PrinterTexto CentrarTextoEncab(MensajeEncabData, X0, x1), Y0, MensajeEncabData
      Y0 = Y0 + 0.8
   End If
   
   If SQLMsg1 <> "" Then
      Printer.FontSize = 12
      PrinterTexto CentrarTextoEncab(SQLMsg1, X0, x1), Y0, SQLMsg1
      Y0 = Y0 + 0.7
   End If
   Printer.FontBold = False
   If SQLMsg2 <> "" Then
      Printer.FontSize = 12
      PrinterTexto CentrarTextoEncab(SQLMsg2, X0, x1), Y0, SQLMsg2
      Y0 = Y0 + 0.7
   End If
   If SQLMsg3 <> "" Then
      Printer.FontSize = 14
      PrinterTexto CentrarTextoEncab(SQLMsg3, X0, x1), Y0, SQLMsg3
      Y0 = Y0 + 0.8
   End If
   If SQLMsg4 <> "" Then
      Printer.FontSize = 10
      PrinterTexto X0, Y0, SQLMsg4
      Y0 = Y0 + 0.6
   End If
   PosLinea = Y0
   Printer.FontSize = PorteLetra
   Printer.FontName = LetraAnterior
End Sub

Public Function Visualizar_Notas_Periodo(LstPeriodo As ListBox) As String
     SQLTAI = Ninguno
     SQLAIC = Ninguno
     SQLAGC = Ninguno
     SQLL = Ninguno
     SQLExaP = Ninguno
     SQLNotas = Ninguno
     SQLBim1 = Ninguno
     SQLBim2 = Ninguno
     SQLBim3 = Ninguno
     SQLExamen = Ninguno
     SQLPromQ = Ninguno
     Visualizar_Notas_Periodo = LstPeriodo.Text
     If OpcPeriodo("PF", LstPeriodo) Then
        Visualizar_Notas_Periodo = "Periodo Final de Quimestres"
        OpcionNotas = 5
     End If
     If OpcPeriodo("PQBim1", LstPeriodo) Then
        SQLTAI = "PQTAI1"
        SQLAIC = "PQAIC1"
        SQLAGC = "PQAGC1"
        SQLL = "PQL1"
        SQLExaP = "PQExaP1"
        SQLNotas = "PQBim1"
         
        SQLBim1 = "PQBim1"
        SQLBim2 = "PQBim2"
        SQLBim3 = "PQBim3"
        SQLProm = "PQBim1"
        SQLExamen = "ExamenPQ"
        SQLPromQ = "PromPQ"
        
        SQLConductaQ = "ConductaPQ1"
        SQLDias = "PQDias1"
        SQLFJ = "PQBFJ1"
        SQLFI = "PQBFI1"
        SQLAtrasos = "PQBA1"
        OpcionNotas = 1
     End If
     If OpcPeriodo("PQBim2", LstPeriodo) Then
        SQLTAI = "PQTAI2"
        SQLAIC = "PQAIC2"
        SQLAGC = "PQAGC2"
        SQLL = "PQL2"
        SQLExaP = "PQExaP2"
        SQLNotas = "PQBim2"
         
        SQLBim1 = "PQBim1"
        SQLBim2 = "PQBim2"
        SQLBim3 = "PQBim3"
        SQLProm = "PQBim2"
        SQLExamen = "ExamenPQ"
        SQLPromQ = "PromPQ"
        
        SQLConductaQ = "ConductaPQ2"
        SQLDias = "PQDias2"
        SQLFJ = "PQBFJ2"
        SQLFI = "PQBFI2"
        SQLAtrasos = "PQBA2"
        OpcionNotas = 2
     End If
     If OpcPeriodo("PQBim3", LstPeriodo) Then
        SQLTAI = "PQTAI3"
        SQLAIC = "PQAIC3"
        SQLAGC = "PQAGC3"
        SQLL = "PQL3"
        SQLExaP = "PQExaP3"
        SQLNotas = "PQBim3"
        
        SQLBim1 = "PQBim1"
        SQLBim2 = "PQBim2"
        SQLBim3 = "PQBim3"
        SQLProm = "PQBim3"
        SQLExamen = "ExamenPQ"
        SQLPromQ = "PromPQ"
        
        SQLConductaQ = "ConductaPQ3"
        SQLDias = "PQDias3"
        SQLFJ = "PQBFJ3"
        SQLFI = "PQBFI3"
        SQLAtrasos = "PQBA3"
        OpcionNotas = 3
     End If
     If OpcPeriodo("PQ", LstPeriodo) Then
        SQLTAI = "PQTAI3"
        SQLAIC = "PQAIC3"
        SQLAGC = "PQAGC3"
        SQLL = "PQL3"
        SQLExaP = "PQExaP3"
        SQLNotas = "PQBim3"
        
        SQLBim1 = "PQBim1"
        SQLBim2 = "PQBim2"
        SQLBim3 = "PQBim3"
        SQLProm = "PQBim3"
        SQLExamen = "ExamenPQ"
        SQLPromQ = "PromPQ"
        
        SQLConductaQ = "ConductaPQ1"
        SQLDias = "PQDias1"
        SQLFJ = "PQBFJ1"
        SQLFI = "PQBFI1"
        SQLAtrasos = "PQBA1"
        OpcionNotas = 4
     End If
     If OpcPeriodo("SQBim1", LstPeriodo) Then
        SQLTAI = "SQTAI1"
        SQLAIC = "SQAIC1"
        SQLAGC = "SQAGC1"
        SQLL = "SQL1"
        SQLExaP = "SQExaP1"
        SQLNotas = "SQBim1"
         
        SQLBim1 = "SQBim1"
        SQLBim2 = "SQBim2"
        SQLBim3 = "SQBim3"
        SQLProm = "SQBim1"
        SQLExamen = "ExamenSQ"
        SQLPromQ = "PromSQ"
        
        SQLConductaQ = "ConductaSQ1"
        SQLDias = "SQDias1"
        SQLFJ = "SQBFJ1"
        SQLFI = "SQBFI1"
        SQLAtrasos = "SQBA1"
        OpcionNotas = 1
     End If
     If OpcPeriodo("SQBim2", LstPeriodo) Then
        SQLTAI = "SQTAI2"
        SQLAIC = "SQAIC2"
        SQLAGC = "SQAGC2"
        SQLL = "SQL2"
        SQLExaP = "SQExaP2"
        SQLNotas = "SQBim2"
         
        SQLBim1 = "SQBim1"
        SQLBim2 = "SQBim2"
        SQLBim3 = "SQBim3"
        SQLProm = "SQBim2"
        SQLExamen = "ExamenSQ"
        SQLPromQ = "PromSQ"
        
        SQLConductaQ = "ConductaSQ2"
        SQLDias = "SQDias2"
        SQLFJ = "SQBFJ2"
        SQLFI = "SQBFI2"
        SQLAtrasos = "SQBA2"
        OpcionNotas = 2
     End If
     If OpcPeriodo("SQBim3", LstPeriodo) Then
        SQLTAI = "SQTAI3"
        SQLAIC = "SQAIC3"
        SQLAGC = "SQAGC3"
        SQLL = "SQL3"
        SQLExaP = "SQExaP3"
        SQLNotas = "SQBim3"
         
        SQLBim1 = "SQBim1"
        SQLBim2 = "SQBim2"
        SQLBim3 = "SQBim3"
        SQLProm = "SQBim3"
        SQLExamen = "ExamenSQ"
        SQLPromQ = "PromSQ"
        
        SQLConductaQ = "ConductaSQ3"
        SQLDias = "SQDias3"
        SQLFJ = "SQBFJ3"
        SQLFI = "SQBFI3"
        SQLAtrasos = "SQBA3"
        OpcionNotas = 3
     End If
     If OpcPeriodo("SQ", LstPeriodo) Then
        SQLTAI = "SQTAI3"
        SQLAIC = "SQAIC3"
        SQLAGC = "SQAGC3"
        SQLL = "SQL3"
        SQLExaP = "SQExaP3"
        SQLNotas = "SQBim3"
         
        SQLBim1 = "SQBim1"
        SQLBim2 = "SQBim2"
        SQLBim3 = "SQBim3"
        SQLProm = "SQBim3"
        SQLExamen = "ExamenSQ"
        SQLPromQ = "PromSQ"
        SQLConductaQ = "ConductaSQ1"
        
        SQLDias = "SQDias1"
        SQLFJ = "SQBFJ1"
        SQLFI = "SQBFI1"
        SQLAtrasos = "SQBA1"
        OpcionNotas = 4
     End If
     If OpcPeriodo("TQBim1", LstPeriodo) Then
        SQLBim1 = "TQBim1"
        SQLBim2 = "TQBim2"
        SQLBim3 = "TQBim3"
        SQLExamen = "ExamenTQ"
        SQLPromQ = "PromTQ"
        OpcionNotas = 1
     End If
     If OpcPeriodo("TQBim2", LstPeriodo) Then
        SQLBim1 = "TQBim1"
        SQLBim2 = "TQBim2"
        SQLBim3 = "TQBim3"
        SQLExamen = "ExamenTQ"
        SQLPromQ = "PromTQ"
        OpcionNotas = 2
     End If
     If OpcPeriodo("TQBim3", LstPeriodo) Then
        SQLBim1 = "TQBim1"
        SQLBim2 = "TQBim2"
        SQLBim3 = "TQBim3"
        SQLExamen = "ExamenTQ"
        SQLPromQ = "PromTQ"
        OpcionNotas = 3
     End If
     If OpcPeriodo("TQ", LstPeriodo) Then
        SQLBim1 = "TQBim1"
        SQLBim2 = "TQBim2"
        SQLBim3 = "TQBim3"
        SQLExamen = "ExamenTQ"
        SQLPromQ = "PromTQ"
        OpcionNotas = 4
     End If
End Function

Public Sub Leer_Notas_Parciales(CodMat As String, Curso As String, LstParcial As ListBox)
Dim Strgs As String
Dim Cod_Prof As String
Dim CodMatP As String
Dim AdoRegs As ADODB.Recordset

    CadenaParcial = Visualizar_Notas_Periodo(LstParcial)
    Cod_Prof = Ninguno
    Select Case CodMat
      Case "998", "999"
           sSQL = "DELETE * " _
                & "FROM Asiento_AS "
      Case Else
           sSQL = "DELETE * " _
                & "FROM Asiento_N "
    End Select
    sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodMat = '" & CodMat & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Ejecutar_SQL_SP sSQL
    Contador = 0
    'C.Sexo DESC
    Set AdoRegs = New ADODB.Recordset
    AdoRegs.CursorType = adOpenStatic
    AdoRegs.CursorLocation = adUseClient
    CodMatP = Ninguno
    Strgs = "SELECT * " _
          & "FROM Catalogo_Estudiantil " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND MidStrg(CodigoE,1," & Len(Curso) & ") = '" & Curso & "' " _
          & "AND CodMat = '" & CodMat & "' "
    Strgs = CompilarSQL(Strgs)
    AdoRegs.open Strgs, AdoStrCnn, , , adCmdText
    With AdoRegs
     If .RecordCount > 0 Then
         Cod_Prof = AdoRegs.fields("Profesor")
         CodMatP = AdoRegs.fields("CodMatP")
     End If
    End With
    AdoRegs.Close
    Strgs = "SELECT C.Cliente,C.Grupo,C.Sexo,CM.* " _
          & "FROM Clientes As C,Clientes_Matriculas As CM " _
          & "WHERE CM.Item = '" & NumEmpresa & "' " _
          & "AND CM.Periodo = '" & Periodo_Contable & "' " _
          & "AND CM.Grupo_No = '" & Curso & "' " _
          & "AND CM.Codigo = C.Codigo " _
          & "ORDER BY C.Cliente "
    Strgs = CompilarSQL(Strgs)
    AdoRegs.open Strgs, AdoStrCnn, , , adCmdText
    With AdoRegs
     If .RecordCount > 0 Then
         'MsgBox .RecordCount
         Do While Not .EOF
            Contador = Contador + 1
            Select Case CodMat
              Case "998", "999"
                   SetAdoAddNew "Asiento_AS"
              Case Else
                   SetAdoAddNew "Asiento_N"
            End Select
            SetAdoFields "Id_No", CByte(Contador)
            SetAdoFields "CodMat", CodMat
            SetAdoFields "Codigo", .fields("Codigo")
            SetAdoFields "Alumno", .fields("Cliente")
            SetAdoFields "Profesor", Cod_Prof
            SetAdoFields "Item", NumEmpresa
            SetAdoFields "CodigoU", CodigoUsuario
            SetAdoUpdate
            If Contador > 254 Then Contador = 254
           .MoveNext
         Loop
     End If
    End With
    AdoRegs.Close
    If FormatoLibreta = "QUIMESTRE" Then
       CadenaParcial = Visualizar_Notas_Periodo(LstParcial)
        Select Case CodMat
          Case "998", "999"
               If SQL_Server Then
                  sSQL = "UPDATE Asiento_AS " _
                       & "SET " & SQLConductaQ & " = TN." & SQLConductaQ & ", " _
                       & SQLDias & " = TN." & SQLDias & ", " _
                       & SQLFJ & " = TN." & SQLFJ & ", " _
                       & SQLFI & " = TN." & SQLFI & ", " _
                       & SQLAtrasos & " = TN." & SQLAtrasos & " " _
                       & "FROM Asiento_AS As AN,Trans_Asistencia As TN "
                Else
               End If
          Case Else
               If SQL_Server Then
                  sSQL = "UPDATE Asiento_N "
                  If OpcionNotas = 5 Then
                     sSQL = sSQL & "SET Nota_Grado = TN.Nota_Grado, " _
                          & "Supletorio = TN.Supletorio, " _
                          & "Remedial = TN.Remedial "
                  Else
                     sSQL = sSQL & "SET TAI = TN." & SQLTAI & ", " _
                          & "AIC = TN." & SQLAIC & ", " _
                          & "AGC = TN." & SQLAGC & ", " _
                          & "LECCIONES = TN." & SQLL & ", " _
                          & "EXAMEN = TN." & SQLExaP & ", " _
                          & "Examen_Q = TN." & SQLExamen & ", " _
                          & "Nota_Grado = TN.Nota_Grado, " _
                          & "Supletorio = TN.Supletorio "
                  End If
                  If CodMatP = Ninguno Then
                     sSQL = sSQL & "FROM Asiento_N As AN,Trans_Notas As TN "
                  Else
                     sSQL = sSQL & "FROM Asiento_N As AN,Trans_Notas_Auxiliares As TN "
                  End If
               Else
               End If
               
        End Select
        sSQL = sSQL & "WHERE TN.Item = '" & NumEmpresa & "' " _
             & "AND TN.Periodo = '" & Periodo_Contable & "' " _
             & "AND TN.CodE = '" & Curso & "' " _
             & "AND TN.CodMat = '" & CodMat & "' " _
             & "AND TN.CodMat = AN.CodMat " _
             & "AND TN.Item = AN.Item " _
             & "AND TN.Codigo = AN.Codigo "
    Else
        Select Case CodMat
          Case "998", "999"
               If SQL_Server Then
                  sSQL = "UPDATE Asiento_AS " _
                       & "SET PQBFJ1 = TN.PQBFJ1," _
                       & "PQBFI1 = TN.PQBFI1," _
                       & "PQBA1 = TN.PQBA1," _
                       & "ConductaPQ1 = TN.ConductaPQ1," _
                       & "PQBFJ2 = TN.PQBFJ2," _
                       & "PQBFI2 = TN.PQBFI2," _
                       & "PQBA2 = TN.PQBA2," _
                       & "ConductaPQ2 = TN.ConductaPQ2,"
                 sSQL = sSQL _
                       & "SQBFJ1 = TN.SQBFJ1," _
                       & "SQBFI1 = TN.SQBFI1," _
                       & "SQBA1 = TN.SQBA1," _
                       & "ConductaSQ1 = TN.ConductaSQ1," _
                       & "SQBFJ2 = TN.SQBFJ2," _
                       & "SQBFI2 = TN.SQBFI2," _
                       & "SQBA2 = TN.SQBA2," _
                       & "ConductaSQ2 = TN.ConductaSQ2,"
                 sSQL = sSQL _
                       & "TQBFJ1 = TN.TQBFJ1," _
                       & "TQBFI1 = TN.TQBFI1," _
                       & "TQBA1 = TN.TQBA1," _
                       & "ConductaTQ1 = TN.ConductaTQ1," _
                       & "TQBFJ2 = TN.TQBFJ2," _
                       & "TQBFI2 = TN.TQBFI2," _
                       & "TQBA2 = TN.TQBA2," _
                       & "ConductaTQ2 = TN.ConductaTQ2 "
                 sSQL = sSQL & "FROM Asiento_AS As AN,Trans_Asistencia As TN "
               Else
                  sSQL = "UPDATE Asiento_AS As AN,Trans_Asistencia As TN " _
                       & "SET AN.PQBFJ1 = TN.PQBFJ1," _
                       & "AN.PQBFI1 = TN.PQBFI1," _
                       & "AN.PQBA1 = TN.PQBA1," _
                       & "AN.ConductaPQ1 = TN.ConductaPQ1," _
                       & "AN.PQBFJ2 = TN.PQBFJ2," _
                       & "AN.PQBFI2 = TN.PQBFI2," _
                       & "AN.PQBA2 = TN.PQBA2," _
                       & "AN.ConductaPQ2 = TN.ConductaPQ2,"
                  sSQL = sSQL _
                       & "AN.SQBFJ1 = TN.SQBFJ1," _
                       & "AN.SQBFI1 = TN.SQBFI1," _
                       & "AN.SQBA1 = TN.SQBA1," _
                       & "AN.ConductaSQ1 = TN.ConductaSQ1," _
                       & "AN.SQBFJ2 = TN.SQBFJ2," _
                       & "AN.SQBFI2 = TN.SQBFI2," _
                       & "AN.SQBA2 = TN.SQBA2," _
                       & "AN.ConductaSQ2 = TN.ConductaSQ2,"
                  sSQL = sSQL _
                       & "AN.TQBFJ1 = TN.TQBFJ1," _
                       & "AN.TQBFI1 = TN.TQBFI1," _
                       & "AN.TQBA1 = TN.TQBA1," _
                       & "AN.ConductaTQ1 = TN.ConductaTQ1," _
                       & "AN.TQBFJ2 = TN.TQBFJ2," _
                       & "AN.TQBFI2 = TN.TQBFI2," _
                       & "AN.TQBA2 = TN.TQBA2," _
                       & "AN.ConductaTQ2 = TN.ConductaTQ2 "
               End If
               sSQL = sSQL & "WHERE TN.Item = '" & NumEmpresa & "' " _
                     & "AND TN.Periodo = '" & Periodo_Contable & "' " _
                     & "AND TN.CodE = '" & Curso & "' " _
                     & "AND TN.CodMat = '" & CodMat & "' " _
                     & "AND TN.CodMat = AN.CodMat " _
                     & "AND TN.Item = AN.Item " _
                     & "AND TN.Codigo = AN.Codigo "
          Case Else
               If SQL_Server Then
                  sSQL = "UPDATE Asiento_N " _
                       & "SET PQBim1 = TN.PQBim1," _
                       & "PQBim2 = TN.PQBim2," _
                       & "SQBim1 = TN.SQBim1," _
                       & "SQBim2 =  TN.SQBim2," _
                       & "TQBim1 = TN.TQBim1," _
                       & "TQBim2 =  TN.TQBim2," _
                       & "ExamenPQ = TN.ExamenPQ," _
                       & "ExamenSQ = TN.ExamenSQ," _
                       & "ExamenTQ = TN.ExamenTQ," _
                       & "ConductaPQ1 = TN.ConductaPQ1," _
                       & "ConductaPQ2 = TN.ConductaPQ2," _
                       & "ConductaSQ1 = TN.ConductaSQ1," _
                       & "ConductaSQ2 = TN.ConductaSQ2," _
                       & "ConductaTQ1 = TN.ConductaTQ1," _
                       & "ConductaTQ2 = TN.ConductaTQ2," _
                       & "Nota_Grado = TN.Nota_Grado," _
                       & "Supletorio  = TN.Supletorio "
                  If CodMatP = Ninguno Then
                     sSQL = sSQL & "FROM Asiento_N As AN,Trans_Notas As TN "
                  Else
                     sSQL = sSQL & "FROM Asiento_N As AN,Trans_Notas_Auxiliares As TN "
                  End If
               Else
                  If CodMatP = Ninguno Then
                     sSQL = "UPDATE Asiento_N As AN,Trans_Notas As TN "
                  Else
                     sSQL = "UPDATE Asiento_N As AN,Trans_Notas_Auxiliares As TN "
                  End If
                  sSQL = sSQL _
                       & "SET AN.PQBim1 = TN.PQBim1," _
                       & "AN.PQBim2 = TN.PQBim2," _
                       & "AN.SQBim1 = TN.SQBim1," _
                       & "AN.SQBim2 =  TN.SQBim2," _
                       & "AN.TQBim1 = TN.TQBim1," _
                       & "AN.TQBim2 =  TN.TQBim2," _
                       & "AN.ExamenPQ = TN.ExamenPQ," _
                       & "AN.ExamenSQ = TN.ExamenSQ," _
                       & "AN.ExamenTQ = TN.ExamenTQ," _
                       & "AN.ConductaPQ1 = TN.ConductaPQ1," _
                       & "AN.ConductaPQ2 = TN.ConductaPQ2," _
                       & "AN.ConductaSQ1 = TN.ConductaSQ1," _
                       & "AN.ConductaSQ2 = TN.ConductaSQ2," _
                       & "AN.ConductaTQ1 = TN.ConductaTQ1," _
                       & "AN.ConductaTQ2 = TN.ConductaTQ2," _
                       & "AN.Nota_Grado = TN.Nota_Grado," _
                       & "AN.Supletorio  = TN.Supletorio "
               End If
               sSQL = sSQL & "WHERE TN.Item = '" & NumEmpresa & "' " _
                     & "AND TN.Periodo = '" & Periodo_Contable & "' " _
                     & "AND TN.CodE = '" & Curso & "' " _
                     & "AND TN.CodMat = '" & CodMat & "' " _
                     & "AND TN.CodMat = AN.CodMat " _
                     & "AND TN.Item = AN.Item " _
                     & "AND TN.Codigo = AN.Codigo "
        End Select
    End If
    Ejecutar_SQL_SP sSQL
End Sub


''Public Sub Leer_Materias_Por_Curso(Curso As String)
''Dim AdoRegs As ADODB.Recordset
''Dim SQLQuery As String
''  SQLQuery = "SELECT CE.CodMat,CM.Materia,CE.CodMatP,CM.C,CM.I,CM.P,CM.SDiv " _
''           & "FROM Catalogo_Estudiantil As CE,Catalogo_Materias As CM " _
''           & "WHERE CE.Item = '" & NumEmpresa & "' " _
''           & "AND CE.Periodo = '" & Periodo_Contable & "' " _
''           & "AND MidStrg(CE.CodigoE,1," & Len(Curso) & ") = '" & Curso & "' " _
''           & "AND CE.CodMatP = '.' " _
''           & "AND CE.TC = 'M' " _
''           & "AND CE.Item = CM.Item " _
''           & "AND CE.Periodo = CM.Periodo " _
''           & "AND CE.CodMat = CM.CodMat " _
''           & "ORDER BY CE.CodigoE "
''  SQLQuery = CompilarSQL(SQLQuery)
''  Set AdoRegs = New ADODB.Recordset
''  AdoRegs.CursorType = adOpenStatic
''  AdoRegs.CursorLocation = adUseClient
''  AdoRegs.Open SQLQuery, AdoStrCnn, , , adCmdText
''  With AdoRegs
''   If .RecordCount > 0 Then
''       ReDim VectMateria(.RecordCount + 1) As TipoMaterias
''       ContNotas = 0
''       Do While Not .EOF
''          VectMateria(ContNotas).CodigoMat = .Fields("CodMat")
''          VectMateria(ContNotas).Materias = .Fields("Materia")
''          ContNotas = ContNotas + 1
''         .MoveNext
''       Loop
''       VectMateria(ContNotas).Materias = "PROMEDIO"
''       ContNotas = ContNotas + 1
''   End If
''  End With
''  AdoRegs.Close
''End Sub

Public Sub Listar_Notas_Alunmos(AdoAut As Adodc, _
                                CodMat As String, _
                                Curso As String, _
                                LstParcial As ListBox)
                                
  CadenaParcial = Visualizar_Notas_Periodo(LstParcial)
  With AdoAut.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      'MsgBox Tipo_Nota & vbCrLf & Todo_los_Campos
       If FormatoLibreta = "QUIMESTRE" Then
          sSQL = "SELECT Id_No, Alumno, "
          Select Case CodMat
            Case "998", "999"
                 sSQL = sSQL _
                      & SQLConductaQ & ", " & SQLDias & ", " & SQLFJ & ", " & SQLFI & ", " & SQLAtrasos & ", " _
                      & "CodMat, Profesor, Item, Codigo, CodigoU " _
                      & "FROM Asiento_AS "
            Case Else
                 If OpcionNotas = 4 Then
                    sSQL = sSQL & "Examen_Q, "
                 ElseIf OpcionNotas = 5 Then
                    sSQL = sSQL & "Nota_Grado, Supletorio, Remedial, "
                 Else
                    sSQL = sSQL & "TAI, AIC, AGC, LECCIONES, Examen,FALTAS_J,FALTAS_I,ATRASOS, "
                 End If
                 sSQL = sSQL _
                      & "CodMat, Profesor, Item, Codigo, CodigoU " _
                      & "FROM Asiento_N "
          End Select
       Else
          sSQL = "SELECT Id_No,Alumno,"
          Select Case CodMat
            Case "998", "999"
                 Select Case OpcionNotas
                   Case 1: sSQL = sSQL & "ConductaPQ1,PQBFJ1,PQBFI1,PQBA1,"
                   Case 2: sSQL = sSQL & "ConductaPQ2,PQBFJ2,PQBFI2,PQBA2,"
                   Case 3: sSQL = sSQL & "ConductaSQ1,SQBFJ1,SQBFI1,SQBA1,"
                   Case 4: sSQL = sSQL & "ConductaSQ2,SQBFJ2,SQBFI2,SQBA2,"
                   Case 5: sSQL = sSQL & "ConductaTQ1,TQBFJ1,TQBFI1,TQBA1,"
                   Case 6: sSQL = sSQL & "ConductaTQ2,TQBFJ2,TQBFI2,TQBA2,"
                 End Select
                 sSQL = sSQL & "CodMat,Profesor,Item,CodigoU " _
                      & "FROM Asiento_AS "
            Case Else
                 Select Case OpcionNotas
                   Case 1: sSQL = sSQL & "PQBim1,ConductaPQ1,"
                   Case 2: sSQL = sSQL & "PQBim1,PQBim2,ExamenPQ,ConductaPQ1,ConductaPQ2,"
                   Case 3: sSQL = sSQL & "SQBim1,ConductaSQ1,"
                   Case 4: sSQL = sSQL & "SQBim1,SQBim2,ExamenSQ,ConductaSQ1,ConductaSQ2,"
                   Case 5: sSQL = sSQL & "TQBim1,ConductaTQ1,"
                   Case 6: sSQL = sSQL & "TQBim1,TQBim2,ExamenTQ,ConductaTQ1,ConductaTQ2,"
                   Case 7: sSQL = sSQL & "ExamenPQ,ExamenSQ,Supletorio,"
                   Case Else
                        sSQL = sSQL _
                             & "PQBim1,PQBim2,ExamenPQ," _
                             & "SQBim1,SQBim2,ExamenSQ," _
                             & "TQBim1,TQBim2,ExamenTQ," _
                             & "ConductaPQ1,ConductaPQ2," _
                             & "ConductaSQ1,ConductaSQ2," _
                             & "ConductaTQ1,ConductaTQ2," _
                             & "ExamenSQ,Supletorio,"
                 End Select
                 sSQL = sSQL & "CodMat,Profesor,Item,CodigoU " _
                      & "FROM Asiento_N "
            End Select
       End If
       sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND CodMat = '" & CodMat & "' " _
            & "AND CodigoU = '" & CodigoUsuario & "' " _
            & "ORDER BY Id_No "
   Else
      sSQL = ""
   End If
  End With
End Sub

Public Sub Imprimir_Notas_Profesor(Datas As Adodc, _
                                   FormaImp As Byte, _
                                   SizeLetra As Integer, _
                                   Optional EsCampoCorto As Boolean)
Dim IdCampo As Integer
On Error GoTo Errorhandler

Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoArialNarrow, Orientacion_Pagina, EsCampoCorto
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
If .RecordCount > 0 Then
    .MoveFirst
     EncabezadoData Datas
     Printer.FontSize = SizeLetra
     Printer.FontName = TipoArialNarrow
     Do While Not .EOF
        'MsgBox Printer.FontName
        For IdCampo = 0 To .fields.Count - 1
            Select Case .fields(IdCampo).Name
              Case "CodMat", "Item", "CodigoU", "Codigo": ' no hacer nada
              Case Else
                   PrinterFields Ancho(IdCampo), PosLinea, .fields(IdCampo)
            End Select
        Next IdCampo
        'PrinterAllFields CantCampos, PosLinea, Datas, True, False
        PosLinea = PosLinea + 0.36
        If Cuadricula Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Gris
           PosLinea = PosLinea + 0.05
        End If
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
           Printer.FontName = TipoArialNarrow
        End If
       .MoveNext
     Loop
End If
End With
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos), Negro, True
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Imprimir_Mejores_Alumnos(Datas As Adodc, _
                                   FormaImp As Byte, _
                                   SizeLetra As Integer, _
                                   Optional EsCampoCorto As Boolean)
Dim IdCampo As Integer
On Error GoTo Errorhandler

Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoArialNarrow, Orientacion_Pagina, EsCampoCorto
Ancho(0) = 1.5
Ancho(1) = 14
Ancho(2) = 16.5
Ancho(CantCampos) = 19
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
If .RecordCount > 0 Then
    .MoveFirst
     EncabezadoData Datas
     Printer.FontSize = SizeLetra
     Printer.FontName = TipoArialNarrow
     Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
     PosLinea = PosLinea + 0.05
     Contador = 0
     Do While Not .EOF
        Contador = Contador + 1
        PrinterVariables Ancho(0), PosLinea, Format$(Contador, "00") & ".-"
        PrinterFields Ancho(0) + 0.6, PosLinea, .fields("Estudiante")
        PrinterFields Ancho(1), PosLinea, .fields("Curso")
        PrinterFields Ancho(2), PosLinea, .fields("Promedio")
        PosLinea = PosLinea + 0.4
        Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
        PosLinea = PosLinea + 0.05
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
           Printer.FontName = TipoArialNarrow
        End If
       .MoveNext
     Loop
End If
End With
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos), Negro, True
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Imprimir_Acta_Matricula(NombreAlumno As String)
Dim DataReg As ADODB.Recordset
Dim CadCta As String
Dim CadCtaSup As String
Dim NuevoDoc As Boolean
Dim MesActual As Byte
Dim NumeroLineas As Single
Dim NumFile As Long
Dim LineaDeTexto As String
Dim Matricula As String
Dim Folio As String
On Error GoTo Errorhandler

 HoraSistema = Time
 CodigoCliente = Ninguno
 NivelNo = Ninguno
 Codigo = Ninguno
 Codigo1 = Ninguno
 Codigo2 = Ninguno
 Codigo3 = Ninguno
 Codigo4 = Ninguno
 Matricula = Ninguno
 Folio = Ninguno
 CodigoB = Ninguno
 CodigoP = Ninguno
 CodigoL = Ninguno
 Cta_Sup = Ninguno
'Activamos el espacio de consulta
 Set DataReg = New ADODB.Recordset
 DataReg.CursorType = adOpenStatic
 DataReg.CursorLocation = adUseClient
 
 sSQL = "SELECT C.Grupo,C.DirNumero,C.Telefono,C.Celular,C.Direccion,C.Ciudad,C.Actividad,CM.* " _
      & "FROM Clientes As C,Clientes_Matriculas As CM " _
      & "WHERE C.Cliente = '" & NombreAlumno & "' " _
      & "AND CM.Periodo = '" & Periodo_Contable & "' " _
      & "AND CM.Item = '" & NumEmpresa & "' " _
      & "AND CM.Codigo = C.Codigo "
 DataReg.open sSQL, AdoStrCnn, , , adCmdText
 With DataReg
  If .RecordCount > 0 Then
      CodigoCliente = .fields("Codigo")
      Codigo = .fields("Representante_Alumno")
      Codigo1 = .fields("Nombre_Padre")
      Codigo2 = .fields("Nombre_Madre")
      Codigo3 = .fields("Profesion_P")
      Codigo4 = .fields("Profesion_M")
      NivelNo = .fields("Grupo")
      Cta_Sup = CodigoCuentaSup(NivelNo)
      Matricula = Format$(.fields("Matricula_No"), "000000000")
      Folio = Format$(.fields("Folio_No"), "000000000")
      CICliente = .fields("Matricula_No")
      DireccionCli = .fields("Direccion")
      DirCliente = ""
      If Len(.fields("Domicilio")) > 1 Then DirCliente = DirCliente & .fields("Domicilio")
      If Len(.fields("Lugar_Trabajo_P")) > 1 Then DirCliente = DirCliente & " - " & .fields("Lugar_Trabajo_P")
      If Val(.fields("Celular")) > 0 Then DirCliente = DirCliente & " - " & .fields("Celular")
      If DirCliente = "" Then DirCliente = Ninguno
      Codigos = .fields("Telefono")
      CodigoL = .fields("Ciudad")
      TextoProc = .fields("Actividad")
      CodigoA = .fields("Actividad")
      CodigoL = .fields("Actividad")
      'MsgBox TextoProc
      Mifecha = FechaStrg(.fields("Fecha_N"))
  End If
 End With
 DataReg.Close
 sSQL = "SELECT * " _
      & "FROM Catalogo_Estudiantil " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "ORDER BY CodigoE "
 DataReg.open sSQL, AdoStrCnn, , , adCmdText
 With DataReg
  If .RecordCount > 0 Then
     .MoveFirst
     .Find ("CodigoE = '" & NivelNo & "' ")
      If Not .EOF Then CodigoB = .fields("Detalle")
     .MoveFirst
     .Find ("CodigoE = '" & Cta_Sup & "' ")
      If Not .EOF Then CodigoP = .fields("Detalle")
  End If
 End With
 DataReg.Close
'''     Director
'''     Secretario1
'''     Rector
'''     Secretario2
'''     NombreComercia
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION DE ACTA DE MATRICULA"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   Escala_Centimetro 1, TipoTimes, 8
   ReDim Ancho(2) As Single
   InicioX = 0.5: InicioY = 1
   Ancho(0) = 1.5: AnchoPapel = 18
   Pagina = 1: Documento = 1
   MensajeEncabData = Anio_Lectivo
   SQLMsg1 = "A C T A     D E     M A T R I C U L A"
   SQLMsg2 = ""
   SQLMsg3 = ""
   EncabezadoSimple 1.5, 19
   PosLinea = PosLinea + 0.1
   Printer.Line (2, PosLinea)-(18, PosLinea), Negro
   PosLinea = PosLinea + 0.05
   Printer.Line (2, PosLinea)-(18, PosLinea), Negro
   PosLinea = PosLinea + 0.5
   NumeroLineas = PosLinea
   Printer.FontBold = True
   Printer.FontSize = 10
   Printer.FontName = TipoCourier
  'Foto del Alumno
   Printer.Line (14.5, PosLinea)-(18, PosLinea + 4), Negro, B
   Printer.Line (14.55, PosLinea + 0.05)-(17.95, PosLinea + 3.95), Negro, B
   Printer.Line (14.6, PosLinea + 0.1)-(17.9, PosLinea + 3.9), Negro, B
   
   RutaDestino = RutaSistema & "\FOTOS\" & CodigoCliente & ".JPG"
   PrinterPaint RutaDestino, 14.7, PosLinea + 0.2, 3.1, 3.6
   PrinterTexto 2, PosLinea, "FOLIO:"
   PrinterTexto 2, PosLinea + 0.8, "MATRICULA:"
   PrinterTexto 2, PosLinea + 1.6, "EL ALUMNO(A):"
   PrinterTexto 2, PosLinea + 2.4, "NACIDO(A) EN:"
   PrinterTexto 2, PosLinea + 3.2, "CON FECHA DE NACIMIENTO:"
   PrinterTexto 2, PosLinea + 4, "QUEDA MATRICULADO(A) EN:"
   PrinterTexto 2, PosLinea + 4.8, "UNA VEZ QUE HA LLENADO LOS REQUISITOS DE LEY."
   PrinterTexto 2, PosLinea + 5.6, "ES HIJO(A) DE:"
   PrinterTexto 2, PosLinea + 6.4, "DE PROFESION:"
   PrinterTexto 2, PosLinea + 7.2, "OCUPACIÓN:"
   PrinterTexto 2, PosLinea + 8, "Y DE:"
   PrinterTexto 2, PosLinea + 8.8, "DE PROFESION:"
   PrinterTexto 2, PosLinea + 9.6, "OCUPACION:"
   PrinterTexto 2, PosLinea + 10.4, "EL REPRESENTANTE ES:"
   PrinterTexto 2, PosLinea + 11.2, "DOMICILIO:"
   PrinterTexto 2, PosLinea + 12, "TELEFONOS:"
   Printer.FontBold = False
  'Lineas de Compromisos:
   NumFile = FreeFile
   PFil = PosLinea + 12.5
   LineaDeTexto = RutaSistema & "\DOCUMENT\ActaMatricula.txt"
   Open LineaDeTexto For Input As #NumFile
     Do While Not EOF(NumFile)
        Line Input #NumFile, LineaDeTexto
        LineaDeTexto = LineaDeTexto & " "
        PrinterTexto 2, PFil, LineaDeTexto
        PFil = PFil + 0.4
     Loop
   Close #NumFile
   PrinterTexto 2, PosLinea + 20, "_____________________"
   PrinterTexto 9, PosLinea + 20, "_____________________"
   PrinterTexto 15, PosLinea + 20, "_____________________"
   PrinterTexto 2, PosLinea + 20.45, "   REPRESENTANTE"
   PrinterTexto 9, PosLinea + 20.45, "     RECTOR(A)"
   PrinterTexto 15, PosLinea + 20.45, "    SECRETARIA"
   'PosLinea = NumeroLineas
   PrinterTexto 6, PosLinea, Folio
   PrinterTexto 6, PosLinea + 0.8, Matricula
   PrinterTexto 6, PosLinea + 1.6, NombreAlumno
   PrinterTexto 6, PosLinea + 2.4, CodigoL
   PrinterTexto 7.5, PosLinea + 3.2, Mifecha
   PrinterTexto 7.5, PosLinea + 4, DireccionCli
   PrinterTexto 6, PosLinea + 5.6, Codigo1
   PrinterTexto 6, PosLinea + 6.4, Codigo3
   PrinterTexto 6, PosLinea + 7.2, TextoProc
   
   PrinterTexto 6, PosLinea + 8, Codigo2
   PrinterTexto 6, PosLinea + 8.8, Codigo4
   PrinterTexto 6, PosLinea + 9.6, Codigo4
   PrinterTexto 7, PosLinea + 10.4, Codigo
   PrinterTexto 7, PosLinea + 11.2, DirCliente
   PrinterTexto 7, PosLinea + 12, Codigos
   
   RatonNormal
   MensajeEncabData = ""
   Printer.EndDoc
   Exit Sub
Errorhandler:
             RatonNormal
             ErrorDeImpresion
             Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Imprimir_Acta_Matricula_Periodos(NombreAlumno As String, Optional TGrupo_No As String)
Dim DataReg As ADODB.Recordset
Dim CadCta As String
Dim CadCtaSup As String
Dim NuevoDoc As Boolean
Dim MesActual As Byte
Dim NumeroLineas As Single
Dim NumFile As Long
Dim LineaDeTexto As String
On Error GoTo Errorhandler
 HoraSistema = Time
 CodigoCliente = Ninguno
 NivelNo = Ninguno
 Codigo = Ninguno
 Codigo1 = Ninguno
 Codigo2 = Ninguno
 Codigo3 = Ninguno
 Codigo4 = Ninguno
 CodigoB = Ninguno
 CodigoP = Ninguno
 CodigoL = Ninguno
 Cta_Sup = Ninguno
'Activamos el espacio de consulta
 Set DataReg = New ADODB.Recordset
 DataReg.CursorType = adOpenStatic
 DataReg.CursorLocation = adUseClient
 
 sSQL = "SELECT C.Cliente,C.Grupo,C.DirNumero,C.Telefono,C.Celular,C.Direccion,C.Ciudad,C.Actividad,CM.* " _
      & "FROM Clientes As C,Clientes_Matriculas As CM "
 If TGrupo_No <> "" Then
    sSQL = sSQL & "WHERE C.Grupo = '" & TGrupo_No & "' "
 Else
    sSQL = sSQL & "WHERE C.Cliente = '" & NombreAlumno & "' "
 End If
 sSQL = sSQL & "AND CM.Periodo = '" & Periodo_Contable & "' " _
      & "AND CM.Item = '" & NumEmpresa & "' " _
      & "AND CM.Codigo = C.Codigo " _
      & "ORDER BY C.Grupo,C.Cliente "
 DataReg.open sSQL, AdoStrCnn, , , adCmdText
 With DataReg
  If .RecordCount > 0 Then
      Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
      If TGrupo_No <> "" Then
         Titulo = "IMPRESION DE ACTAS DE MATRICULAS"
      Else
         Titulo = "IMPRESION DE ACTA DE MATRICULA"
      End If
      Bandera = False
      SetPrinters.Show 1
      If PonImpresoraDefecto(SetNombrePRN) Then
         Escala_Centimetro 1, TipoTimes, 8
         ReDim Ancho(2) As Single
         InicioX = 0.5: InicioY = 1
         Ancho(0) = 1.5: AnchoPapel = 18
         Pagina = 0: Documento = 1
         'EncabezadoSimple 1, 18
         Do While Not .EOF
         
         MensajeEncabData = "AÑO LECTIVO " & Anio_Lectivo
         SQLMsg1 = "A C T A   D E   M A T R I C U L A    No. " & .fields("Matricula_No")
         SQLMsg2 = ""
         SQLMsg3 = ""
         SQLMsg4 = ""
         'If Val(MidStrg(.Fields("Grupo"), 1, 1)) > 1 Then
            Encabezado_Institucion 1, 19
            PosLinea = 5
         'Else
         '   Printer.FontSize = 16
         '   PrinterTexto CentrarTexto(MensajeEncabData, 19), 3.5, MensajeEncabData
         '   PrinterTexto CentrarTexto(SQLMsg1, 19), 4.5, SQLMsg1
         '   PosLinea = 8
         'End If
         SQLMsg1 = ""
         PosLinea = PosLinea + 2
         NumeroLineas = PosLinea
         Printer.FontItalic = False
         Printer.FontBold = False
         Printer.FontSize = 11
         Printer.FontName = TipoArial
        'Foto del Alumno
'''         Printer.Line (14.5, PosLinea)-(18, PosLinea + 4), Negro, B
'''         Printer.Line (14.55, PosLinea + 0.05)-(17.95, PosLinea + 3.95), Negro, B
'''         Printer.Line (14.6, PosLinea + 0.1)-(17.9, PosLinea + 3.9), Negro, B
'''         RutaDestino = RutaSistema & "\FOTOS\" & CodigoCliente & ".JPG"
'''         PrinterPaint RutaDestino, 14.7, PosLinea + 0.2, 3.1, 3.6
        'Datos del Alumno

         Cadena = "En la " & Institucion1 & " " & Institucion2 & " " _
                & "De Conformidad con el Reglamento a la Ley Orgánica de Educación Intercultural, se matricula la estudiante: "
         NumeroLineas = PrinterLineasMayor(2, PosLinea, Cadena, 17.5, 0.5)
         PosLinea = PosLinea_Aux + 0.6
         Printer.FontSize = 14
         Printer.FontBold = True
         PrinterCentrarTexto 20, PosLinea, UCaseStrg(.fields("Cliente"))
         Printer.FontBold = False
         PosLinea = PosLinea + 1
         Printer.FontSize = 11
         Printer.FontUnderline = True
         Printer.FontBold = True
         Cadena = Leer_Datos_del_Curso(.fields("Grupo"))
         NumeroLineas = PrinterLineasMayor(2, PosLinea, TrimStrg(Cadena), 17.5, 0.5)
         PosLinea = PosLinea_Aux + 0.6
         Printer.FontUnderline = False
         Printer.FontBold = False
        '=============================================================================================
         Printer.FontBold = True
         PrinterTexto 2, PosLinea, "Lugar de Nacimiento:"
         PrinterTexto 15, PosLinea, "Fecha:"
         Printer.FontBold = False
         PrinterTexto 6.5, PosLinea, .fields("Lugar_Nac")
         PrinterTexto 16.5, PosLinea, .fields("Fecha_N")
         PosLinea = PosLinea + 0.5
         Printer.FontBold = True
         PrinterTexto 2, PosLinea, "Dirección Domicilio:"
         Printer.FontBold = False
         PrinterTexto 6.5, PosLinea, .fields("Domicilio")
         PosLinea = PosLinea + 0.5
         Printer.FontBold = True
         PrinterTexto 2, PosLinea, "Teléfonos:"
         PrinterTexto 13, PosLinea, "Nacionalidad:"
         Printer.FontBold = False
         PrinterTexto 6.5, PosLinea, .fields("Telefono")
         PrinterTexto 16, PosLinea, .fields("Nacionalidad")
         PosLinea = PosLinea + 0.5
         Printer.FontBold = True
         PrinterTexto 2, PosLinea, "Plantel que Proviene:"
         Printer.FontBold = False
         PrinterTexto 6.5, PosLinea, .fields("Procedencia")
         PosLinea = PosLinea + 0.8
         Printer.FontSize = 13
         Printer.FontBold = True
         PrinterTexto 2, PosLinea, "DATOS DEL PADRE"
         Printer.FontBold = False
         Printer.FontSize = 11
         PosLinea = PosLinea + 0.6
         Printer.FontBold = True
         PrinterTexto 2, PosLinea, "Nombre:"
         PrinterTexto 15, PosLinea, "Teléfono:"
         Printer.FontBold = False
         PrinterTexto 4.2, PosLinea, .fields("Nombre_Padre")
         PrinterTexto 17, PosLinea, .fields("Telefono_Trabajo_P")
         PosLinea = PosLinea + 0.5
         Printer.FontBold = True
         PrinterTexto 2, PosLinea, "Nacionalidad:"
         PrinterTexto 10, PosLinea, "Ocupación:"
         Printer.FontBold = False
         PrinterTexto 4.7, PosLinea, .fields("Nacionalidad_P")
         PrinterTexto 12.5, PosLinea, .fields("Profesion_P")
         PosLinea = PosLinea + 0.8
         Printer.FontSize = 13
         Printer.FontBold = True
         PrinterTexto 2, PosLinea, "DATOS DE LA MADRE"
         Printer.FontBold = False
         Printer.FontSize = 11
         PosLinea = PosLinea + 0.6
         Printer.FontBold = True
         PrinterTexto 2, PosLinea, "Nombre:"
         PrinterTexto 15, PosLinea, "Teléfono:"
         Printer.FontBold = False
         PrinterTexto 4.2, PosLinea, .fields("Nombre_Madre")
         PrinterTexto 17, PosLinea, .fields("Telefono_Trabajo_M")
         PosLinea = PosLinea + 0.5
         Printer.FontBold = True
         PrinterTexto 2, PosLinea, "Nacionalidad:"
         PrinterTexto 10, PosLinea, "Ocupación:"
         Printer.FontBold = False
         PrinterTexto 4.7, PosLinea, .fields("Nacionalidad_M")
         PrinterTexto 12.5, PosLinea, .fields("Profesion_M")
         PosLinea = PosLinea + 0.8
         Printer.FontSize = 13
         Printer.FontBold = True
         PrinterTexto 2, PosLinea, "DATOS DEL REPRESENTANTE"
         Printer.FontBold = False
         Printer.FontSize = 11
         PosLinea = PosLinea + 0.6
         Printer.FontBold = True
         PrinterTexto 2, PosLinea, "Nombre:"
         PrinterTexto 15, PosLinea, "Teléfono:"
         Printer.FontBold = False
         PrinterTexto 4.2, PosLinea, .fields("Representante_Alumno")
         PrinterTexto 17, PosLinea, .fields("Telefono_R")
         PosLinea = PosLinea + 0.5
         Printer.FontBold = True
         PrinterTexto 2, PosLinea, "Telefono:"
         PrinterTexto 10, PosLinea, "Ocupación:"
         Printer.FontBold = False
         PrinterTexto 4.2, PosLinea, .fields("Celular")
         PrinterTexto 12.5, PosLinea, .fields("Profesion_R")
         PosLinea = PosLinea + 1
         Cadena = "El infrascrito, representante de la estudiante matriculada, declara que se encuentra conforme " _
                & "con los datos que anteceden y firma sometiéndose a las disposiciones del citado reglamento."
         NumeroLineas = PrinterLineasMayor(2, PosLinea, Cadena, 17.5, 0.5)
         'PosLinea = PosLinea + (0.5 * NumeroLineas)
         PosLinea = PosLinea_Aux + 0.6
         PrinterTexto 2, PosLinea, "Lugar y Fecha:"
         Printer.FontBold = False
         PrinterTexto 5, PosLinea, FechaStrgCiudad(FechaComp)
         PosLinea = PosLinea + 3
         Printer.FontBold = True
         PrinterTexto 2, PosLinea, "_____________________"
         PrinterTexto 14.5, PosLinea, "_____________________"
         PrinterTexto 2, PosLinea + 0.45, "   REPRESENTANTE"
         PrinterTexto 15, PosLinea + 0.45, "    SECRETARIA"
         Printer.FontBold = False
         Printer.NewPage
         .MoveNext
         Loop
         RatonNormal
         MensajeEncabData = ""
         Printer.EndDoc
         DataReg.Close
      Exit Sub
Errorhandler:
             RatonNormal
             ErrorDeImpresion
             Exit Sub
     End If
   Else
    RatonNormal
   End If
 End With
End Sub

Public Sub Imprimir_Hoja_Matricula(NombreAlumno As String)
Dim DataReg As ADODB.Recordset
Dim CadCta As String
Dim CadCtaSup As String
Dim NuevoDoc As Boolean
Dim MesActual As Byte
Dim NumeroLineas As Single
Dim AnchoCentrar As Single
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION DE HOJA DE MATRICULA"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   Escala_Centimetro 1, TipoTimes, 8
   ReDim Ancho(2) As Single
   Set DataReg = New ADODB.Recordset
   DataReg.CursorType = adOpenDynamic
   DataReg.CursorLocation = adUseClient
   sSQL = "SELECT C.Cliente,C.Grupo,C.DirNumero,C.Telefono,C.Celular,CC.Descripcion,C.Ciudad,CM.* " _
        & "FROM Clientes As C,Clientes_Matriculas As CM,Catalogo_Cursos As CC " _
        & "WHERE C.Cliente = '" & NombreAlumno & "' " _
        & "AND CM.Periodo = '" & Periodo_Contable & "' " _
        & "AND CM.Item = '" & NumEmpresa & "' " _
        & "AND CM.Codigo = C.Codigo " _
        & "AND CM.Item = CC.Item " _
        & "AND CM.Periodo = CC.Periodo "
   sSQL = CompilarSQL(sSQL)
   DataReg.open sSQL, AdoStrCnn, , , adCmdText
   With DataReg
    If .RecordCount > 0 Then
        Cadena = Leer_Datos_del_Curso(.fields("Grupo_No"))
        Codigo4 = MidStrg(.fields("Grupo_No"), 1, 4)
   InicioX = 0.5: InicioY = 1
   Ancho(0) = 1.5: AnchoPapel = 18
   Pagina = 1: Documento = 1
   MensajeEncabData = Anio_Lectivo
   SQLMsg1 = "H O J A     D E     M A T R I C U L A"
   SQLMsg2 = ""
   SQLMsg3 = ""
   MensajeEncabData = Anio_Lectivo
   EncabezadoSimple 1, 18
   PosLinea = PosLinea + 0.1
   Printer.Line (1, PosLinea)-(18, PosLinea), Negro
   PosLinea = PosLinea + 0.05
   Printer.Line (1, PosLinea)-(18, PosLinea), Negro
   PosLinea = PosLinea + 1
   NumeroLineas = PosLinea
   Printer.FontBold = True
   Printer.FontSize = 14
   Printer.FontName = TipoTimes
  'Foto del Alumno(a)
   Printer.Line (14.5, PosLinea + 1)-(18, PosLinea + 5), Negro, B
   Printer.Line (14.55, PosLinea + 1.05)-(17.95, PosLinea + 4.95), Negro, B
   Printer.Line (14.6, PosLinea + 1.1)-(17.9, PosLinea + 4.9), Negro, B
   RutaDestino = RutaSistema & "\FOTOS\" & CodigoCliente & ".JPG"
   PrinterPaint RutaDestino, 14.7, PosLinea + 1.2, 3.1, 3.6
  'Datos del Alumno(a)
   Printer.FontSize = 14
   PosLinea = PosLinea - 0.2
   PrinterTexto 2, PosLinea, "EL ALUMNO(A):"
   Printer.FontSize = 12
   PosLinea = PosLinea + 2
   PrinterTexto 2, PosLinea, "CURSO:"
   PosLinea = PosLinea + 2
   Select Case Codigo4
     Case "1.00" To "2.99"
          PrinterTexto 2, PosLinea, "CICLO:"
          PosLinea = PosLinea + 2
     Case "3.00" To "3.99"
          PrinterTexto 2, PosLinea, "ESPECIALIDAD:"
          PosLinea = PosLinea + 2
   End Select
   PrinterTexto 2, PosLinea, "MATRICULA No."
   PosLinea = PosLinea + 2
   PrinterTexto 2, PosLinea, "FOLIO No."
   Printer.FontBold = False
   AnchoCentrar = SetPapelAncho / 2
   
   PosColumna = (AnchoCentrar / 2) - (Printer.TextWidth(Director) / 2)
   PrinterTexto PosColumna, 22, String$(Len(Director), "_")
   PosColumna = (AnchoCentrar / 2) - (Printer.TextWidth(Secretario1) / 2)
   PrinterTexto AnchoCentrar + PosColumna, 22, String$(Len(Secretario1), "_")
   Select Case Codigo4
     Case "0.00" To "1.99"
          PosColumna = (AnchoCentrar / 2) - (Printer.TextWidth(Director) / 2)
          PrinterTexto PosColumna, 22.5, Director
          PosColumna = (AnchoCentrar / 2) - (Printer.TextWidth(Secretario1) / 2)
          PrinterTexto AnchoCentrar + PosColumna, 22.5, Secretario1
          Printer.FontBold = True
          PosColumna = (AnchoCentrar / 2) - (Printer.TextWidth(TextoRector) / 2)
          PrinterTexto PosColumna, 23, TextoRector
          PosColumna = (AnchoCentrar / 2) - (Printer.TextWidth(TextoSecretario1) / 2)
          PrinterTexto AnchoCentrar + PosColumna, 23, TextoSecretario1
     Case "2.00" To "3.99"
          PosColumna = (AnchoCentrar / 2) - (Printer.TextWidth(Rector) / 2)
          PrinterTexto PosColumna, 22.5, Rector
          PosColumna = (AnchoCentrar / 2) - (Printer.TextWidth(Secretario2) / 2)
          PrinterTexto AnchoCentrar + PosColumna, 22.5, Secretario2
          Printer.FontBold = True
          PosColumna = (AnchoCentrar / 2) - (Printer.TextWidth(TextoRector) / 2)
          PrinterTexto PosColumna, 23, TextoRector
          PosColumna = (AnchoCentrar / 2) - (Printer.TextWidth(TextoSecretario2) / 2)
          PrinterTexto AnchoCentrar + PosColumna, 23, TextoSecretario2
   End Select
   Printer.FontSize = 9
   PrinterTexto 2, 25, "NOTA: DOCUMENTO NO VALIDO SIN LA FIRMA Y SELLO DE LA INSTITUCION"
   Printer.FontBold = False
   Printer.FontUnderline = False
        PosLinea = NumeroLineas
        Printer.FontSize = 14
        PosLinea = PosLinea - 0.2
        PrinterTexto 6, PosLinea, .fields("Cliente")
        Printer.FontSize = 12
        PosLinea = PosLinea + 2
        'MsgBox .Fields("Nivel")
        PrinterTexto 6, PosLinea, Dato_Curso.Descripcion
        PosLinea = PosLinea + 2
        Select Case Codigo4
          Case "1.00" To "1.99"
''               Cadena = SinEspaciosDer(.Fields("Especialidad"))
''               Cadena = MidStrg(.Fields("Nivel"), 1, Len(.Fields("Nivel")) - Len(Cadena))
               PrinterTexto 6, PosLinea, Dato_Curso.Especialidad
               PosLinea = PosLinea + 2
        
          Case "2.00" To "3.99"
''               Cadena = SinEspaciosDer(.Fields("Especialidad"))
''               Cadena = MidStrg(.Fields("Nivel"), 1, Len(.Fields("Nivel")) - Len(Cadena))
               PrinterTexto 6, PosLinea, Dato_Curso.Especialidad
               PosLinea = PosLinea + 2
        End Select
        PrinterTexto 6, PosLinea, Format$(.fields("Matricula_No"), "000000000")
        PosLinea = PosLinea + 2
        PrinterTexto 6, PosLinea, Format$(.fields("Folio_No"), "000000000")
        PosLinea = PosLinea + 2
        PrinterTexto 2, 17, UCaseStrg(NombreCiudad) & ", " & FechaStrg(.fields("Fecha_M"))
  End If
 End With
 DataReg.Close
   RatonNormal
   MensajeEncabData = ""
   Printer.EndDoc
   Exit Sub
Errorhandler:
             RatonNormal
             ErrorDeImpresion
             Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Imprimir_Hoja_Matricula_Periodos(NombreAlumno As String)
Dim DataReg As ADODB.Recordset
Dim CadCta As String
Dim CadCtaSup As String
Dim NuevoDoc As Boolean
Dim MesActual As Byte
Dim NumeroLineas As Single
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION DE HOJA DE MATRICULA"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   Escala_Centimetro 1, TipoTimes, 8
   ReDim Ancho(2) As Single
   Set DataReg = New ADODB.Recordset
   DataReg.CursorType = adOpenDynamic
   DataReg.CursorLocation = adUseClient
   sSQL = "SELECT C.Cliente,C.Grupo,C.DirNumero,C.Telefono,C.Celular,CC.Descripcion,C.Ciudad,CM.* " _
        & "FROM Clientes As C,Clientes_Matriculas As CM,Catalogo_Cursos As CC " _
        & "WHERE C.Cliente = '" & NombreAlumno & "' " _
        & "AND CM.Periodo = '" & Periodo_Contable & "' " _
        & "AND CM.Item = '" & NumEmpresa & "' " _
        & "AND CM.Codigo = C.Codigo " _
        & "AND CM.Item = CC.Item " _
        & "AND CM.Periodo = CC.Periodo "
   sSQL = CompilarSQL(sSQL)
   DataReg.open sSQL, AdoStrCnn, , , adCmdText
   With DataReg
    If .RecordCount > 0 Then
        Codigo4 = MidStrg(.fields("Grupo_No"), 1, 4)
   InicioX = 0.5: InicioY = 1
   Ancho(0) = 1.5: AnchoPapel = 18
   Pagina = 0: Documento = 1
   MensajeEncabData = Anio_Lectivo
   SQLMsg1 = "H O J A     D E     M A T R I C U L A"
   SQLMsg2 = ""
   SQLMsg3 = ""
   MensajeEncabData = Anio_Lectivo
   EncabezadoSimple 1, 18
   PosLinea = PosLinea + 0.1
   Printer.Line (1, PosLinea)-(18, PosLinea), Negro
   PosLinea = PosLinea + 0.05
   Printer.Line (1, PosLinea)-(18, PosLinea), Negro
   PosLinea = PosLinea + 1
   NumeroLineas = PosLinea
   Printer.FontBold = True
   Printer.FontSize = 14
   Printer.FontName = TipoTimes
  'Foto del Alumno(a)
   Printer.Line (14.5, PosLinea + 1)-(18, PosLinea + 5), Negro, B
   Printer.Line (14.55, PosLinea + 1.05)-(17.95, PosLinea + 4.95), Negro, B
   Printer.Line (14.6, PosLinea + 1.1)-(17.9, PosLinea + 4.9), Negro, B
   RutaDestino = RutaSistema & "\FOTOS\" & CodigoCliente & ".JPG"
   PrinterPaint RutaDestino, 14.7, PosLinea + 1.2, 3.1, 3.6
  'Datos del Alumno(a)
   Printer.FontSize = 14
   PosLinea = PosLinea - 0.2
   PrinterTexto 1, PosLinea, "EL ALUMNO(A):"
   Printer.FontSize = 12
   PosLinea = PosLinea + 2
   PrinterTexto 1, PosLinea, "CURSO:"
   PosLinea = PosLinea + 2
   Select Case Codigo4
     Case "2.00" To "3.99"
          PrinterTexto 1, PosLinea, "ESPECIALIDAD:"
          PosLinea = PosLinea + 2
   End Select
   PrinterTexto 1, PosLinea, "MATRICULA No."
   PosLinea = PosLinea + 2
   PrinterTexto 1, PosLinea, "FOLIO No."
   Printer.FontBold = False
   PrinterTexto 5, 22, "_____________________"
   PrinterTexto 12, 22, "_____________________"
   Select Case Codigo4
     Case "0.00" To "1.99"
          PrinterTexto 5, 22.5, Director
          PrinterTexto 12, 22.5, Secretario1
          Printer.FontBold = True
          PrinterTexto 5, 23, "DIRECTOR(A)"
          PrinterTexto 12, 23, "SECRETARIO(A)"
     Case "2.00" To "3.99"
          PrinterTexto 5, 22.5, Rector
          PrinterTexto 12, 22.5, Secretario2
          Printer.FontBold = True
          PrinterTexto 5, 23, "RECTOR(A)"
          PrinterTexto 12, 23, "SECRETARIO(A)"
   End Select
   Printer.FontSize = 9
   PrinterTexto 1.5, 25, "NOTA: DOCUMENTO NO VALIDO SIN LA FIRMA Y SELLO DE LA INSTITUCION"
   Printer.FontBold = False
   Printer.FontUnderline = False
        PosLinea = NumeroLineas
        Printer.FontSize = 14
        PosLinea = PosLinea - 0.2
        PrinterTexto 5, PosLinea, .fields("Cliente")
        Printer.FontSize = 12
        PosLinea = PosLinea + 2
        'MsgBox .Fields("Nivel")
        PrinterTexto 5, PosLinea, .fields("Nivel")
        PosLinea = PosLinea + 2
        Select Case Codigo4
          Case "2.00" To "3.99"
''               Cadena = SinEspaciosDer(.Fields("Especialidad"))
''               Cadena = MidStrg(.Fields("Nivel"), 1, Len(.Fields("Nivel")) - Len(Cadena))
               PrinterTexto 5, PosLinea, .fields("Especialidad")
               PosLinea = PosLinea + 2
        End Select
        PrinterTexto 5, PosLinea, Format$(.fields("Matricula_No"), "000000000")
        PosLinea = PosLinea + 2
        PrinterTexto 5, PosLinea, Format$(.fields("Folio_No"), "000000000")
        PosLinea = PosLinea + 2
        PrinterTexto 1, 17, UCaseStrg(NombreCiudad) & ", " & FechaStrg(.fields("Fecha_M"))
  End If
 End With
 DataReg.Close
   RatonNormal
   MensajeEncabData = ""
   Printer.EndDoc
   Exit Sub
Errorhandler:
             RatonNormal
             ErrorDeImpresion
             Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Imprimir_Registro_Matricula(NombreAlumno As String)
Dim DataReg As ADODB.Recordset
Dim CadCta As String
Dim CadCtaSup As String
Dim NuevoDoc As Boolean
Dim MesActual As Byte
Dim NumeroLineas As Single
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION DE HOJA DE MATRICULA"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   Escala_Centimetro 1, TipoTimes, 8
   ReDim Ancho(2) As Single
  'Activamos el espacio de consulta
   Set DataReg = New ADODB.Recordset
   DataReg.CursorType = adOpenStatic
   DataReg.CursorLocation = adUseClient
   
   sSQL = "SELECT C.Cliente,CM.Grupo_No,C.DirNumero,C.Telefono,C.Celular,C.Direccion,C.Ciudad,CM.* " _
        & "FROM Clientes As C,Clientes_Matriculas As CM " _
        & "WHERE C.Cliente = '" & NombreAlumno & "' " _
        & "AND CM.Periodo = '" & Periodo_Contable & "' " _
        & "AND CM.Item = '" & NumEmpresa & "' " _
        & "AND CM.Codigo = C.Codigo "
   DataReg.open sSQL, AdoStrCnn, , , adCmdText
   With DataReg
    If .RecordCount > 0 Then
        Cadena = Leer_Datos_del_Curso(.fields("Grupo_No"))
        NivelNo = .fields("Grupo_No")
    End If
   End With
   
   InicioX = 0.5: InicioY = 0.5
   Ancho(0) = 1.5: AnchoPapel = 20
   Pagina = 1: Documento = 1
   MensajeEncabData = Anio_Lectivo
   SQLMsg1 = "H O J A     D E     M A T R I C U L A"
   SQLMsg2 = ""
   SQLMsg3 = ""
   MensajeEncabData = Anio_Lectivo
   EncabezadoSimple 1, 20
   PosLinea = PosLinea + 0.1
   Printer.Line (1, PosLinea)-(18, PosLinea), Negro
   PosLinea = PosLinea + 0.05
   Printer.Line (1, PosLinea)-(18, PosLinea), Negro
   PosLinea = PosLinea + 0.5
   NumeroLineas = PosLinea
   Printer.FontBold = True
   Printer.FontSize = 14
   Printer.FontName = TipoVerdana
  'Foto del Alumno(a)
   Printer.Line (14.5, PosLinea + 1)-(18, PosLinea + 5), Negro, B
   Printer.Line (14.55, PosLinea + 1.05)-(17.95, PosLinea + 4.95), Negro, B
   Printer.Line (14.6, PosLinea + 1.1)-(17.9, PosLinea + 4.9), Negro, B
   RutaDestino = RutaSistema & "\FOTOS\" & CodigoCliente & ".JPG"
   PrinterPaint RutaDestino, 14.7, PosLinea + 1.2, 3.1, 3.6
  'Datos del Alumno(a)
   Printer.FontSize = 12
   Printer.FontUnderline = True
   PrinterTexto 1, PosLinea + 0, "REGISTRO DE MATRICULA"
   Printer.FontUnderline = False
   Printer.FontSize = 10
   PrinterTexto 1, PosLinea + 1, "MATRICULA No."
   PrinterTexto 1, PosLinea + 1.5, "FOLIO No."
   PrinterTexto 1, PosLinea + 2, "FECHA DE MATRICULA:"
   PrinterTexto 1, PosLinea + 2.5, "ESPECIALIDAD:"
   PrinterTexto 1, PosLinea + 3, "AÑO:"
   PrinterTexto 1, PosLinea + 3.5, "SECCION:"
   PrinterTexto 1, PosLinea + 4, "NIVEL:"
   Printer.FontSize = 12
   Printer.FontUnderline = True
   PrinterTexto 1, PosLinea + 5, "DATOS PERSONALES"
   Printer.FontUnderline = False
   Printer.FontSize = 10
   PrinterTexto 1, PosLinea + 6, "NOMBRES Y APELLIDOS:"
   PrinterTexto 1, PosLinea + 6.5, "CEDULA:"
   PrinterTexto 1, PosLinea + 7, "LUGAR DE NACIMIENTO:"
   PrinterTexto 1, PosLinea + 7.5, "FECHA DE NACIMIENTO:"
   PrinterTexto 1, PosLinea + 8, "DOMICILIO:"
   PrinterTexto 1, PosLinea + 8.5, "TELEFONO:"
   Printer.FontSize = 12
   Printer.FontUnderline = True
   PrinterTexto 1, PosLinea + 9.5, "DATOS FAMILIARES"
   Printer.FontUnderline = False
   Printer.FontSize = 10
   PrinterTexto 1, PosLinea + 10.5, "NOMBRE DEL PADRE:"
   PrinterTexto 1, PosLinea + 11, "NACIONALIDAD:"
   PrinterTexto 1, PosLinea + 11.5, "PROFESION:"
   PrinterTexto 1, PosLinea + 12, "LUGAR DE TRABAJO:"
   PrinterTexto 1, PosLinea + 12.5, "TELEFONO DE TRABAJO:"
   'Printer.Line (1, PosLinea + 13)-(17.9, PosLinea + 13), Negro, B
   PrinterTexto 1, PosLinea + 13.5, "NOMBRE DEL MADRE:"
   PrinterTexto 1, PosLinea + 14, "NACIONALIDAD:"
   PrinterTexto 1, PosLinea + 14.5, "PROFESION:"
   PrinterTexto 1, PosLinea + 15, "LUGAR DE TRABAJO:"
   PrinterTexto 1, PosLinea + 15.5, "TELEFONO DE TRABAJO:"
   'Printer.Line (1, PosLinea + 15.5)-(17.9, PosLinea + 15.5), Negro, B
   PrinterTexto 1, PosLinea + 16.5, "REPRESENTANTE:"
   PrinterTexto 1, PosLinea + 17, "CEDULA DE IDENTIDAD:"
   PrinterTexto 1, PosLinea + 17.5, "TELEFONO:"
   PrinterTexto 1, PosLinea + 18, "OBSERVACIONES:"
   Printer.FontBold = False
   PrinterTexto 3, PosLinea + 21, "________________________"
   PrinterTexto 12, PosLinea + 21, "________________________"
   Select Case NivelNo
     Case "0" To "1.99"
          PrinterTexto 3, PosLinea + 21.5, " " & Director
          PrinterTexto 12, PosLinea + 21.5, " " & Secretario1
          PrinterTexto 3, PosLinea + 22, " " & TextoDirector
          PrinterTexto 12, PosLinea + 22, " " & TextoSecretario1
     Case "2.00" To "3.99"
          PrinterTexto 3, PosLinea + 21.5, " " & Rector
          PrinterTexto 12, PosLinea + 21.5, " " & Secretario2
          PrinterTexto 3, PosLinea + 22, " " & TextoRector
          PrinterTexto 12, PosLinea + 22, " " & TextoSecretario2
     Case "4.00" To "9.99"
          PrinterTexto 3, 21.5, PosLinea + " " & TextoRector
          PrinterTexto 12, 21.5, PosLinea + " " & Secretario3
   End Select
   Printer.FontSize = 9
   PrinterTexto 1.5, PosLinea + 23, "NOTA: DOCUMENTO NO VALIDO SIN LA FIRMA Y SELLO DE LA INSTITUCION"
  'Activamos el espacio de consulta
    ''Set DataReg = New ADODB.Recordset
    ''DataReg.CursorType = adOpenStatic
    ''DataReg.CursorLocation = adUseClient
    ''
    ''sSQL = "SELECT * " _
    ''     & "FROM Catalogo_Estudiantil " _
    ''     & "WHERE Item = '" & NumEmpresa & "' " _
    ''     & "AND Periodo = '" & Periodo_Contable & "' " _
    ''     & "ORDER BY CodigoE "
    ''DataReg.Open sSQL, AdoStrCnn, , , adCmdText
    ''With DataReg
    '' If .RecordCount > 0 Then
    ''    .MoveFirst
    ''    .Find ("CodigoE = '" & NivelNo & "' ")
    ''     If Not .EOF Then CodigoB = .Fields("Detalle")
    ''    .MoveFirst
    ''    .Find ("CodigoE = '" & Cta_Sup & "' ")
    ''     If Not .EOF Then CodigoP = .Fields("Detalle")
    '' End If
    ''End With
    ''DataReg.Close
   Printer.FontSize = 10
   Printer.FontBold = False
   Printer.FontUnderline = False
   
   With DataReg
    If .RecordCount > 0 Then
        PrinterTexto 7, PosLinea + 1, Format$(.fields("Matricula_No"), "0000")
        PrinterTexto 7, PosLinea + 1.5, Format$(.fields("Folio_No"), "0000")
        PrinterTexto 7, PosLinea + 2, FechaStrg(.fields("Fecha_M"))
        PrinterTexto 7, PosLinea + 2.5, Dato_Curso.Especialidad
        PrinterTexto 7, PosLinea + 3, Dato_Curso.Descripcion  ' .Fields("Direccion")
        PrinterTexto 7, PosLinea + 3.5, Dato_Curso.Seccion
        PrinterTexto 7, PosLinea + 4, .fields("Grupo_No")
    
        PrinterTexto 7, PosLinea + 6, .fields("Cliente")
        PrinterTexto 7, PosLinea + 6.5, .fields("CI")
        PrinterTexto 7, PosLinea + 7, .fields("Lugar_Nac")
        PrinterTexto 7, PosLinea + 7.5, FechaStrg(.fields("Fecha_N"))
        PrinterTexto 7, PosLinea + 8, .fields("Domicilio")
        PrinterTexto 7, PosLinea + 8.5, .fields("Telefono")
    
        PrinterTexto 7, PosLinea + 10.5, .fields("Nombre_Padre")
        PrinterTexto 7, PosLinea + 11, .fields("Nacionalidad_P")
        PrinterTexto 7, PosLinea + 11.5, .fields("Profesion_P")
        PrinterTexto 7, PosLinea + 12, .fields("Lugar_Trabajo_P")
        PrinterTexto 7, PosLinea + 12.5, .fields("Telefono_Trabajo_P")
        
        PrinterTexto 7, PosLinea + 13.5, .fields("Nombre_Madre")
        PrinterTexto 7, PosLinea + 14, .fields("Nacionalidad_M")
        PrinterTexto 7, PosLinea + 14.5, .fields("Profesion_M")
        PrinterTexto 7, PosLinea + 15, .fields("Lugar_Trabajo_M")
        PrinterTexto 7, PosLinea + 15.5, .fields("Telefono_Trabajo_M")
        
        PrinterTexto 7, PosLinea + 16.5, .fields("Representante_Alumno")
        PrinterTexto 7, PosLinea + 17, .fields("Cedula_R")
        PrinterTexto 7, PosLinea + 17.5, .fields("Telefono_R")
        PrinterTexto 7, PosLinea + 18, .fields("Observaciones")
        Cta_Sup = CodigoCuentaSup(NivelNo)
  End If
 End With
 DataReg.Close
   RatonNormal
   MensajeEncabData = ""
   Printer.EndDoc
   Exit Sub
Errorhandler:
             RatonNormal
             ErrorDeImpresion
             Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Imprimir_Certificado_Matricula(NombreAlumno As String, Optional TGrupo_No As String)
Dim DataReg As ADODB.Recordset
Dim CadCta As String
Dim CadCtaSup As String
Dim NuevoDoc As Boolean
Dim MesActual As Byte
Dim NumeroLineas As Single
Dim AnchoDib As Single
Dim LogoMinisterio As String

On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION DE HOJA DE MATRICULA"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   Escala_Centimetro 1, TipoTimes, 8
   ReDim Ancho(2) As Single
   InicioX = 0.5: InicioY = 1
   Ancho(0) = 1.5: AnchoPapel = 18
   Pagina = 0: Documento = 1
   
  'Activamos el espacio de consulta
   Set DataReg = New ADODB.Recordset
   DataReg.CursorType = adOpenStatic
   DataReg.CursorLocation = adUseClient
   
   sSQL = "SELECT * " _
        & "FROM Catalogo_Cursos " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "ORDER BY Curso "
   DataReg.open sSQL, AdoStrCnn, , , adCmdText
   
   'MsgBox NivelNo
   
   
   With DataReg
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("Curso = '" & NivelNo & "' ")
        If Not .EOF Then CodigoB = .fields("Descripcion")
       .MoveFirst
       .Find ("Curso = '" & Cta_Sup & "' ")
        If Not .EOF Then CodigoP = .fields("Descripcion")
    End If
   End With
   DataReg.Close
   
   
   Printer.FontBold = False
   Printer.FontUnderline = False
  'Listamos los datos de la Matricula
   sSQL = "SELECT C.Cliente,C.Grupo,C.DirNumero,C.Telefono,C.Celular,C.Direccion,C.Sexo,C.Ciudad,CM.* " _
        & "FROM Clientes As C,Clientes_Matriculas As CM " _
        & "WHERE CM.Periodo = '" & Periodo_Contable & "' " _
        & "AND CM.Item = '" & NumEmpresa & "' "
   If TGrupo_No <> "" Then
      sSQL = sSQL & "AND C.Grupo = '" & TGrupo_No & "' "
   Else
      sSQL = sSQL & "AND C.Cliente = '" & NombreAlumno & "' "
   End If
   sSQL = sSQL & "AND CM.Codigo = C.Codigo " _
        & "ORDER BY C.Grupo,C.Cliente "
   DataReg.open sSQL, AdoStrCnn, , , adCmdText
   With DataReg
    If .RecordCount > 0 Then
        Do While Not .EOF
             NivelNo = .fields("Grupo_No")
             PosLinea = 1
             AnchoDib = 20
             LogoMinisterio = RutaSistema & "\LOGOS\MINISEDU.JPG"
             PrinterPaint LogoTipo, 1, PosLinea, 4, PosLinea + 2
             PrinterPaint LogoMinisterio, 16.5, PosLinea, 2.6, PosLinea + 1.6
             Printer.FontBold = True
             Printer.FontSize = 12
             PrinterCentrarTexto AnchoDib, PosLinea, Institucion1
             PosLinea = PosLinea + 0.6
             Printer.FontSize = 9
             PrinterCentrarTexto AnchoDib, PosLinea, Institucion2
             PosLinea = PosLinea + 0.5
             Printer.FontSize = 8
             Cadena = ULCase(Direccion) & " - Teléfono: " & Telefono1
             PrinterCentrarTexto AnchoDib, PosLinea, Cadena
             PosLinea = PosLinea + 0.5
             Cadena = EmailEmpresa
             If Len(Codigo_AMIE) > 1 Then Cadena = Cadena & String(20, " ") & "Codigo AMIE: " & Codigo_AMIE
             PrinterCentrarTexto AnchoDib, PosLinea, Cadena
             PosLinea = PosLinea + 0.5
             Printer.FontSize = 9
             PrinterCentrarTexto AnchoDib, PosLinea, UCaseStrg(NombreCiudad) & "-" & UCaseStrg(NombrePais)
             Printer.FontSize = 12
             PosLinea = PosLinea + 0.5
             PrinterCentrarTexto AnchoDib, PosLinea, TextoWeb
             Printer.FontSize = 9
             PosLinea = PosLinea + 0.6
             PrinterCentrarTexto AnchoDib, PosLinea, "Año Lectivo " & Anio_Lectivo
             PosLinea = PosLinea + 0.5
             PrinterCentrarTexto AnchoDib, PosLinea, TextoLeyenda
             Printer.FontSize = 12
             PosLinea = PosLinea + 1
             MensajeEncabData = ""
             SQLMsg3 = "M A T R I C U L A  No. " & .fields("Matricula_No")
             PrinterCentrarTexto AnchoDib, PosLinea, SQLMsg3
             PosLinea = PosLinea + 1
             Printer.FontSize = 18
             SQLMsg1 = "C E R T I F I C O"
             PrinterCentrarTexto AnchoDib, PosLinea, SQLMsg1
             PosLinea = PosLinea + 1.5
             Printer.FontSize = 10
             PrinterTexto 15, PosLinea, "JORNADA " & Dato_Curso.Seccion
             PosLinea = PosLinea + 1
             NumeroLineas = PosLinea
             Printer.FontName = TipoTimes
            'Datos del Alumno(a)
             Printer.FontBold = False
             Mifecha = .fields("Fecha_M")
             Cadena = ""
             If .fields("Sexo") = "M" Then
                 Cadena = Cadena & "Ha sido matriculado "
             Else
                 Cadena = Cadena & "Ha sido matriculada "
             End If
             Cadena = Cadena & Leer_Datos_del_Curso(NivelNo) _
                    & " de esta Unidad Educativa, con fecha " & FechaStrg(FechaComp) _
                    & " para el período lectivo " & Anio_Lectivo & " previo el cumplimiento de los requisitos" _
                    & " legales y reglamentarios. Dicha matrícula consta en el Folio " & .fields("Folio_No") _
                    & " del libro No. "
             If .fields("Grupo_No") < "3" Then Cadena = Cadena & "1" Else Cadena = Cadena & "2"
             Printer.FontSize = 12
             Printer.FontBold = False
             If .fields("Sexo") = "M" Then
                 PrinterTexto 2, PosLinea, "Que el Señor "
             Else
                 PrinterTexto 2, PosLinea, "Que la Señorita "
             End If
             Printer.FontBold = True
             Printer.FontSize = 15
             PrinterTexto 5.5, PosLinea, .fields("Cliente")
             Printer.FontSize = 14
             PosLinea = PosLinea + 1
             Printer.FontBold = False
             NumeroLineas = PrinterLineasMayor(2, PosLinea, Cadena, 17.5, 1)
            'PosLinea = PosLinea + (0.5 * NumeroLineas)
             PosLinea = PosLinea_Aux + 0.2
             PosLinea = PosLinea + 1
             PrinterTexto 2, PosLinea, FechaStrgCiudad(FechaComp)
             PosLinea = PosLinea + 3
             PrinterTexto 3, PosLinea, "_____________________"
             PrinterTexto 12, PosLinea, "_____________________"
             PosLinea = PosLinea + 0.6
            'MsgBox NivelNo
             Select Case NivelNo
               Case "0" To "1.99"
                    PrinterTexto 3, PosLinea, "  " & Director
                    PrinterTexto 12, PosLinea, "  " & Secretario1
                    PosLinea = PosLinea + 0.6
                    PrinterTexto 3, PosLinea, "  " & TextoDirector
                    PrinterTexto 12, PosLinea, "  " & TextoSecretario1
               Case "2.00" To "3.99"
                    PrinterTexto 3, PosLinea, "  " & Rector
                    PrinterTexto 12, PosLinea, "  " & Secretario2
                    PosLinea = PosLinea + 0.6
                    PrinterTexto 3, PosLinea, "  " & TextoRector
                    PrinterTexto 12, PosLinea, "  " & TextoSecretario2
               Case "4.00" To "9.99"
                    PrinterTexto 3, PosLinea, "  " & TextoRector
                    PrinterTexto 12, PosLinea, "  " & Secretario3
             End Select
             Printer.FontSize = 9
             Printer.FontBold = False
             Printer.FontUnderline = False
             Printer.NewPage
            .MoveNext
        Loop
    End If
   End With
   DataReg.Close
   RatonNormal
   MensajeEncabData = ""
   Printer.EndDoc
   Exit Sub
Errorhandler:
             RatonNormal
             ErrorDeImpresion
             Exit Sub
Else
    RatonNormal
End If
End Sub

Public Function Numero_De_Matricula(TGrupo_No As String) As Integer
Dim DataReg As ADODB.Recordset
Dim Num_Matric As Integer
 'Activamos el espacio de consulta
  RatonReloj
  Set DataReg = New ADODB.Recordset
  DataReg.CursorType = adOpenStatic
  DataReg.CursorLocation = adUseClient
  Num_Matric = 0
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE FA <> " & Val(adFalse) & " "
  Select Case MidStrg(TGrupo_No, 1, 2)
    Case "1."
         sSQL = sSQL & "AND MidStrg(Grupo,1,2) = '1.' "
    Case "2."
         sSQL = sSQL & "AND MidStrg(Grupo,1,2) = '2.' "
    Case "3."
         sSQL = sSQL & "AND MidStrg(Grupo,1,2) = '3.' "
    Case Else
         sSQL = sSQL & "AND MidStrg(Grupo,1,2) > '3.' "
  End Select
  sSQL = sSQL & "AND ISNUMERIC(DirNumero) <> 0 " _
       & "ORDER BY DirNumero DESC "
  sSQL = CompilarSQL(sSQL)
  DataReg.open sSQL, AdoStrCnn, , , adCmdText
  With DataReg
   If .RecordCount > 0 Then Num_Matric = Val(.fields("DirNumero")) + 1
  End With
  DataReg.Close
  Numero_De_Matricula = Num_Matric
  RatonNormal
End Function

Public Sub Encabezado_Materias(AdoMaterias As Adodc, _
                               CantCampos As Integer, _
                               Anchos As Single, _
                               alto As Single, _
                               Optional SegundaPagina As Boolean)
Dim InicX As Single
Dim InicY As Single
Dim PosLineaAux As Single
PosLinea = 0.01
Encabezado 0.01, AnchoPapel
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
Printer.FontBold = False
If SQLMsg1 <> "" Then
   Printer.FontSize = 10
   PrinterTexto CentrarTextoEncab(SQLMsg1, 0.01, AnchoPapel), PosLinea, SQLMsg1
   PosLinea = PosLinea + 0.45
End If
If SQLMsg2 <> "" Then
   Printer.FontSize = 9
   PrinterTexto 0.5, PosLinea, SQLMsg2
   PosLinea = PosLinea + 0.4
End If
If SQLMsg3 <> "" Then
   Printer.FontSize = 8
   PrinterTexto 0.5, PosLinea, SQLMsg3
   PosLinea = PosLinea + 0.4
End If
PosLinea = PosLinea + 0.05
PosLineaAux = PosLinea
Printer.FontSize = 9
Printer.FontBold = True
InicX = 5
TextoBusqueda = ""
With AdoMaterias.Recordset
    For I = 1 To CantCampos - 1
        Cadena = ""
        Contador = 1
        For J = 1 To Len(.fields(I).Name)
            Cadena = Cadena & MidStrg(.fields(I).Name, J, 1)
            If Printer.TextWidth(Cadena) > 2.5 Then
               Do While MidStrg(Cadena, J, 1) <> " " And J > 1
                  J = J - 1
               Loop
               Cadena = TrimStrg(MidStrg(Cadena, 1, J))
               TextoBusqueda = TextoBusqueda & Cadena & vbCrLf
               Cadena = ""
               Contador = Contador + 1
            End If
        Next J
        TextoBusqueda = TextoBusqueda & Cadena & vbCrLf
        If Contador = 1 Then TextoBusqueda = TextoBusqueda & "_" & vbCrLf
    Next I
End With
RutaBackup = RutaSistema & "\FORMATOS\PARALELO\P" & MidStrg(Codigo, 1, 1) & MidStrg(Codigo, 3, 2) & MidStrg(Codigo, 6, 2) & ".BMP"
PrinterPaint RutaBackup, 8, PosLinea, Anchos, alto
Printer.FontSize = 8
PrinterTexto 1, PosLinea + 0.1, "CURSO:"
PrinterTexto 3, PosLinea + 0.1, Codigo
Printer.FontSize = 16
PrinterTexto 1, PosLinea + 1.1, "A L U M N O S"
Printer.FontBold = False
PosLinea = PosLinea + alto
Printer.Line (Ancho(0), PosLineaAux)-(Ancho(1) + Anchos, PosLinea), Negro, B
Printer.Line (Ancho(0) + 0.01, PosLineaAux + 0.01)-(Ancho(1) + Anchos - 0.01, PosLinea - 0.01), Negro, B
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Sub EncabezadoDisc(Datas As Adodc, _
                          CantCampos As Integer, _
                          Anchos As Single, _
                          alto As Single, _
                          Optional SegundaPagina As Boolean)
Dim InicX As Single
Dim InicY As Single
Dim PosLineaAux As Single
PosLinea = 0.01
Encabezado 0.01, AnchoPapel
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
Printer.FontBold = False
If SQLMsg1 <> "" Then
   Printer.FontSize = 10
   PrinterTexto CentrarTextoEncab(SQLMsg1, 0.01, AnchoPapel), PosLinea, SQLMsg1
   PosLinea = PosLinea + 0.45
End If
If SQLMsg2 <> "" Then
   Printer.FontSize = 9
   PrinterTexto 0.5, PosLinea, SQLMsg2
   PosLinea = PosLinea + 0.4
End If
If SQLMsg3 <> "" Then
   Printer.FontSize = 8
   PrinterTexto 0.5, PosLinea, SQLMsg3
   PosLinea = PosLinea + 0.4
End If
PosLinea = PosLinea + 0.05
PosLineaAux = PosLinea
Printer.FontSize = 9
Printer.FontBold = True
InicX = 5
'Codigo = "3.03.CB"
RutaBackup = RutaSistema & "\FORMATOS\DISCIPLINA\D" & MidStrg(Codigo, 1, 1) & MidStrg(Codigo, 3, 2) & MidStrg(Codigo, 6, 2) & ".BMP"
'MsgBox RutaBackup
Anchos = Anchos - 0.3
PrinterPaint RutaBackup, 8, PosLinea, Anchos, alto
Printer.FontSize = 8
PrinterTexto 1, PosLinea + 0.1, "CURSO:"
PrinterTexto 3, PosLinea + 0.1, Codigo
Printer.FontSize = 16
PrinterTexto 1, PosLinea + 1.1, "A L U M N O S"
Printer.FontBold = False
PosLinea = PosLinea + alto
Printer.Line (Ancho(0), PosLineaAux)-(Ancho(1) + Anchos, PosLinea), Negro, B
Printer.Line (Ancho(0) + 0.01, PosLineaAux + 0.01)-(Ancho(1) + Anchos - 0.01, PosLinea - 0.01), Negro, B
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Sub Imprimir_Notas(Datas As Adodc, _
                          VectMat() As String, _
                         Optional EsCampoCorto As Boolean)
On Error GoTo Errorhandler
RatonReloj
Dim PointColor
Dim PictTexto, Texto, CampoTexto, LineaDeTexto As String
Dim AnchoDeLinea As Single
Dim IJ As Integer
Dim AnchoMaximo As Single
Dim AltoMaximo As Single
Dim SizeLetra As Integer
Dim VectProm(30) As Single
Dim ContVert(30) As Integer
SizeLetra = 8
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, 1, EsCampoCorto
For I = 1 To 30
    VectProm(I) = 0
    ContVert(I) = 0
Next I

PCol = 8
Ancho(0) = 0.5
Ancho(1) = PCol
Si_No = True
For I = 2 To CantCampos
    If Si_No Then
       PCol = PCol + 0.68
       Si_No = False
    Else
       PCol = PCol + 0.64
    End If
    Ancho(I) = PCol
Next I
Ancho(CantCampos) = Ancho(CantCampos) + 0.3
  Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
AltoMaximo = 2.47
AnchoMaximo = 12.82
Encabezado_Materias Datas, CantCampos, AnchoMaximo, AltoMaximo
PosLinea = 6.2
Printer.FontSize = SizeLetra
Contador = 0
With Datas.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     Do While Not .EOF
        'MsgBox CantCampos
        Contador = Contador + 1
        Printer.FontBold = False
        Printer.FontItalic = False
        Printer.FontUnderline = False
        PrinterTexto Ancho(0) + 0.05, PosLinea, Format$(Contador, "00") & ".-"
        PrinterTexto Ancho(0) + 0.55, PosLinea, .fields(0)
        For I = 1 To CantCampos - 1
            Printer.FontBold = False
            Printer.FontItalic = False
            Printer.FontUnderline = False
            If .fields(I) < 15 Then
                Printer.FontBold = True
                Printer.FontItalic = True
                Printer.FontUnderline = True
            End If
            If .fields(I) <> 0 Then
               'MsgBox .Fields(I).Name
                CodigoA = .fields(I).Name
                Si_No = False
                If MidStrg(CodigoA, Len(CodigoA), 1) = "_" Then Si_No = True
                If MidStrg(CodigoA, Len(CodigoA) - 1, 1) = "_" Then Si_No = True
                If Si_No Then
                   Select Case .fields(I)
                     Case 1 To 11: CodigoP = " I"
                     Case 12 To 13: CodigoP = " R"
                     Case 14 To 15: CodigoP = " B"
                     Case 16 To 18: CodigoP = "MB"
                     Case 19 To 20: CodigoP = " S"
                   End Select
                   PrinterTexto Ancho(I) + 0.05, PosLinea, CodigoP
                Else
                   If I = CantCampos - 1 Then
                      PrinterTexto Ancho(I) + 0.1, PosLinea, Format$(.fields(I), "00.00")
                   Else
                      PrinterTexto Ancho(I) + 0.1, PosLinea, Format$(.fields(I), "00")
                   End If
                   VectProm(I) = VectProm(I) + .fields(I)
                   ContVert(I) = ContVert(I) + 1
                End If
            End If
        Next I
        For I = 0 To CantCampos
            Printer.Line (Ancho(I), PosLinea - 0.01)-(Ancho(I), PosLinea + 0.35), Negro
        Next I
        PosLinea = PosLinea + 0.35
        Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
        PosLinea = PosLinea + 0.05
        If PosLinea >= LimiteAlto Then
           Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
           Printer.NewPage
           Encabezado_Materias Datas, CantCampos, AnchoMaximo, AltoMaximo
           PosLinea = 6.2
           Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
 End If
End With
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
Printer.FontItalic = False
Printer.FontUnderline = False
PosLinea = PosLinea + 0.1
PrinterTexto Ancho(0) + 4.5, PosLinea, "TOTAL PROMEDIOS:"
Printer.FontSize = SizeLetra - 2
For I = 1 To CantCampos - 1
    If ContVert(I) = 0 Then ContVert(I) = 1
    VectProm(I) = VectProm(I) / ContVert(I)
    If I = (CantCampos - 1) Then Printer.FontSize = SizeLetra
    If VectProm(I) <> 0 Then PrinterTexto Ancho(I) + 0.01, PosLinea, Format$(VectProm(I), "00.00")
Next I
PosLinea = PosLinea + 0.4
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 1.5
Printer.FontSize = SizeLetra
Printer.Line (2.5, PosLinea)-(6, PosLinea), Negro
Printer.Line (8.5, PosLinea)-(12, PosLinea), Negro
PosLinea = PosLinea + 0.05
PrinterTexto 3, PosLinea, CodigoBenef
PrinterTexto 9, PosLinea, CodigoCorresp
PosLinea = PosLinea + 0.35
If TipoDoc = "1" Then
   PrinterTexto 3, PosLinea, "DIRECTOR(A)"
   PrinterTexto 9, PosLinea, "SECRETARIO(A)"
End If
If TipoDoc = "2" Then
   PrinterTexto 3, PosLinea, "RECTOR(A)"
   PrinterTexto 9, PosLinea, "SECRETARIO(A)"
End If
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Imprimir_Disciplina(Datas As Adodc, _
                               VectMat() As String, _
                               Optional EsCampoCorto As Boolean)
On Error GoTo Errorhandler
RatonReloj
Dim PointColor
Dim PictTexto, Texto, CampoTexto, LineaDeTexto As String
Dim AnchoDeLinea As Single
Dim IJ As Integer
Dim AnchoMaximo As Single
Dim AltoMaximo As Single
Dim SizeLetra As Integer
Dim VectProm(30) As Single
Dim ContVert(30) As Integer
SizeLetra = 8
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0

DataAnchoCampos InicioX, Datas, SizeLetra, TipoArialNarrow, 1, EsCampoCorto
For I = 1 To 30
    VectProm(I) = 0
    ContVert(I) = 0
Next I

PCol = 8
Ancho(0) = 0.5
Ancho(1) = PCol
Si_No = True
For I = 2 To CantCampos
    PCol = PCol + 0.54
'''    If Si_No Then
'''       PCol = PCol + 0.54
'''       Si_No = False
'''    Else
'''       PCol = PCol + 0.55
'''    End If
    Ancho(I) = PCol
Next I
Ancho(CantCampos) = Ancho(CantCampos) + 0.3
  Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
AltoMaximo = 2.47
AnchoMaximo = 12.82
EncabezadoDisc Datas, CantCampos, AnchoMaximo, AltoMaximo
PosLinea = 6.2
Printer.FontSize = SizeLetra
Printer.FontName = TipoArialNarrow
Contador = 0
With Datas.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     Do While Not .EOF
        'MsgBox CantCampos
        Contador = Contador + 1
        Printer.FontBold = False
        Printer.FontItalic = False
        Printer.FontUnderline = False
        PrinterTexto Ancho(0) + 0.05, PosLinea, Format$(Contador, "00") & ".-"
        PrinterTexto Ancho(0) + 0.55, PosLinea, .fields(0)
        For I = 1 To CantCampos - 1
            Printer.FontBold = False
            Printer.FontItalic = False
            Printer.FontUnderline = False
            If .fields(I) <> 0 Then
                If I = (CantCampos - 1) Then
                   PrinterTexto Ancho(I) + 0.05, PosLinea, Format$(.fields(I), "00.00")
                Else
                   PrinterTexto Ancho(I) + 0.05, PosLinea, Format$(.fields(I), "00")
                End If
                VectProm(I) = VectProm(I) + .fields(I)
                ContVert(I) = ContVert(I) + 1
            End If
        Next I
        For I = 0 To CantCampos
            Printer.Line (Ancho(I), PosLinea - 0.01)-(Ancho(I), PosLinea + 0.35), Negro
        Next I
        PosLinea = PosLinea + 0.35
        Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
        PosLinea = PosLinea + 0.05
        If PosLinea >= LimiteAlto Then
           Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
           Printer.NewPage
           EncabezadoDisc Datas, CantCampos, AnchoMaximo, AltoMaximo
           PosLinea = 6.2
           Printer.FontSize = SizeLetra
           Printer.FontName = TipoArialNarrow
        End If
       .MoveNext
     Loop
 End If
End With
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.1
PrinterTexto Ancho(0) + 4.5, PosLinea, "TOTAL PROMEDIOS:"
Printer.FontSize = SizeLetra - 2
For I = 1 To CantCampos - 1
    If ContVert(I) = 0 Then ContVert(I) = 1
    VectProm(I) = VectProm(I) / ContVert(I)
    If I = (CantCampos - 1) Then Printer.FontSize = SizeLetra
    If VectProm(I) <> 0 Then PrinterTexto Ancho(I) + 0.01, PosLinea, Format$(VectProm(I), "00.00")
Next I
PosLinea = PosLinea + 0.4
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub GenerarVerNotas(Datas As Adodc, _
                           FinDoc As Boolean, _
                           FormaImp As Byte, _
                           SizeLetra As Integer, _
                           VectMat() As String, _
                           Optional EsCampoCorto As Boolean)
On Error GoTo Errorhandler
RatonReloj
Dim PointColor
Dim PictTexto, Texto, CampoTexto, LineaDeTexto As String
Dim AnchoDeLinea As Single
Dim IJ As Integer
Dim AnchoMaximo As Single
Dim AltoMaximo As Single

Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, 1, EsCampoCorto
'CantCampos = 20
'ReDim Ancho(CantCampos) As Single
PCol = 7
Ancho(0) = 0.5
Ancho(1) = PCol
Si_No = True
For I = 2 To CantCampos
    If Si_No Then
       PCol = PCol + 0.68
       Si_No = False
    Else
       PCol = PCol + 0.64
    End If
    Ancho(I) = PCol
Next I
''  AltoMaximo = 3.53
''  AnchoMaximo = 18.31
  Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
    .MoveFirst
     AltoMaximo = 2.47
     AnchoMaximo = 12.82
     Encabezado_Materias Datas, CantCampos, AnchoMaximo, AltoMaximo
     PosLinea = 4.1
     
     PosLinea = PosLinea + 0.05
     Printer.FontSize = SizeLetra
     Do While Not .EOF
        'MsgBox CantCampos
        PrinterTexto Ancho(0) + 0.2, PosLinea, .fields(0)
        For I = 0 To CantCampos - 1
            PrinterTexto Ancho(I) + 0.1, PosLinea, Format$(.fields(I), "00")
        Next I
        For I = 0 To CantCampos
            Printer.Line (Ancho(I), PosLinea - 0.01)-(Ancho(I), PosLinea + 0.35), Negro
        Next I
        PosLinea = PosLinea + 0.35
        Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
        PosLinea = PosLinea + 0.05
        If PosLinea >= LimiteAlto Then
           Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
           Printer.NewPage
           Encabezado_Materias Datas, CantCampos, AnchoMaximo, AltoMaximo
           Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
End With
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.05
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
RatonNormal
MensajeEncabData = ""
If FinDoc Then Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Actualizar_Notas_del_Curso(CodMat As String, Curso As String)
Dim Strgs As String
Dim Cod_Prof As String
Dim CodMatP As String
Dim AdoRegs As ADODB.Recordset
   
   RatonReloj
    Set AdoRegs = New ADODB.Recordset
    AdoRegs.CursorType = adOpenStatic
    AdoRegs.CursorLocation = adUseClient
    CodMatP = Ninguno
    Strgs = "SELECT * " _
          & "FROM Catalogo_Estudiantil " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND MidStrg(CodigoE,1," & Len(Curso) & ") = '" & Curso & "' " _
          & "AND CodMat = '" & CodMat & "' "
    Strgs = CompilarSQL(Strgs)
    AdoRegs.open Strgs, AdoStrCnn, , , adCmdText
    With AdoRegs
     If .RecordCount > 0 Then
         Cod_Prof = AdoRegs.fields("Profesor")
         CodMatP = AdoRegs.fields("CodMatP")
     End If
    End With
    AdoRegs.Close
    
   RatonReloj
  'MsgBox CodMat & vbCrLf & Curso
   If FormatoLibreta = "QUIMESTRE" Then
      Select Case CodMat
        Case "998", "999"
             If SQL_Server Then
                sSQL = "UPDATE Trans_Asistencia " _
                     & "SET " & SQLConductaQ & " = AN." & SQLConductaQ & ", " _
                     & SQLDias & " = AN." & SQLDias & ", " _
                     & SQLFJ & " = AN." & SQLFJ & ", " _
                     & SQLFI & " = AN." & SQLFI & ", " _
                     & SQLAtrasos & " = AN." & SQLAtrasos & " " _
                     & "FROM Trans_Asistencia As TN,Asiento_AS As AN "
             Else
             
             End If
        Case Else
             If SQL_Server Then
                If CodMatP = Ninguno Then
                   sSQL = "UPDATE Trans_Notas "
                Else
                   sSQL = "UPDATE Trans_Notas_Auxiliares "
                End If
                If OpcionNotas = 4 Then
                   sSQL = sSQL & "SET " & SQLExamen & " = AN.Examen_Q "
                ElseIf OpcionNotas = 5 Then
                   sSQL = sSQL _
                        & "SET Nota_Grado = AN.Nota_Grado," _
                        & "Supletorio = AN.Supletorio," _
                        & "Remedial = AN.Remedial "
                Else
                   sSQL = sSQL _
                        & "SET " & SQLTAI & " = AN.TAI," _
                        & SQLAIC & " = AN.AIC," _
                        & SQLAGC & " = AN.AGC," _
                        & SQLL & " = AN.LECCIONES," _
                        & SQLExaP & " = AN.EXAMEN "
                End If
                If CodMatP = Ninguno Then
                   sSQL = sSQL & "FROM Trans_Notas As TN, Asiento_N As AN "
                Else
                   sSQL = sSQL & "FROM Trans_Notas_Auxiliares As TN, Asiento_N As AN "
                End If
             Else
             End If
      End Select
      sSQL = sSQL & "WHERE TN.Item = '" & NumEmpresa & "' " _
           & "AND TN.Periodo = '" & Periodo_Contable & "' " _
           & "AND TN.CodE = '" & Curso & "' " _
           & "AND TN.CodMat = '" & CodMat & "' " _
           & "AND TN.CodMat = AN.CodMat " _
           & "AND TN.Item = AN.Item " _
           & "AND TN.Codigo = AN.Codigo "
   Else
        Select Case CodMat
          Case "998", "999"
               If SQL_Server Then
                  sSQL = "UPDATE Trans_Asistencia " _
                       & "SET PQBFJ1 = AN.PQBFJ1," _
                       & "PQBFI1 = AN.PQBFI1," _
                       & "PQBA1 = AN.PQBA1," _
                       & "ConductaPQ1 = AN.ConductaPQ1," _
                       & "PQBFJ2 = AN.PQBFJ2," _
                       & "PQBFI2 = AN.PQBFI2," _
                       & "PQBA2 = AN.PQBA2," _
                       & "ConductaPQ2 = AN.ConductaPQ2,"
                  sSQL = sSQL _
                       & "SQBFJ1 = AN.SQBFJ1," _
                       & "SQBFI1 = AN.SQBFI1," _
                       & "SQBA1 = AN.SQBA1," _
                       & "ConductaSQ1 = AN.ConductaSQ1," _
                       & "SQBFJ2 = AN.SQBFJ2," _
                       & "SQBFI2 = AN.SQBFI2," _
                       & "SQBA2 = AN.SQBA2," _
                       & "ConductaSQ2 = AN.ConductaSQ2,"
                  sSQL = sSQL _
                       & "TQBFJ1 = AN.TQBFJ1," _
                       & "TQBFI1 = AN.TQBFI1," _
                       & "TQBA1 = AN.TQBA1," _
                       & "ConductaTQ1 = AN.ConductaTQ1," _
                       & "TQBFJ2 = AN.TQBFJ2," _
                       & "TQBFI2 = AN.TQBFI2," _
                       & "TQBA2 = AN.TQBA2," _
                       & "ConductaTQ2 = AN.ConductaTQ2 " _
                       & "FROM Trans_Asistencia As TN,Asiento_AS As AN "
               Else
                  sSQL = "UPDATE Trans_Asistencia As TN,Asiento_AS As AN " _
                       & "SET TN.PQBFJ1 = AN.PQBFJ1," _
                       & "TN.PQBFI1 = AN.PQBFI1," _
                       & "TN.PQBA1 = AN.PQBA1," _
                       & "TN.ConductaPQ1 = AN.ConductaPQ1," _
                       & "TN.PQBFJ2 = AN.PQBFJ2," _
                       & "TN.PQBFI2 = AN.PQBFI2," _
                       & "TN.PQBA2 = AN.PQBA2," _
                       & "TN.ConductaPQ2 = AN.ConductaPQ2,"
                  sSQL = sSQL _
                       & "TN.SQBFJ1 = AN.SQBFJ1," _
                       & "TN.SQBFI1 = AN.SQBFI1," _
                       & "TN.SQBA1 = AN.SQBA1," _
                       & "TN.ConductaSQ1 = AN.ConductaSQ1," _
                       & "TN.SQBFJ2 = AN.SQBFJ2," _
                       & "TN.SQBFI2 = AN.SQBFI2," _
                       & "TN.SQBA2 = AN.SQBA2," _
                       & "TN.ConductaSQ2 = AN.ConductaSQ2,"
                  sSQL = sSQL _
                       & "TN.TQBFJ1 = AN.TQBFJ1," _
                       & "TN.TQBFI1 = AN.TQBFI1," _
                       & "TN.TQBA1 = AN.TQBA1," _
                       & "TN.ConductaTQ1 = AN.ConductaTQ1," _
                       & "TN.TQBFJ2 = AN.TQBFJ2," _
                       & "TN.TQBFI2 = AN.TQBFI2," _
                       & "TN.TQBA2 = AN.TQBA2," _
                       & "TN.ConductaTQ2 = AN.ConductaTQ2 "
               End If
               sSQL = sSQL & "WHERE TN.Item = '" & NumEmpresa & "' " _
                    & "AND TN.Periodo = '" & Periodo_Contable & "' " _
                    & "AND TN.CodE = '" & Curso & "' " _
                    & "AND TN.CodMat = '" & CodMat & "' " _
                    & "AND AN.CodigoU = '" & CodigoUsuario & "' " _
                    & "AND TN.CodMat = AN.CodMat " _
                    & "AND TN.Item = AN.Item " _
                    & "AND TN.Codigo = AN.Codigo "
          Case Else
               If SQL_Server Then
                  If CodMatP = Ninguno Then
                     sSQL = "UPDATE Trans_Notas "
                  Else
                     sSQL = "UPDATE Trans_Notas_Auxiliares "
                  End If
                  sSQL = sSQL _
                       & "SET PQBim1 = AN.PQBim1," _
                       & "PQBim2 = AN.PQBim2," _
                       & "SQBim1 = AN.SQBim1," _
                       & "SQBim2 =  AN.SQBim2," _
                       & "TQBim1 = AN.TQBim1," _
                       & "TQBim2 =  AN.TQBim2," _
                       & "ExamenPQ = AN.ExamenPQ," _
                       & "ExamenSQ = AN.ExamenSQ," _
                       & "ExamenTQ = AN.ExamenTQ," _
                       & "ConductaPQ1 = AN.ConductaPQ1," _
                       & "ConductaPQ2 = AN.ConductaPQ2," _
                       & "ConductaSQ1 = AN.ConductaSQ1," _
                       & "ConductaSQ2 = AN.ConductaSQ2," _
                       & "ConductaTQ1 = AN.ConductaTQ1," _
                       & "ConductaTQ2 = AN.ConductaTQ2," _
                       & "Nota_Grado = AN.Nota_Grado," _
                       & "Supletorio  = AN.Supletorio "
                  If CodMatP = Ninguno Then
                     sSQL = sSQL & "FROM Trans_Notas As TN, Asiento_N As AN "
                  Else
                     sSQL = sSQL & "FROM Trans_Notas_Auxiliares As TN, Asiento_N As AN "
                  End If
               Else
                  If CodMatP = Ninguno Then
                     sSQL = "UPDATE Trans_Notas As TN, Asiento_N As AN "
                  Else
                     sSQL = "UPDATE Trans_Notas_Auxiliares As TN, Asiento_N As AN "
                  End If
                  sSQL = sSQL _
                       & "SET TN.PQBim1 = AN.PQBim1," _
                       & "TN.PQBim2 = AN.PQBim2," _
                       & "TN.SQBim1 = AN.SQBim1," _
                       & "TN.SQBim2 =  AN.SQBim2," _
                       & "TN.TQBim1 = AN.TQBim1," _
                       & "TN.TQBim2 =  AN.TQBim2," _
                       & "TN.ExamenPQ = AN.ExamenPQ," _
                       & "TN.ExamenSQ = AN.ExamenSQ," _
                       & "TN.ExamenTQ = AN.ExamenTQ," _
                       & "TN.ConductaPQ1 = AN.ConductaPQ1," _
                       & "TN.ConductaPQ2 = AN.ConductaPQ2," _
                       & "TN.ConductaSQ1 = AN.ConductaSQ1," _
                       & "TN.ConductaSQ2 = AN.ConductaSQ2," _
                       & "TN.ConductaTQ1 = AN.ConductaTQ1," _
                       & "TN.ConductaTQ2 = AN.ConductaTQ2," _
                       & "TN.Nota_Grado = AN.Nota_Grado," _
                       & "TN.Supletorio  = AN.Supletorio "
               End If
               sSQL = sSQL & "WHERE TN.Item = '" & NumEmpresa & "' " _
                    & "AND TN.Periodo = '" & Periodo_Contable & "' " _
                    & "AND TN.CodE = '" & Curso & "' " _
                    & "AND TN.CodMat = '" & CodMat & "' " _
                    & "AND AN.CodigoU = '" & CodigoUsuario & "' " _
                    & "AND TN.CodMat = AN.CodMat " _
                    & "AND TN.Item = AN.Item " _
                    & "AND TN.Codigo = AN.Codigo "
        End Select
   End If
   Ejecutar_SQL_SP sSQL
   
   sSQL = "DELETE * " _
        & "FROM Asiento_N " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "AND CodMat = '" & CodMat & "' "
   Ejecutar_SQL_SP sSQL
  
   sSQL = "DELETE * " _
        & "FROM Asiento_AS " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "AND CodMat = '" & CodMat & "' "
   Ejecutar_SQL_SP sSQL
   RatonNormal
End Sub

Public Sub Imprimir_Nomina_Notas(Datas As Adodc, _
                                 AdoAut As Adodc, _
                                 Optional EsCampoCorto As Boolean, _
                                 Optional EsActaGrado As Boolean)
Dim SizeLetra As Integer
Dim AnchoCol As Single

On Error GoTo Errorhandler

Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, Orientacion_Pagina, EsCampoCorto
Pagina = 1
For I = 2 To CantCampos - 1
    Ancho(I) = Ancho(I) + 2
Next I
Ancho(CantCampos) = LimiteAncho
'Iniciamos la impresion
'0925751992
Printer.FontBold = False
With Datas.Recordset
If .RecordCount > 0 Then
    .MoveFirst
    AnchoCol = 1
    Encabezado 1.5, 19
    Printer.FontName = TipoTimes
    Printer.FontBold = False
    If SQLMsg1 <> "" Then
       Printer.FontSize = 12
       PrinterTexto 1.5, PosLinea, SQLMsg1
       PosLinea = PosLinea + 0.45
    End If
    If SQLMsg2 <> "" Then
       Printer.FontSize = 10
       PrinterTexto 1.5, PosLinea, SQLMsg2
       PosLinea = PosLinea + 0.4
    End If
    If SQLMsg3 <> "" Then
       Printer.FontSize = 8
       PrinterTexto 1.5, PosLinea, SQLMsg3
       PosLinea = PosLinea + 0.4
    End If
    PosLinea = PosLinea + 0.05
    Printer.FontSize = 8
    Printer.FontBold = True
    Printer.FontItalic = False
        If AdoAut.Recordset.RecordCount > 0 Then
           AdoAut.Recordset.MoveFirst
           PrinterTexto 2.5, PosLinea, "A P E L L I D O S   Y   N O M B R E S"
           PosColumna = 10.5
           If EsActaGrado Then
              PrinterTexto PosColumna, PosLinea, "Examen Grado"
           Else
               Select Case .fields("CodMat")
                 Case "998", "999"
                      If AdoAut.Recordset.fields("NPQP1") Then
                         PrinterTexto PosColumna, PosLinea, "ConductaPQ1"
                         PosColumna = PosColumna + AnchoCol
                         PrinterTexto PosColumna, PosLinea, "PQBFJ1"
                         PosColumna = PosColumna + AnchoCol
                         PrinterTexto PosColumna, PosLinea, "PQBFI1"
                         PosColumna = PosColumna + AnchoCol
                         PrinterTexto PosColumna, PosLinea, "PQBA1"
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NPQP2") Then
                         PrinterTexto PosColumna, PosLinea, "ConductaPQ2"
                         PosColumna = PosColumna + AnchoCol
                         PrinterTexto PosColumna, PosLinea, "PQBFJ2"
                         PosColumna = PosColumna + AnchoCol
                         PrinterTexto PosColumna, PosLinea, "PQBFI2"
                         PosColumna = PosColumna + AnchoCol
                         PrinterTexto PosColumna, PosLinea, "PQBA2"
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NSQP1") Then
                         PrinterTexto PosColumna, PosLinea, "ConductaSQ1"
                         PosColumna = PosColumna + AnchoCol
                         PrinterTexto PosColumna, PosLinea, "SQBFJ1"
                         PosColumna = PosColumna + AnchoCol
                         PrinterTexto PosColumna, PosLinea, "SQBFI1"
                         PosColumna = PosColumna + AnchoCol
                         PrinterTexto PosColumna, PosLinea, "SQBA1"
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NSQP2") Then
                         PrinterTexto PosColumna, PosLinea, "ConductaSQ2"
                         PosColumna = PosColumna + AnchoCol
                         PrinterTexto PosColumna, PosLinea, "SQBFJ2"
                         PosColumna = PosColumna + AnchoCol
                         PrinterTexto PosColumna, PosLinea, "SQBFI2"
                         PosColumna = PosColumna + AnchoCol
                         PrinterTexto PosColumna, PosLinea, "SQBA2"
                         PosColumna = PosColumna + AnchoCol
                      End If
                 Case Else
                      If AdoAut.Recordset.fields("NPQP1") Then
                         PrinterTexto PosColumna, PosLinea, "PQBim1"
                         PosColumna = PosColumna + AnchoCol
                         PrinterTexto PosColumna, PosLinea, "ConductaPQ1"
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NPQP2") Then
                         PrinterTexto PosColumna, PosLinea, "PQBim2"
                         PosColumna = PosColumna + AnchoCol
                         PrinterTexto PosColumna, PosLinea, "ConductaPQ2"
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NPQEX") Then
                         PrinterTexto PosColumna, PosLinea, "ExamenPQ"
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NSQP1") Then
                         PrinterTexto PosColumna, PosLinea, "SQBim1"
                         PosColumna = PosColumna + AnchoCol
                         PrinterTexto PosColumna, PosLinea, "ConductaSQ1"
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NSQP2") Then
                         PrinterTexto PosColumna, PosLinea, "SQBim2"
                         PosColumna = PosColumna + AnchoCol
                         PrinterTexto PosColumna, PosLinea, "ConductaSQ2"
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NSQEX") Then
                         PrinterTexto PosColumna, PosLinea, "ExamenSQ"
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NSUPL") Then
                         PrinterTexto PosColumna, PosLinea, "Supletorio"
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NGRADO") Then
                         PrinterTexto PosColumna, PosLinea, "Nota_Grado"
                         PosColumna = PosColumna + AnchoCol
                      End If
               End Select
           End If
        End If
    Printer.FontBold = False
    PosLinea = PosLinea + 0.4
     Printer.FontSize = 9
     Imprimir_Linea_H PosLinea, 1.5, 19, Negro, True
     PosLinea = PosLinea + 0.05
     Contador = 1
     Printer.FontName = TipoCourierNew
     Do While Not .EOF
        PrinterTexto 1.4, PosLinea, Format$(Contador, "00") & ".-"
        PrinterFields 2.3, PosLinea, .fields("Alumno")
        PosColumna = 10.5
        If AdoAut.Recordset.RecordCount > 0 Then
           AdoAut.Recordset.MoveFirst
           'AdoAut.Recordset.Fields ("NGRADO")
           If EsActaGrado Then
              PrinterFields PosColumna, PosLinea, .fields("Examen"), True
           Else
               Select Case .fields("CodMat")
                 Case "998", "999"
                      If AdoAut.Recordset.fields("NPQP1") Then
                         PrinterFields PosColumna, PosLinea, .fields("ConductaPQ1"), True
                         PosColumna = PosColumna + AnchoCol
                         PrinterFields PosColumna, PosLinea, .fields("PQBFJ1"), True
                         PosColumna = PosColumna + AnchoCol
                         PrinterFields PosColumna, PosLinea, .fields("PQBFI1"), True
                         PosColumna = PosColumna + AnchoCol
                         PrinterFields PosColumna, PosLinea, .fields("PQBA1"), True
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NPQP2") Then
                         PrinterFields PosColumna, PosLinea, .fields("ConductaPQ2"), True
                         PosColumna = PosColumna + AnchoCol
                         PrinterFields PosColumna, PosLinea, .fields("PQBFJ2"), True
                         PosColumna = PosColumna + AnchoCol
                         PrinterFields PosColumna, PosLinea, .fields("PQBFI2"), True
                         PosColumna = PosColumna + AnchoCol
                         PrinterFields PosColumna, PosLinea, .fields("PQBA2"), True
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NSQP1") Then
                         PrinterFields PosColumna, PosLinea, .fields("ConductaSQ1"), True
                         PosColumna = PosColumna + AnchoCol
                         PrinterFields PosColumna, PosLinea, .fields("SQBFJ1"), True
                         PosColumna = PosColumna + AnchoCol
                         PrinterFields PosColumna, PosLinea, .fields("SQBFI1"), True
                         PosColumna = PosColumna + AnchoCol
                         PrinterFields PosColumna, PosLinea, .fields("SQBA1"), True
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NSQP2") Then
                         PrinterFields PosColumna, PosLinea, .fields("ConductaSQ2"), True
                         PosColumna = PosColumna + AnchoCol
                         PrinterFields PosColumna, PosLinea, .fields("SQBFJ2"), True
                         PosColumna = PosColumna + AnchoCol
                         PrinterFields PosColumna, PosLinea, .fields("SQBFI2"), True
                         PosColumna = PosColumna + AnchoCol
                         PrinterFields PosColumna, PosLinea, .fields("SQBA2"), True
                         PosColumna = PosColumna + AnchoCol
                      End If
                 Case Else
                      If AdoAut.Recordset.fields("NPQP1") Then
                         PrinterFields PosColumna, PosLinea, .fields("PQBim1"), True
                         PosColumna = PosColumna + AnchoCol
                         PrinterFields PosColumna, PosLinea, .fields("ConductaPQ1"), True
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NPQP2") Then
                         PrinterFields PosColumna, PosLinea, .fields("PQBim2"), True
                         PosColumna = PosColumna + AnchoCol
                         PrinterFields PosColumna, PosLinea, .fields("ConductaPQ2"), True
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NPQEX") Then
                         PrinterFields PosColumna, PosLinea, .fields("ExamenPQ"), True
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NSQP1") Then
                         PrinterFields PosColumna, PosLinea, .fields("SQBim1"), True
                         PosColumna = PosColumna + AnchoCol
                         PrinterFields PosColumna, PosLinea, .fields("ConductaSQ1"), True
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NSQP2") Then
                         PrinterFields PosColumna, PosLinea, .fields("SQBim2"), True
                         PosColumna = PosColumna + AnchoCol
                         PrinterFields PosColumna, PosLinea, .fields("ConductaSQ2"), True
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NSQEX") Then
                         PrinterFields PosColumna, PosLinea, .fields("ExamenSQ"), True
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NSUPL") Then
                         PrinterFields PosColumna, PosLinea, .fields("Supletorio"), True
                         PosColumna = PosColumna + AnchoCol
                      End If
                      If AdoAut.Recordset.fields("NGRADO") Then
                         PrinterFields PosColumna, PosLinea, .fields("Nota_Grado"), True
                         PosColumna = PosColumna + AnchoCol
                      End If
               End Select
           End If
        End If
        PosLinea = PosLinea + 0.36
        Contador = Contador + 1
       .MoveNext
     Loop
End If
End With
Imprimir_Linea_H PosLinea, 1.5, 19, Negro, True
PosLinea = PosLinea + 1.5
PrinterTexto 2, PosLinea, String$(25, "_")
PosLinea = PosLinea + 0.4
PrinterTexto 2.6, PosLinea, "FIRMA RESPONSABLE"
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Imprimir_Promedio_Notas(Datas As Adodc, _
                                   SizeLetra As Integer, _
                                   Optional EsCampoCorto As Boolean, _
                                   Optional PromTotal As Boolean)
On Error GoTo Errorhandler

Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, Orientacion_Pagina, True
Pagina = 1
If PromTotal Then
   Ancho(1) = Ancho(1) - 1
   Ancho(2) = Ancho(2) - 0.5
   Distancia = Ancho(2)
   For I = 3 To CantCampos - 1
       If Distancia <= AnchoPapel Then
          Distancia = Distancia + 1.6
          Ancho(I) = Distancia
       End If
   Next I
Else
   Ancho(2) = Ancho(2) + 0.5
   Distancia = Ancho(2)
   For I = 3 To CantCampos - 1
       If Distancia <= AnchoPapel Then
          Distancia = Distancia + 1.7
          Ancho(I) = Distancia
       End If
   Next I
   Ancho(CantCampos - 1) = 0
   Ancho(CantCampos - 2) = 0
End If
Ancho(CantCampos) = LimiteAncho
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
If .RecordCount > 0 Then
    .MoveFirst
     Encabezado_Institucion Ancho(0), Ancho(CantCampos)
     PrinterAllFields CantCampos, PosLinea, Datas, True, True
     PosLinea = PosLinea + 0.4
     Printer.FontSize = SizeLetra
     Do While Not .EOF
        PrinterAllFields CantCampos, PosLinea, Datas, True, False
        PosLinea = PosLinea + 0.36
        If Cuadricula Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Gris
           PosLinea = PosLinea + 0.05
        End If
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
           Printer.NewPage
           Encabezado_Institucion Ancho(0), Ancho(CantCampos)
           PrinterAllFields CantCampos, PosLinea, Datas, True, True
           PosLinea = PosLinea + 0.4
           Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
End If
End With
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos), Negro, True
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Function Tipo_Acceso_Educativo(Tipo_Tabla As String, NombreCampo As String) As String
Dim CadAcceso As String
  CadAcceso = ""
  If Opc_Primaria = False And Opc_Secundaria = False And Opc_Bachillerato = False Then
     CadAcceso = " "
  Else
     If Opc_Primaria = True And Opc_Secundaria = False And Opc_Bachillerato = False Then
        CadAcceso = "AND MidStrg(" & Tipo_Tabla & NombreCampo & ",1,1) IN ('1') "
     End If
     If Opc_Primaria = False And Opc_Secundaria = True And Opc_Bachillerato = False Then
        CadAcceso = "AND MidStrg(" & Tipo_Tabla & NombreCampo & ",1,1) IN ('2') "
     End If
     If Opc_Primaria = False And Opc_Secundaria = False And Opc_Bachillerato = True Then
        CadAcceso = "AND MidStrg(" & Tipo_Tabla & NombreCampo & ",1,1) IN ('3') "
     End If
     If Opc_Primaria = True And Opc_Secundaria = True And Opc_Bachillerato = False Then
        CadAcceso = "AND MidStrg(" & Tipo_Tabla & NombreCampo & ",1,1) IN ('1','2') "
     End If
     If Opc_Primaria = True And Opc_Secundaria = False And Opc_Bachillerato = True Then
        CadAcceso = "AND MidStrg(" & Tipo_Tabla & NombreCampo & ",1,1) IN ('1','3') "
     End If
     If Opc_Primaria = False And Opc_Secundaria = True And Opc_Bachillerato = True Then
        CadAcceso = "AND MidStrg(" & Tipo_Tabla & NombreCampo & ",1,1) IN ('2','3') "
     End If
     If Opc_Primaria = True And Opc_Secundaria = True And Opc_Bachillerato = True Then
        CadAcceso = "AND MidStrg(" & Tipo_Tabla & NombreCampo & ",1,1) IN ('1','2','3') "
     End If
     If CodigoUsuario = "0702164179" Then
        CadAcceso = "AND MidStrg(" & Tipo_Tabla & NombreCampo & ",1,1) IN ('1','2','3','4','5') "
     End If
  End If
 'MsgBox CodigoUsuario & vbCrLf & CadAcceso
  Tipo_Acceso_Educativo = CadAcceso
End Function

Public Sub Imprimir_Nomina_Alumnos(Datas As Adodc, _
                                   FechaInicial As String, _
                                   FechaFinal As String, _
                                   Optional EsCampoCorto As Boolean, _
                                   Optional Sin_Salto_Pagina As Boolean, _
                                   Optional TipoNomina As Byte)
On Error GoTo Errorhandler
Dim SizeLetra As Integer
SizeLetra = 8
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
J = Redondear((CFechaLong(FechaFinal) - CFechaLong(FechaInicial)) / 30) + 1
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, Orientacion_Pagina, EsCampoCorto
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
'MsgBox Sin_Salto_Pagina
If TipoNomina = 2 Then KE = 17
If TipoNomina = 3 Then KE = 5
With Datas.Recordset
If .RecordCount > 0 Then
    .MoveFirst
     Contador = 0
     Codigo = .fields("Grupo")
     Codigo3 = .fields("Materia")
     Codigo1 = ""
     Codigo2 = ""
     Cadena = ""
     Select Case MidStrg(.fields("Grupo"), 1, 1)
       Case "1": Codigo1 = UCaseStrg(.fields("Curso"))
       Case "2": Codigo1 = UCaseStrg(.fields("Curso"))
       Case "3":
                'MsgBox MidStrg(.Fields("Grupo_No"), 3, 2)
                Select Case MidStrg(.fields("Grupo"), 3, 2)
                  Case "01": Codigo1 = "PRIMER CURSO CICLO DIVERSIFICADO DEL " & UCaseStrg(.fields("Bachiller"))
                  Case "02": Codigo1 = "SEGUNDO CURSO CICLO DIVERSIFICADO DEL " & UCaseStrg(.fields("Bachiller"))
                  Case "03": Codigo1 = "TERCER CURSO CICLO DIVERSIFICADO DEL " & UCaseStrg(.fields("Bachiller"))
                End Select
     End Select
     If .fields("Especialidad") <> Ninguno Then Codigo2 = "ESPECIALIZACIÓN " & UCaseStrg(.fields("Especialidad"))
     'MensajeEncabData = Cadena
     Codigo3 = .fields("Materia")
     Codigo4 = .fields("Profesor")
     Encabezado_Lista_Alumnos SizeLetra, FechaInicial, Codigo, Codigo1, Codigo2, TipoNomina
     Do While Not .EOF
        Printer.FontItalic = False
        Printer.FontSize = SizeLetra
        Printer.FontName = TipoArialNarrow
        If Codigo3 <> .fields("Materia") Then
          'PosLinea = PosLinea - 0.1
           Imprimir_Linea_H PosLinea, 2, 19.5
           If TipoNomina = 2 Or TipoNomina = 3 Then
              InicioY = InicioY + 0.05
              Imprimir_Linea_V 2, InicioY, PosLinea
              Imprimir_Linea_V 2.55, InicioY, PosLinea
              Imprimir_Linea_V 19.5, InicioY, PosLinea
              PosColumna = 10.5
              Imprimir_Linea_V PosColumna, InicioY, PosLinea
              For I = 1 To KE
                  Imprimir_Linea_V PosColumna, InicioY, PosLinea
                  If TipoNomina = 3 Then
                     PosColumna = PosColumna + 1.5
                  Else
                     PosColumna = PosColumna + 0.6
                  End If
              Next I
           End If
           Printer.NewPage
           Codigo = .fields("Grupo")
           Codigo1 = ""
           Codigo2 = ""
           Cadena = ""
           Select Case MidStrg(.fields("Grupo"), 1, 1)
             Case "1": Codigo1 = UCaseStrg(.fields("Curso"))
             Case "2": Codigo1 = UCaseStrg(.fields("Curso"))
             Case "3":
                     'MsgBox MidStrg(.Fields("Grupo_No"), 3, 2)
                      Select Case MidStrg(.fields("Grupo"), 3, 2)
                        Case "01": Codigo1 = "PRIMER CURSO CICLO DIVERSIFICADO DEL " & UCaseStrg(.fields("Bachiller"))
                        Case "02": Codigo1 = "SEGUNDO CURSO CICLO DIVERSIFICADO DEL " & UCaseStrg(.fields("Bachiller"))
                        Case "03": Codigo1 = "TERCER CURSO CICLO DIVERSIFICADO DEL " & UCaseStrg(.fields("Bachiller"))
                      End Select
           End Select
           If .fields("Especialidad") <> Ninguno Then Codigo2 = "ESPECIALIZACIÓN " & UCaseStrg(.fields("Especialidad"))
           Codigo3 = .fields("Materia")
           Codigo4 = .fields("Profesor")
           Encabezado_Lista_Alumnos SizeLetra, FechaInicial, Codigo, Codigo1, Codigo2, TipoNomina
           Contador = 0
        End If
        
        If Codigo <> .fields("Grupo") Then
           Imprimir_Linea_H PosLinea, 2, 19.5
           If TipoNomina = 2 Or TipoNomina = 3 Then
              InicioY = InicioY + 0.05
              Imprimir_Linea_V 2, InicioY, PosLinea
              Imprimir_Linea_V 2.55, InicioY, PosLinea
              Imprimir_Linea_V 19.5, InicioY, PosLinea
              PosColumna = 10.5
              Imprimir_Linea_V PosColumna, InicioY, PosLinea
              For I = 1 To KE
                  Imprimir_Linea_V PosColumna, InicioY, PosLinea
                  If TipoNomina = 3 Then
                     PosColumna = PosColumna + 1.5
                  Else
                     PosColumna = PosColumna + 0.6
                  End If
              Next I
           End If
           Printer.NewPage
           Contador = 0
           Codigo = .fields("Grupo")
           Codigo1 = ""
           Codigo2 = ""
           Cadena = ""
           Select Case MidStrg(.fields("Grupo"), 1, 1)
             Case "1": Codigo1 = UCaseStrg(.fields("Curso"))
             Case "2": Codigo1 = UCaseStrg(.fields("Curso"))
             Case "3":
                      Select Case MidStrg(.fields("Grupo"), 3, 2)
                        Case "01": Codigo1 = "PRIMER CURSO CICLO DIVERSIFICADO DEL " & UCaseStrg(.fields("Bachiller"))
                        Case "02": Codigo1 = "SEGUNDO CURSO CICLO DIVERSIFICADO DEL " & UCaseStrg(.fields("Bachiller"))
                        Case "03": Codigo1 = "TERCER CURSO CICLO DIVERSIFICADO DEL " & UCaseStrg(.fields("Bachiller"))
                      End Select
           End Select
           If .fields("Especialidad") <> Ninguno Then Codigo2 = "ESPECIALIZACIÓN " & UCaseStrg(.fields("Especialidad"))
           'MensajeEncabData = Cadena
           Encabezado_Lista_Alumnos SizeLetra, FechaInicial, Codigo, Codigo1, Codigo2, TipoNomina
        End If
        'PosLinea = PosLinea + 0.05
        Contador = Contador + 1
        PrinterTexto 2, PosLinea, Format$(Contador, "00")
        PrinterTexto 2.6, PosLinea, .fields("Cliente")
        PosLinea = PosLinea + 0.4
        If TipoNomina = 2 Or TipoNomina = 3 Then Imprimir_Linea_H PosLinea, 2, 19.5
        PosLinea = PosLinea + 0.1
        K = .fields("No_") + 1
        If PosLinea >= LimiteAlto Then
           'PosLinea = PosLinea - 0.1
           Imprimir_Linea_H PosLinea, 2, 19.5
           If TipoNomina = 2 Or TipoNomina = 3 Then
              InicioY = InicioY + 0.05
              Imprimir_Linea_V 2, InicioY, PosLinea
              Imprimir_Linea_V 2.55, InicioY, PosLinea
              Imprimir_Linea_V 19.5, InicioY, PosLinea
              PosColumna = 10.5
              Imprimir_Linea_V PosColumna, InicioY, PosLinea
              For I = 1 To KE
                  Imprimir_Linea_V PosColumna, InicioY, PosLinea
                  If TipoNomina = 3 Then
                     PosColumna = PosColumna + 1.5
                  Else
                     PosColumna = PosColumna + 0.6
                  End If
              Next I
           End If
           Printer.NewPage
           Codigo = .fields("Grupo")
           Codigo1 = ""
           Codigo2 = ""
           Cadena = ""
           Select Case MidStrg(.fields("Grupo"), 1, 1)
             Case "1": Codigo1 = UCaseStrg(.fields("Curso"))
             Case "2": Codigo1 = UCaseStrg(.fields("Curso"))
             Case "3":
                     'MsgBox MidStrg(.Fields("Grupo_No"), 3, 2)
                      Select Case MidStrg(.fields("Grupo"), 3, 2)
                        Case "01": Codigo1 = "PRIMER CURSO CICLO DIVERSIFICADO DEL " & UCaseStrg(.fields("Bachiller"))
                        Case "02": Codigo1 = "SEGUNDO CURSO CICLO DIVERSIFICADO DEL " & UCaseStrg(.fields("Bachiller"))
                        Case "03": Codigo1 = "TERCER CURSO CICLO DIVERSIFICADO DEL " & UCaseStrg(.fields("Bachiller"))
                      End Select
           End Select
           If .fields("Especialidad") <> Ninguno Then Codigo2 = "ESPECIALIZACIÓN " & UCaseStrg(.fields("Especialidad"))
           Codigo3 = .fields("Materia")
           Codigo4 = .fields("Profesor")
           Encabezado_Lista_Alumnos SizeLetra, FechaInicial, Codigo, Codigo1, Codigo2, TipoNomina
        End If
       .MoveNext
     Loop
     PosColumna = 10.5
     If TipoNomina = 2 Or TipoNomina = 3 Then
        For I = 1 To KE
            Imprimir_Linea_V PosColumna, InicioY, PosLinea
            If TipoNomina = 3 Then
               PosColumna = PosColumna + 1.5
            Else
               PosColumna = PosColumna + 0.6
            End If
        Next I
     End If
End If
End With
Imprimir_Linea_H PosLinea, 2, 19.5
PosLinea = PosLinea - 0.1
If TipoNomina = 2 Or TipoNomina = 3 Then
   InicioY = InicioY + 0.05
   Imprimir_Linea_V 2, InicioY, PosLinea
   Imprimir_Linea_V 2.55, InicioY, PosLinea
   Imprimir_Linea_V 19.5, InicioY, PosLinea
   PosColumna = 10.5
   Imprimir_Linea_V PosColumna, InicioY, PosLinea
   If J > 1 Then
      For I = 1 To KE
          Imprimir_Linea_V PosColumna, InicioY, PosLinea
          If TipoNomina = 3 Then
             PosColumna = PosColumna + 1.5
          Else
             PosColumna = PosColumna + 0.6
          End If
      Next I
   End If
   PosLinea = PosLinea + 1.6
   PrinterTexto 2, PosLinea, "FECHA:"
   PosLinea = PosLinea + 0.3
   Imprimir_Linea_H PosLinea, 3, 8
   Imprimir_Linea_H PosLinea, 15, 18
   PosLinea = PosLinea + 0.05
   PrinterTexto 15, PosLinea, "FIRMA DEL PROFESOR:"
End If
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Function Leer_Datos_del_Curso(Curso As String, Optional Tipo_Texto As Byte, Optional SubNotas As Boolean) As String
Dim Curso_Aux As String
Dim Idx As Integer
'Si Tipo_Texo = 0 Con paralelo
'               1 Sin Paralelo
'               2 Sin "EN EL" articulo

Dim AdoDBCurso As ADODB.Recordset
Dim sSQLCurso As String
 'Activamos el espacio de consulta
  RatonReloj
  Set AdoDBCurso = New ADODB.Recordset
  AdoDBCurso.CursorType = adOpenStatic
  AdoDBCurso.CursorLocation = adUseClient
  If Len(Curso) = 7 Then
    'Consultamos los datos del Curso
     sSQLCurso = "SELECT * " _
               & "FROM Catalogo_Cursos " _
               & "WHERE Periodo = '" & Periodo_Contable & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Curso = '" & Curso & "' "
     sSQLCurso = CompilarSQL(sSQLCurso)
     AdoDBCurso.open sSQLCurso, AdoStrCnn, , , adCmdText
     With AdoDBCurso
      If .RecordCount > 0 Then
          Dato_Curso.Curso_Anio = .fields("Descripcion")
          Dato_Curso.Bachiller = .fields("Bachiller")
          Dato_Curso.Ciclo = .fields("Ciclo")
          Dato_Curso.Codigo_Titulo = .fields("Codigo_Titulo")
          Dato_Curso.Curso = .fields("Curso")
          Dato_Curso.Paralelo = .fields("Paralelo")
          Dato_Curso.Descripcion = .fields("Descripcion")
          Dato_Curso.Especialidad = .fields("Especialidad")
          Dato_Curso.Figura_Profesional = .fields("Figura_Profesional")
          Dato_Curso.Seccion = .fields("Seccion")
          Dato_Curso.Tipo_Titulo = .fields("Tipo_Titulo")
          Dato_Curso.Titulo = .fields("Titulo")
          Dato_Curso.Curso_Superior = .fields("Curso_Superior")
          Dato_Curso.Nombre_Largo = ""
          If Tipo_Texto = 2 Then Dato_Curso.Nombre_Largo = "EN EL "
          Dato_Curso.Nombre_Largo = Dato_Curso.Nombre_Largo & Dato_Curso.Bachiller
          If Dato_Curso.Especialidad <> Ninguno Then Dato_Curso.Nombre_Largo = Dato_Curso.Nombre_Largo & " " & Dato_Curso.Especialidad
          If Dato_Curso.Figura_Profesional <> Ninguno Then Dato_Curso.Nombre_Largo = Dato_Curso.Nombre_Largo & " " & Dato_Curso.Figura_Profesional
          If Tipo_Texto = 0 And Dato_Curso.Paralelo <> Ninguno Then
             Dato_Curso.Nombre_Largo = Dato_Curso.Nombre_Largo & " " & Dato_Curso.Paralelo
          End If
         'MsgBox Dato_Curso.Nombre_Largo
      End If
     End With
     AdoDBCurso.Close
     
     sSQLCurso = "SELECT Descripcion " _
               & "FROM Catalogo_Cursos " _
               & "WHERE Periodo = '" & Periodo_Contable & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Curso = '" & CodigoCuentaSup(Curso) & "' "
     sSQLCurso = CompilarSQL(sSQLCurso)
     AdoDBCurso.open sSQLCurso, AdoStrCnn, , , adCmdText
     With AdoDBCurso
      If .RecordCount > 0 Then
          Dato_Curso.Curso_Texto = .fields("Descripcion")
      End If
     End With
     AdoDBCurso.Close
     
     Idx = 1
    'Consultamos los alumnos del curso
     sSQLCurso = "SELECT C.Cliente,C.Sexo,C.Codigo,C.CI_RUC,CM.Grupo_No " _
               & "FROM Clientes As C,Clientes_Matriculas As CM " _
               & "WHERE CM.Periodo = '" & Periodo_Contable & "' " _
               & "AND CM.Item = '" & NumEmpresa & "' " _
               & "AND CM.Grupo_No = '" & Curso & "' " _
               & "AND C.FA <> " & Val(adFalse) & " " _
               & "AND C.Codigo = CM.Codigo " _
               & "ORDER BY C.Cliente "
     sSQLCurso = CompilarSQL(sSQLCurso)
     AdoDBCurso.open sSQLCurso, AdoStrCnn, , , adCmdText
     With AdoDBCurso
      If .RecordCount > 0 Then
          Dato_Curso.ContAlumnos = .RecordCount
          ReDim Dato_Curso.CodigoC(1 To Dato_Curso.ContAlumnos) As String
          ReDim Dato_Curso.Sexo(1 To Dato_Curso.ContAlumnos) As String
          ReDim Dato_Curso.Alumno(1 To Dato_Curso.ContAlumnos) As String
          ReDim Dato_Curso.CI_RUC(1 To Dato_Curso.ContAlumnos) As String
          ReDim Dato_Curso.NotaPQ(1 To Dato_Curso.ContAlumnos) As Currency
          ReDim Dato_Curso.NotaSQ(1 To Dato_Curso.ContAlumnos) As Currency
          ReDim Dato_Curso.NotaTQ(1 To Dato_Curso.ContAlumnos) As Currency
          ReDim Dato_Curso.NotaFinal(1 To Dato_Curso.ContAlumnos) As Currency
          Do While Not .EOF
             Dato_Curso.NotaPQ(Idx) = 0
             Dato_Curso.NotaSQ(Idx) = 0
             Dato_Curso.NotaTQ(Idx) = 0
             Dato_Curso.NotaFinal(Idx) = 0
             Dato_Curso.CodigoC(Idx) = .fields("Codigo")
             Dato_Curso.CI_RUC(Idx) = .fields("CI_RUC")
             Dato_Curso.Alumno(Idx) = .fields("Cliente")
             Dato_Curso.Sexo(Idx) = .fields("Sexo")
             Idx = Idx + 1
            .MoveNext
          Loop
      End If
     End With
     AdoDBCurso.Close
     Idx = 1
    'Consultamos las materias del curso que se promedian al Ministerio de Educacion
     sSQLCurso = "SELECT CE.*,CM.Materia " _
               & "FROM Catalogo_Estudiantil As CE,Catalogo_Materias As CM " _
               & "WHERE CE.Periodo = '" & Periodo_Contable & "' " _
               & "AND CE.Item = '" & NumEmpresa & "' " _
               & "AND MidStrg(CE.CodigoE,1,7) = '" & Curso & "' " _
               & "AND CE.TC = 'M' "
     If SubNotas Then
        sSQLCurso = sSQLCurso & "AND CE.CodMatP <> '" & Ninguno & "' "
     Else
        sSQLCurso = sSQLCurso & "AND CE.CodMatP = '" & Ninguno & "' "
     End If
     sSQLCurso = sSQLCurso _
               & "AND CE.Item = CM.Item " _
               & "AND CE.Periodo = CM.Periodo " _
               & "AND CE.CodMat = CM.CodMat " _
               & "ORDER BY CE.CodigoE "
     sSQLCurso = CompilarSQL(sSQLCurso)
     AdoDBCurso.open sSQLCurso, AdoStrCnn, , , adCmdText
     With AdoDBCurso
      If .RecordCount > 0 Then
          Dato_Curso.ContMat = .RecordCount
          ReDim Dato_Curso.Materia(1 To Dato_Curso.ContMat) As String
          ReDim Dato_Curso.CodMat(1 To Dato_Curso.ContMat) As String
          ReDim Dato_Curso.PosXMat(1 To Dato_Curso.ContMat) As Single
          Do While Not .EOF
             Dato_Curso.PosXMat(Idx) = 0
             Dato_Curso.Materia(Idx) = .fields("Materia")
             Dato_Curso.CodMat(Idx) = .fields("CodMat")
             Idx = Idx + 1
            .MoveNext
          Loop
      End If
     End With
     AdoDBCurso.Close
     Idx = 1
    'Consultamos todas las materias del curso incluida las Submaterias
     sSQLCurso = "SELECT CE.*,CM.Materia " _
               & "FROM Catalogo_Estudiantil As CE,Catalogo_Materias As CM " _
               & "WHERE CE.Periodo = '" & Periodo_Contable & "' " _
               & "AND CE.Item = '" & NumEmpresa & "' " _
               & "AND MidStrg(CE.CodigoE,1,7) = '" & Curso & "' " _
               & "AND CE.TC = 'M' " _
               & "AND CE.Item = CM.Item " _
               & "AND CE.Periodo = CM.Periodo " _
               & "AND CE.CodMat = CM.CodMat " _
               & "ORDER BY CE.CodigoE "
     sSQLCurso = CompilarSQL(sSQLCurso)
     AdoDBCurso.open sSQLCurso, AdoStrCnn, , , adCmdText
     With AdoDBCurso
      If .RecordCount > 0 Then
          Dato_Curso.ContMatT = .RecordCount
          ReDim Dato_Curso.MateriaT(1 To Dato_Curso.ContMatT) As String
          ReDim Dato_Curso.CodMatPT(1 To Dato_Curso.ContMatT) As String
          Do While Not .EOF
             Dato_Curso.MateriaT(Idx) = .fields("Materia")
             Dato_Curso.CodMatPT(Idx) = .fields("CodMat")
             Idx = Idx + 1
            .MoveNext
          Loop
      End If
     End With
     AdoDBCurso.Close
    'MsgBox Dato_Curso.Curso_Superior
     Leer_Datos_del_Curso = TrimStrg(Dato_Curso.Nombre_Largo)
  Else
     Leer_Datos_del_Curso = Ninguno
  End If
  RatonNormal
End Function

Public Function Indice_Materia(CodMat As String) As Byte
Dim IdMat As Byte
Dim Posicion As Byte
  Posicion = 0
  For IdMat = 1 To Dato_Curso.ContMat
      If Dato_Curso.CodMat(IdMat) = CodMat Then Posicion = IdMat
  Next IdMat
  Indice_Materia = Posicion
End Function

Public Sub Imprimir_Mejor_Promedio(Datas As Adodc, _
                                   FinDoc As Boolean, _
                                   FormaImp As Byte, _
                                   SizeLetra As Integer, _
                                   Tipo_Impresion As Integer, _
                                   OpcBimestre As Byte, _
                                   Optional Decimales As String)
On Error GoTo Errorhandler
Dim CantDec  As Long
Dim Idx As Long
Dim Jdx As Long
Dim LenCamposDec As Long
Dim CantDecCampo As String
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoArialNarrow, Orientacion_Pagina
Ancho(CantCampos) = AnchoPapel
Pagina = 1
'Iniciamos la impresion
Contador = 0
Printer.FontBold = False

With Datas.Recordset
If .RecordCount > 0 Then
    CantCampos = .fields.Count
    'Array para almacenar el ancho de cada columna
     ReDim Vect_Dec(CantCampos) As Campos_Decimal
    'Enceramos la impresion
     For Idx = 0 To CantCampos - 1
         Vect_Dec(Idx).Campo = .fields(Idx).Name
         Vect_Dec(Idx).CantDec = 2
         Vect_Dec(Idx).AnchoCampo = 0
     Next Idx
     
     For Idx = 0 To CantCampos - 1
        'MsgBox Decimales & vbCrLf & Vect_Dec(Col).Campo & vbCrLf & Vect_Dec(Col).CantDec & vbCrLf & Vect_Dec(Col).AnchoCampo
         If Decimales <> "" Then
            LenCamposDec = Len(.fields(Idx).Name)
            For Jdx = 1 To Len(Decimales)
                If .fields(Idx).Name = MidStrg(Decimales, Jdx, LenCamposDec) Then
                   CantDecCampo = ""
                   For CantDec = Jdx + LenCamposDec To Len(Decimales)
                       If MidStrg(Decimales, CantDec, 1) = "|" Then CantDec = Len(Decimales) + 1
                       CantDecCampo = CantDecCampo & MidStrg(Decimales, CantDec, 1)
                   Next CantDec
                   Vect_Dec(Idx).CantDec = Val(CantDecCampo)
                End If
            Next Jdx
         End If
     Next Idx
    .MoveFirst
     Encabezado_Institucion Ancho(0), AnchoPapel
     Printer.FontSize = SizeLetra
     Printer.FontName = TipoArialNarrow
     Codigo = .fields("Grupo")
     Printer.FontItalic = False
     Printer.FontBold = True
     Printer.FontUnderline = True
     PrinterTexto 1, PosLinea, "E S T U D I A N T E"
     PrinterTexto 8, PosLinea, "C U R S O"
     Select Case OpcBimestre
       Case 1
            PrinterTexto 16.3, PosLinea, "Promedio_PQ"
       Case 2
            PrinterTexto 16.3, PosLinea, "Promedio_SQ"
       Case 3
            PrinterTexto 12.5, PosLinea, "Promedio_PQ"
            PrinterTexto 14.5, PosLinea, "Promedio_SQ"
            PrinterTexto 16.5, PosLinea, "Promedio_TQ"
            PrinterTexto 18.5, PosLinea, "Promedio"
     End Select
     Printer.FontBold = False
     Printer.FontUnderline = False
     PosLinea = PosLinea + 0.5
     'PrinterFields Ancho(0), PosLinea, .Fields("Curso")
     Do While Not .EOF
        If Codigo <> .fields("Grupo") Then Codigo = .fields("Grupo")
        PrinterFields 1, PosLinea, .fields("Alumno")
        Select Case OpcBimestre
          Case 1, 2
               PrinterFields 8, PosLinea, .fields("Curso")
          Case Else
               PrinterFields 8, PosLinea, .fields("Grupo")
        End Select
        Select Case OpcBimestre
          Case 1
               PrinterFields 15.5, PosLinea, .fields("Promedio_PQ")
          Case 2
               PrinterFields 15.5, PosLinea, .fields("Promedio_SQ")
          Case 3
               PrinterFields 11.7, PosLinea, .fields("Tot_PQBim1")
               PrinterFields 13.7, PosLinea, .fields("Tot_PQBim2")
               PrinterFields 15.7, PosLinea, .fields("Tot_PQBim3")
               PrinterFields 17.5, PosLinea, .fields("Tot_PrompQ")
''               PrinterFields 11.7, PosLinea, .Fields("Promedio_PQ")
''               PrinterFields 13.7, PosLinea, .Fields("Promedio_SQ")
''               PrinterFields 15.7, PosLinea, .Fields("Promedio_TQ")
''               PrinterFields 17.5, PosLinea, .Fields("Promedio")
        End Select
        PosLinea = PosLinea + 0.4
        If Cuadricula Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro
           'MsgBox Ancho(0) & vbCrLf & Ancho(CantCampos)
           PosLinea = PosLinea + 0.05
        End If
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
           Printer.NewPage
           Encabezado_Institucion Ancho(0), AnchoPapel
           Printer.FontSize = SizeLetra
           Printer.FontName = TipoArialNarrow
           Codigo = .fields("Grupo")
           Printer.FontItalic = False
           Printer.FontBold = True
           Printer.FontUnderline = True
           PrinterTexto 1, PosLinea, "E S T U D I A N T E"
           PrinterTexto 8, PosLinea, "C U R S O"
           Select Case OpcBimestre
             Case 1
                  PrinterTexto 16.3, PosLinea, "Promedio_PQ"
             Case 2
                  PrinterTexto 16.3, PosLinea, "Promedio_SQ"
             Case 3
                  PrinterTexto 12.5, PosLinea, "Promedio_PQ"
                  PrinterTexto 14.5, PosLinea, "Promedio_SQ"
                  PrinterTexto 16.5, PosLinea, "Promedio_TQ"
                  PrinterTexto 18.5, PosLinea, "Promedio"
           End Select
           Printer.FontBold = False
           Printer.FontUnderline = False
           PosLinea = PosLinea + 0.5
           'PrinterFields Ancho(0), PosLinea, .Fields("Grupo")
        End If
       .MoveNext
     Loop
End If
End With
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos), Negro, True
RatonNormal
MensajeEncabData = ""
If FinDoc Then Printer.EndDoc
Cuadricula = False
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Redondear_Nota(Campo As String, Decimales As Byte)
   sSQL = "UPDATE Trans_Notas "
   If Decimales <= 0 Then
      sSQL = sSQL & "SET " & Campo & " = ROUND(" & Campo & ",0,0) "
   Else
      sSQL = sSQL & "SET " & Campo & " = ROUND(" & Campo & "," & Decimales & ",0) "
   End If
   sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' "
   Ejecutar_SQL_SP sSQL
End Sub

Public Sub Actualizar_Cursos()
    RatonReloj
   'Actualizamos Grupos
    If Periodo_Contable = Ninguno Then
       sSQL = "UPDATE Clientes_Matriculas " _
            & "SET Grupo_No = '.' " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' "
       Ejecutar_SQL_SP sSQL
       If SQL_Server Then
          sSQL = "UPDATE Clientes_Matriculas " _
               & "SET Grupo_No = C.Grupo " _
               & "FROM Clientes_Matriculas As CM,Clientes As C "
       Else
          sSQL = "UPDATE Clientes_Matriculas As CM,Clientes As C " _
               & "SET CM.Grupo_No = C.Grupo "
       End If
       sSQL = sSQL & "WHERE CM.Item = '" & NumEmpresa & "' " _
            & "AND CM.Periodo = '" & Periodo_Contable & "' " _
            & "AND C.FA <> " & Val(adFalse) & " " _
            & "AND CM.Codigo = C.Codigo "
       Ejecutar_SQL_SP sSQL
    End If
    RatonNormal
   'Reindexar Notas y Auxiliares
    
    If SQL_Server Then
       sSQL = "UPDATE Trans_Notas " _
            & "SET Id_No = CE.Id_No " _
            & "FROM Trans_Notas As TN, Catalogo_Estudiantil As CE "
    Else
       sSQL = "UPDATE Trans_Notas As TN, Catalogo_Estudiantil As CE " _
            & "SET TN. Id_No = CE.Id_No "
    End If
    sSQL = sSQL _
         & "WHERE TN.Item = '" & NumEmpresa & "' " _
         & "AND TN.Periodo = '" & Periodo_Contable & "' " _
         & "AND TN.CodE = MidStrg(CE.CodigoE,1,7) " _
         & "AND TN.CodMat = CE.CodMat " _
         & "AND TN.Item = CE.Item " _
         & "AND TN.Periodo = CE.Periodo " _
         & ""
    Ejecutar_SQL_SP sSQL
    
    If SQL_Server Then
       sSQL = "UPDATE Trans_Notas_Auxiliares " _
            & "SET Id_No = CE.Id_No " _
            & "FROM Trans_Notas_Auxiliares As TN, Catalogo_Estudiantil As CE "
    Else
       sSQL = "UPDATE Trans_Notas_Auxiliares As TN, Catalogo_Estudiantil As CE " _
            & "SET TN. Id_No = CE.Id_No "
    End If
    sSQL = sSQL _
         & "WHERE TN.Item = '" & NumEmpresa & "' " _
         & "AND TN.Periodo = '" & Periodo_Contable & "' " _
         & "AND TN.CodE = MidStrg(CE.CodigoE,1,7) " _
         & "AND TN.CodMat = CE.CodMat " _
         & "AND TN.Item = CE.Item " _
         & "AND TN.Periodo = CE.Periodo " _
         & ""
    Ejecutar_SQL_SP sSQL
    RatonNormal
End Sub

Public Sub Autorizar_Ingreso_Notas()
''''Dim IAut As Integer
''''   For IAut = 0 To 11
''''       Autoizar_Notas(IAut) = Ninguno
''''   Next IAut
''''   IAut = 0
''''   sSQL = "SELECT * " _
''''        & "FROM Catalogo_Periodo_Lectivo " _
''''        & "WHERE Item = '" & NumEmpresa & "' " _
''''        & "AND Periodo = '" & Periodo_Contable & "' "
''''   Select_Adodc AdoAutorizar, sSQL
''''   With AdoAutorizar.Recordset
''''    If .RecordCount > 0 Then
''''        If MidStrg(FormatoLibreta, 1, 9) = "TRIMESTRE" Then
''''           LstAutoriza.AddItem "Primer Trimestre Primer Periodo", IAut
''''           LstAutoriza.Selected(IAut) = CBool(.Fields("NPQP1"))
''''           Autoizar_Notas(IAut) = "NPQP1": IAut = IAut + 1
''''           LstAutoriza.AddItem "Primer Trimestre Segundo Periodo", IAut
''''           LstAutoriza.Selected(IAut) = CBool(.Fields("NPQP2"))
''''           Autoizar_Notas(IAut) = "NPQP2": IAut = IAut + 1
''''           LstAutoriza.AddItem "Examen Primer Trimestre", IAut
''''           LstAutoriza.Selected(IAut) = CBool(.Fields("NPQEX"))
''''           Autoizar_Notas(IAut) = "NPQEX": IAut = IAut + 1
''''
''''           LstAutoriza.AddItem "Segundo Trimestre Primer Periodo", IAut
''''           LstAutoriza.Selected(IAut) = CBool(.Fields("NSQP1"))
''''           Autoizar_Notas(IAut) = "NSQP1": IAut = IAut + 1
''''           LstAutoriza.AddItem "Segundo Trimestre Segundo Periodo", IAut
''''           LstAutoriza.Selected(IAut) = CBool(.Fields("NSQP2"))
''''           Autoizar_Notas(IAut) = "NSQP2": IAut = IAut + 1
''''           LstAutoriza.AddItem "Examen Segundo Trimestre", IAut
''''           LstAutoriza.Selected(IAut) = CBool(.Fields("NSQEX"))
''''           Autoizar_Notas(IAut) = "NSQEX": IAut = IAut + 1
''''
''''           LstAutoriza.AddItem "Tercer Trimestre Primer Periodo", IAut
''''           LstAutoriza.Selected(IAut) = CBool(.Fields("NTQP1"))
''''           Autoizar_Notas(IAut) = "NTQP1": IAut = IAut + 1
''''           LstAutoriza.AddItem "Tercer Trimestre Segundo Periodo", IAut
''''           LstAutoriza.Selected(IAut) = CBool(.Fields("NTQP2"))
''''           Autoizar_Notas(IAut) = "NTQP2": IAut = IAut + 1
''''           LstAutoriza.AddItem "Examen Tercer Trimestre", IAut
''''           LstAutoriza.Selected(IAut) = CBool(.Fields("NTQEX"))
''''           Autoizar_Notas(IAut) = "NTQEX": IAut = IAut + 1
''''        Else
''''           LstAutoriza.AddItem "Primer Quimestre Primer Periodo", IAut
''''           LstAutoriza.Selected(IAut) = CBool(.Fields("NPQP1"))
''''           Autoizar_Notas(IAut) = "NPQP1": IAut = IAut + 1
''''           LstAutoriza.AddItem "Primer Quimestre Segundo Periodo", IAut
''''           LstAutoriza.Selected(IAut) = CBool(.Fields("NPQP2"))
''''           Autoizar_Notas(IAut) = "NPQP2": IAut = IAut + 1
''''           LstAutoriza.AddItem "Examen Primer Quimestre", IAut
''''           LstAutoriza.Selected(IAut) = CBool(.Fields("NPQEX"))
''''           Autoizar_Notas(IAut) = "NPQEX": IAut = IAut + 1
''''
''''           LstAutoriza.AddItem "Segundo Quimestre Primer Periodo", IAut
''''           LstAutoriza.Selected(IAut) = CBool(.Fields("NSQP1"))
''''           Autoizar_Notas(IAut) = "NSQP1": IAut = IAut + 1
''''           LstAutoriza.AddItem "Segundo Quimestre Segundo Periodo", IAut
''''           LstAutoriza.Selected(IAut) = CBool(.Fields("NSQP2"))
''''           Autoizar_Notas(IAut) = "NSQP2": IAut = IAut + 1
''''           LstAutoriza.AddItem "Examen Segundo Quimestre", IAut
''''           LstAutoriza.Selected(IAut) = CBool(.Fields("NSQEX"))
''''           Autoizar_Notas(IAut) = "NSQEX": IAut = IAut + 1
''''        End If
''''        LstAutoriza.AddItem "Supletorio", IAut
''''        LstAutoriza.Selected(IAut) = CBool(.Fields("NSUPL"))
''''        Autoizar_Notas(IAut) = "NSUPL": IAut = IAut + 1
''''        LstAutoriza.AddItem "Grado", IAut
''''        LstAutoriza.Selected(IAut) = CBool(.Fields("NGRADO"))
''''        Autoizar_Notas(IAut) = "NGRADO": IAut = IAut + 1
''''    End If
''''   End With
''''   RatonNormal
End Sub

Public Function OpcPeriodo(Notas_P As String, LstPeriodo As ListBox) As Boolean
Dim Resultado As Boolean
    Resultado = False
    If MidStrg(FormatoLibreta, 1, 9) = "TRIMESTRE" Then
       If Notas_P = "PQBim1" And LstPeriodo.Text = "Primer Trimestre Primer Periodo" Then Resultado = True
       If Notas_P = "PQBim2" And LstPeriodo.Text = "Primer Trimestre Segundo Periodo" Then Resultado = True
       If Notas_P = "PQ" And LstPeriodo.Text = "Promedio Primer Trimestre" Then Resultado = True
       
       If Notas_P = "SQBim1" And LstPeriodo.Text = "Segundo Trimestre Primer Periodo" Then Resultado = True
       If Notas_P = "SQBim2" And LstPeriodo.Text = "Segundo Trimestre Segundo Periodo" Then Resultado = True
       If Notas_P = "SQ" And LstPeriodo.Text = "Promedio Segundo Trimestre" Then Resultado = True
       
       If Notas_P = "TQBim1" And LstPeriodo.Text = "Tercer Trimestre Primer Periodo" Then Resultado = True
       If Notas_P = "TQBim2" And LstPeriodo.Text = "Tercer Trimestre Segundo Periodo" Then Resultado = True
       If Notas_P = "TQ" And LstPeriodo.Text = "Promedio Tercer Trimestre" Then Resultado = True
       
       If Notas_P = "PF" And LstPeriodo.Text = "Todos Trimestres" Then Resultado = True
    ElseIf MidStrg(FormatoLibreta, 1, 9) = "QUIMESTRE" Then
       If Notas_P = "PQBim1" And LstPeriodo.Text = "Primer Quimestre Primer Parcial" Then Resultado = True
       If Notas_P = "PQBim2" And LstPeriodo.Text = "Primer Quimestre Segundo Parcial" Then Resultado = True
       If Notas_P = "PQBim3" And LstPeriodo.Text = "Primer Quimestre Tercer Parcial" Then Resultado = True
       If Notas_P = "ExamenPQ" And LstPeriodo.Text = "Primer Quimestre Examen" Then Resultado = True
       If Notas_P = "PQ" And LstPeriodo.Text = "Promedio Primer Quimestre" Then Resultado = True
       
       If Notas_P = "SQBim1" And LstPeriodo.Text = "Segundo Quimestre Primer Parcial" Then Resultado = True
       If Notas_P = "SQBim2" And LstPeriodo.Text = "Segundo Quimestre Segundo Parcial" Then Resultado = True
       If Notas_P = "SQBim3" And LstPeriodo.Text = "Segundo Quimestre Tercer Parcial" Then Resultado = True
       If Notas_P = "ExamenSQ" And LstPeriodo.Text = "Segundo Quimestre Examen" Then Resultado = True
       If Notas_P = "SQ" And LstPeriodo.Text = "Promedio Segundo Quimestre" Then Resultado = True
       
       If Notas_P = "PF" And LstPeriodo.Text = "Todos los Quimestres" Then Resultado = True
    Else
       If Notas_P = "PQBim1" And LstPeriodo.Text = "Primer Quimestre Primer Periodo" Then Resultado = True
       If Notas_P = "PQBim2" And LstPeriodo.Text = "Primer Quimestre Segundo Periodo" Then Resultado = True
       If Notas_P = "PQ" And LstPeriodo.Text = "Promedio Primer Quimestre" Then Resultado = True
       
       If Notas_P = "SQBim1" And LstPeriodo.Text = "Segundo Quimestre Primer Periodo" Then Resultado = True
       If Notas_P = "SQBim2" And LstPeriodo.Text = "Segundo Quimestre Segundo Periodo" Then Resultado = True
       If Notas_P = "SQ" And LstPeriodo.Text = "Promedio Segundo Quimestre" Then Resultado = True
       
       If Notas_P = "PF" And LstPeriodo.Text = "Todos los Periodos" Then Resultado = True
    End If
   OpcPeriodo = Resultado
End Function

Public Function Seleccionar_Periodo(LstPeriodo As ListBox, _
                                    Optional CodigoAlum As String, _
                                    Optional CodCurso As String) As Byte
Dim OpcionNotas As Byte
Dim AdoEstudiante As ADODB.Recordset
    RatonReloj
    If CodCurso = "" Then CodCurso = Ninguno
    If CodigoAlum = "" Then CodigoAlum = Ninguno
    sSQL = "SELECT * " _
         & "FROM Trans_Asistencia " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Codigo = '" & CodigoAlum & "' " _
         & "AND CodE = '" & CodCurso & "' "
    Select_AdoDB AdoEstudiante, sSQL
    Valor = 0
    Atrasos = 0
    Faltas_Just = 0
    Faltas_Injust = 0
    Dias_Laborados = 0
    
    Atrasos1 = 0
    Faltas_Just1 = 0
    Faltas_Injust1 = 0
    Dias_Laborados1 = 0
    
    Atrasos2 = 0
    Faltas_Just2 = 0
    Faltas_Injust2 = 0
    Dias_Laborados2 = 0
    
    Atrasos3 = 0
    Faltas_Just3 = 0
    Faltas_Injust3 = 0
    Dias_Laborados3 = 0
    
    SQLTAI = Ninguno
    SQLAIC = Ninguno
    SQLAGC = Ninguno
    SQLL = Ninguno
    SQLExaP = Ninguno
    SQLNotas = Ninguno
    SQLBim1 = Ninguno
    SQLBim2 = Ninguno
    SQLBim3 = Ninguno
    SQLProm = Ninguno
    SQLExamen = Ninguno
    SQLPromQ = Ninguno
    SQLConductaQ = Ninguno
    SQLQPX = Ninguno
    SQLQEX = Ninguno
    SQLInforme = Ninguno
    CadenaParcial = LstPeriodo.Text
    OpcionNotas = 0
   'Determinamos que tipo de informacion necesitamos presentar
    With AdoEstudiante
        If .RecordCount > 0 Then
            Real1 = .fields("ConductaPQ1")
            Real2 = .fields("ConductaPQ2")
            Real5 = .fields("ConductaPQ3")
            Real3 = .fields("ConductaSQ1")
            Real4 = .fields("ConductaSQ2")
            Real6 = .fields("ConductaSQ3")
        End If
        Select Case CadenaParcial
          Case "Primer Quimestre Primer Parcial", _
               "Primer Quimestre Primer Periodo", _
               "Primer Trimestre Primer Periodo"
               SQLTAI = "PQTAI1"
               SQLAIC = "PQAIC1"
               SQLAGC = "PQAGC1"
               SQLL = "PQL1"
               SQLExaP = "PQExaP1"
               SQLNotas = "PQBim1"
                
               SQLBim1 = "PQBim1"
               SQLBim2 = "PQBim2"
               SQLBim3 = "PQBim3"
               SQLProm = "PQBim1"
               SQLExamen = "ExamenPQ"
               SQLPromQ = "PromPQ"
               SQLConductaQ = "ConductaPQ1"
               SQLInforme = "Informe_PQ1"
               If .RecordCount > 0 Then
                   Valor = .fields("ConductaPQ1")
                   Dias_Laborados1 = .fields("PQDias1")
                   Faltas_Just1 = .fields("PQBFJ1")
                   Faltas_Injust1 = .fields("PQBFI1")
                   Atrasos1 = .fields("PQBA1")
               End If
               OpcionNotas = 1
          Case "Primer Quimestre Segundo Parcial", _
               "Primer Quimestre Segundo Periodo", _
               "Primer Trimestre Segundo Periodo"
               SQLTAI = "PQTAI2"
               SQLAIC = "PQAIC2"
               SQLAGC = "PQAGC2"
               SQLL = "PQL2"
               SQLExaP = "PQExaP2"
               SQLNotas = "PQBim2"
                
               SQLBim1 = "PQBim1"
               SQLBim2 = "PQBim2"
               SQLBim3 = "PQBim3"
               SQLProm = "PQBim2"
               SQLExamen = "ExamenPQ"
               SQLPromQ = "PromPQ"
               SQLConductaQ = "ConductaPQ2"
               SQLInforme = "Informe_PQ2"
               If .RecordCount > 0 Then
                   Valor = .fields("ConductaPQ2")
                   Dias_Laborados2 = .fields("PQDias2")
                   Faltas_Just2 = .fields("PQBFJ2")
                   Faltas_Injust2 = .fields("PQBFI2")
                   Atrasos2 = .fields("PQBA2")
               End If
               OpcionNotas = 2
          Case "Primer Quimestre Tercer Parcial"
               SQLTAI = "PQTAI3"
               SQLAIC = "PQAIC3"
               SQLAGC = "PQAGC3"
               SQLL = "PQL3"
               SQLExaP = "PQExaP3"
               SQLNotas = "PQBim3"
               
               SQLBim1 = "PQBim1"
               SQLBim2 = "PQBim2"
               SQLBim3 = "PQBim3"
               SQLProm = "PQBim3"
               SQLExamen = "ExamenPQ"
               SQLPromQ = "PromPQ"
               SQLConductaQ = "ConductaPQ3"
               SQLInforme = "Informe_PQ3"
               If .RecordCount > 0 Then
                   Valor = .fields("ConductaPQ3")
                   Dias_Laborados3 = .fields("PQDias3")
                   Faltas_Just3 = .fields("PQBFJ3")
                   Faltas_Injust3 = .fields("PQBFI3")
                   Atrasos3 = .fields("PQBA3")
               End If
               OpcionNotas = 3
          Case "Promedio Primer Quimestre", _
               "Promedio Primer Trimestre"
               SQLBim1 = "PQBim1"
               SQLBim2 = "PQBim2"
               SQLBim3 = "PQBim3"
               SQLProm = "PQBim3"
               SQLPromQ = "PromPQ"
               SQLExamen = "ExamenPQ"
               SQLQPX = "PQ_PP"
               SQLQEX = "PQ_PE"
               SQLInforme = "Informe_PQ"
               If .RecordCount > 0 Then
                   Valor = .fields("ConductaPQ3")
                   Dias_Laborados1 = .fields("PQDias1")
                   Faltas_Just1 = .fields("PQBFJ1")
                   Faltas_Injust1 = .fields("PQBFI1")
                   Atrasos1 = .fields("PQBA1")
                   
                   Dias_Laborados2 = .fields("PQDias2")
                   Faltas_Just2 = .fields("PQBFJ2")
                   Faltas_Injust2 = .fields("PQBFI2")
                   Atrasos2 = .fields("PQBA2")
                   
                   Dias_Laborados3 = .fields("PQDias3")
                   Faltas_Just3 = .fields("PQBFJ3")
                   Faltas_Injust3 = .fields("PQBFI3")
                   Atrasos3 = .fields("PQBA3")
                   
                   Dias_Laborados = .fields("PQDias1") + .fields("PQDias2") + .fields("PQDias3")
                   Faltas_Just = .fields("PQBFJ1") + .fields("PQBFJ2") + .fields("PQBFJ3")
                   Faltas_Injust = .fields("PQBFI1") + .fields("PQBFI2") + .fields("PQBFI3")
                   Atrasos = .fields("PQBA1") + .fields("PQBA2") + .fields("PQBA3")
               End If
               OpcionNotas = 4
          Case "Segundo Quimestre Primer Parcial", _
               "Segundo Quimestre Primer Periodo", _
               "Segundo Trimestre Primer Periodo"
               SQLTAI = "SQTAI1"
               SQLAIC = "SQAIC1"
               SQLAGC = "SQAGC1"
               SQLL = "SQL1"
               SQLExaP = "SQExaP1"
               SQLNotas = "SQBim1"
                
               SQLBim1 = "SQBim1"
               SQLBim2 = "SQBim2"
               SQLBim3 = "SQBim3"
               SQLProm = "SQBim1"
               SQLExamen = "ExamenSQ"
               SQLPromQ = "PromSQ"
               SQLConductaQ = "ConductaSQ1"
               SQLInforme = "Informe_SQ1"
               If .RecordCount > 0 Then
                   Valor = .fields("ConductaSQ1")
                   Dias_Laborados1 = .fields("SQDias1")
                   Faltas_Just1 = .fields("SQBFJ1")
                   Faltas_Injust1 = .fields("SQBFI1")
                   Atrasos1 = .fields("SQBA1")
               End If
               OpcionNotas = 1
          Case "Segundo Quimestre Segundo Parcial", _
               "Segundo Quimestre Segundo Periodo", _
               "Segundo Trimestre Segundo Periodo"
               SQLTAI = "SQTAI2"
               SQLAIC = "SQAIC2"
               SQLAGC = "SQAGC2"
               SQLL = "SQL2"
               SQLExaP = "SQExaP2"
               SQLNotas = "SQBim2"
                
               SQLBim1 = "SQBim1"
               SQLBim2 = "SQBim2"
               SQLBim3 = "SQBim3"
               SQLProm = "SQBim2"
               SQLExamen = "ExamenSQ"
               SQLPromQ = "PromSQ"
               SQLConductaQ = "ConductaSQ2"
               SQLInforme = "Informe_SQ2"
               If .RecordCount > 0 Then
                   Valor = .fields("ConductaSQ2")
                   Dias_Laborados2 = .fields("SQDias2")
                   Faltas_Just2 = .fields("SQBFJ2")
                   Faltas_Injust2 = .fields("SQBFI2")
                   Atrasos2 = .fields("SQBA2")
               End If
               OpcionNotas = 2
          Case "Segundo Quimestre Tercer Parcial"
               SQLTAI = "SQTAI3"
               SQLAIC = "SQAIC3"
               SQLAGC = "SQAGC3"
               SQLL = "SQL3"
               SQLExaP = "SQExaP3"
               SQLNotas = "SQBim3"
                
               SQLBim1 = "SQBim1"
               SQLBim2 = "SQBim2"
               SQLBim3 = "SQBim3"
               SQLProm = "SQBim3"
               SQLExamen = "ExamenSQ"
               SQLPromQ = "PromSQ"
               SQLConductaQ = "ConductaSQ3"
               SQLInforme = "Informe_SQ3"
               If .RecordCount > 0 Then
                   Valor = .fields("ConductaSQ3")
                   Dias_Laborados3 = .fields("SQDias3")
                   Faltas_Just3 = .fields("SQBFJ3")
                   Faltas_Injust3 = .fields("SQBFI3")
                   Atrasos3 = .fields("SQBA3")
               End If
               OpcionNotas = 3
          Case "Promedio Segundo Quimestre", _
               "Promedio Segundo Trimestre"
               SQLBim1 = "SQBim1"
               SQLBim2 = "SQBim2"
               SQLBim3 = "SQBim3"
               SQLProm = "SQBim3"
               SQLPromQ = "PromSQ"
               SQLExamen = "ExamenSQ"
               SQLQPX = "SQ_PP"
               SQLQEX = "SQ_PE"
               SQLInforme = "Informe_SQ"
               If .RecordCount > 0 Then
                   Valor = .fields("ConductaSQ3")
                  
                   Dias_Laborados1 = .fields("SQDias1")
                   Faltas_Just1 = .fields("SQBFJ1")
                   Faltas_Injust1 = .fields("SQBFI1")
                   Atrasos1 = .fields("SQBA1")
                   
                   Dias_Laborados2 = .fields("SQDias2")
                   Faltas_Just2 = .fields("SQBFJ2")
                   Faltas_Injust2 = .fields("SQBFI2")
                   Atrasos2 = .fields("SQBA2")
                   
                   Dias_Laborados3 = .fields("SQDias3")
                   Faltas_Just3 = .fields("SQBFJ3")
                   Faltas_Injust3 = .fields("SQBFI3")
                   Atrasos3 = .fields("SQBA3")
                   
                   Dias_Laborados = .fields("SQDias3") + .fields("SQDias3") + .fields("SQDias3")
                   Faltas_Just = .fields("SQBFJ1") + .fields("SQBFJ2") + .fields("SQBFJ3")
                   Faltas_Injust = .fields("SQBFI1") + .fields("SQBFI2") + .fields("SQBFI3")
                   Atrasos = .fields("SQBA1") + .fields("SQBA2") + .fields("SQBA3")
               End If
               OpcionNotas = 4
          Case "Tercer Trimestre Primer Periodo"
               SQLBim1 = "TQBim1"
               SQLBim2 = "TQBim2"
               SQLBim3 = "TQBim3"
               SQLExamen = "ExamenTQ"
               SQLPromQ = "PromTQ"
               SQLConductaQ = "ConductaTQ1"
               If .RecordCount > 0 Then
                   Faltas_Just = .fields("PQBFJ1") + .fields("PQBFJ2") + .fields("PQBFJ3") + .fields("SQBFJ1") + .fields("SQBFJ2") + .fields("SQBFJ3")
                   Faltas_Injust = .fields("PQBFI1") + .fields("PQBFI2") + .fields("PQBFI3") + .fields("SQBFI1") + .fields("SQBFI2") + .fields("SQBFI3")
                   Atrasos = .fields("PQBA1") + .fields("PQBA2") + .fields("PQBA3") + .fields("SQBA1") + .fields("SQBA2") + .fields("SQBA3")
               End If
               OpcionNotas = 1
          Case "Tercer Trimestre Segundo Periodo"
               SQLBim1 = "TQBim1"
               SQLBim2 = "TQBim2"
               SQLBim3 = "TQBim3"
               SQLExamen = "ExamenTQ"
               SQLPromQ = "PromTQ"
               SQLConductaQ = "ConductaTQ2"
               If .RecordCount > 0 Then
                   Faltas_Just = .fields("PQBFJ1") + .fields("PQBFJ2") + .fields("PQBFJ3") + .fields("SQBFJ1") + .fields("SQBFJ2") + .fields("SQBFJ3")
                   Faltas_Injust = .fields("PQBFI1") + .fields("PQBFI2") + .fields("PQBFI3") + .fields("SQBFI1") + .fields("SQBFI2") + .fields("SQBFI3")
                   Atrasos = .fields("PQBA1") + .fields("PQBA2") + .fields("PQBA3") + .fields("SQBA1") + .fields("SQBA2") + .fields("SQBA3")
               End If
               OpcionNotas = 2
          Case "Tercer Trimestre Tercer Periodo"
               SQLBim1 = "TQBim1"
               SQLBim2 = "TQBim2"
               SQLBim3 = "TQBim3"
               SQLExamen = "ExamenTQ"
               SQLPromQ = "PromTQ"
               SQLConductaQ = "ConductaTQ3"
               If .RecordCount > 0 Then
                   Faltas_Just = .fields("PQBFJ1") + .fields("PQBFJ2") + .fields("PQBFJ3") + .fields("SQBFJ1") + .fields("SQBFJ2") + .fields("SQBFJ3")
                   Faltas_Injust = .fields("PQBFI1") + .fields("PQBFI2") + .fields("PQBFI3") + .fields("SQBFI1") + .fields("SQBFI2") + .fields("SQBFI3")
                   Atrasos = .fields("PQBA1") + .fields("PQBA2") + .fields("PQBA3") + .fields("SQBA1") + .fields("SQBA2") + .fields("SQBA3")
               End If
               OpcionNotas = 3
          Case "Promedio Tercer Trimestre"
               SQLBim1 = "TQBim1"
               SQLBim2 = "TQBim2"
               SQLBim3 = "TQBim3"
               SQLExamen = "ExamenTQ"
               SQLPromQ = "PromTQ"
               If .RecordCount > 0 Then
                   Faltas_Just = .fields("PQBFJ1") + .fields("PQBFJ2") + .fields("PQBFJ3") + .fields("SQBFJ1") + .fields("SQBFJ2") + .fields("SQBFJ3")
                   Faltas_Injust = .fields("PQBFI1") + .fields("PQBFI2") + .fields("PQBFI3") + .fields("SQBFI1") + .fields("SQBFI2") + .fields("SQBFI3")
                   Atrasos = .fields("PQBA1") + .fields("PQBA2") + .fields("PQBA3") + .fields("SQBA1") + .fields("SQBA2") + .fields("SQBA3")
               End If
               OpcionNotas = 4
          Case "Todos los Periodos", _
               "Todos Trimestres", _
               "Todos los Quimestres"
               'Presenta todos los promedios finales
               If .RecordCount > 0 Then
                   Faltas_Just = .fields("PQBFJ1") + .fields("PQBFJ2") + .fields("PQBFJ3") + .fields("SQBFJ1") + .fields("SQBFJ2") + .fields("SQBFJ3")
                   Faltas_Injust = .fields("PQBFI1") + .fields("PQBFI2") + .fields("PQBFI3") + .fields("SQBFI1") + .fields("SQBFI2") + .fields("SQBFI3")
                   Atrasos = .fields("PQBA1") + .fields("PQBA2") + .fields("PQBA3") + .fields("SQBA1") + .fields("SQBA2") + .fields("SQBA3")
               End If
               OpcionNotas = 5
        End Select
    End With
    AdoEstudiante.Close
    RatonNormal
    Seleccionar_Periodo = OpcionNotas
End Function

Public Function Equivalencia(Nota As Currency, _
                             Optional Texto_Completo As Boolean, _
                             Optional Texto_Comportamiento As Boolean, _
                             Optional Significado_Letras2 As Boolean, _
                             Optional Significado_Evaluacion2 As Boolean) As String
Dim StrEqu As String
Dim MiNota As Currency
Dim IdEquiv As Byte
   StrEqu = " "
   MiNota = Redondear(Nota, 2)
   If MiNota > 0 Then
      For IdEquiv = 0 To UBound(Equivalencias) - 1
          With Equivalencias(IdEquiv)
            If (.Desde <= MiNota) And (MiNota <= .Hasta) Then
               If Texto_Completo Then
                  StrEqu = .Significado_Equivalencia
               ElseIf Texto_Comportamiento Then
                  StrEqu = .Significado_Evaluacion
               ElseIf Significado_Letras2 Then
                  StrEqu = .Significado_Letras2
               ElseIf Significado_Evaluacion2 Then
                  StrEqu = .Significado_Evaluacion2
               Else
                  StrEqu = .Equivalencia
               End If
            End If
          End With
      Next IdEquiv
   Else
      StrEqu = " "
   End If
   If MiNota = 0 And SinImprimir Then StrEqu = ""
   Equivalencia = StrEqu
End Function

Public Sub Consultamos_Disciplinas(AdoDisciplina As Adodc, Optional Curso As String)
  sSQL = "SELECT CM.Grupo_No, CM.Codigo, " _
       & "TA.ConductaPQ1, TA.ConductaPQ2, TA.ConductaPQ3, " _
       & "TA.ConductaSQ1, TA.ConductaSQ2, TA.ConductaSQ3 " _
       & "FROM Clientes_Matriculas As CM, Trans_Asistencia As TA " _
       & "WHERE CM.Item = '" & NumEmpresa & "' " _
       & "AND CM.Periodo = '" & Periodo_Contable & "' "
  If Curso <> "" Then sSQL = sSQL & "AND CM.Grupo_No = '" & Curso & "' "
  sSQL = sSQL _
       & "AND CM.Item = TA.Item " _
       & "AND CM.Periodo = TA.Periodo " _
       & "AND CM.Codigo = TA.Codigo " _
       & "ORDER BY CM.Codigo "
  Select_Adodc AdoDisciplina, sSQL
End Sub

