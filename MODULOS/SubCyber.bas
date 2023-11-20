Attribute VB_Name = "SubCyber"
Option Explicit

Global TiempoPCs(8) As String
Global TotalPCs(8) As Currency
Global ICabina As Integer

Public Sub Calcular_Total_PC(IPC As Integer, TiempoInicial As Single, TiempoFinal As Single)
Dim PC_Minutos As Integer
Dim PC_Horas As Integer
Dim PC_Minimo As Single
Dim PC_Valor_Hora As Single
Dim I As Integer

   TiempoPCs(IPC) = CDate(TiempoFinal - TiempoInicial)
   PC_Minutos = Minute(TiempoPCs(IPC))
   PC_Horas = Hour(TiempoPCs(IPC))
   PC_Minimo = VCyber_Tiempo(0).Valor
   PC_Valor_Hora = VCyber_Tiempo(Cantidad_Cyber_Tiempo).Valor
   For I = 0 To Cantidad_Cyber_Tiempo
       If (VCyber_Tiempo(I).Desde < PC_Minutos) And (PC_Minutos <= VCyber_Tiempo(I).Hasta) Then TotalPCs(IPC) = VCyber_Tiempo(I).Valor
   Next I
 'Calculamos Valor final de hora a pagar
  'MsgBox Format(TotalPCs(IPC) + (PC_Horas * PC_Valor_Hora), "#,##0.00")
  TotalPCs(IPC) = Format(TotalPCs(IPC) + (PC_Horas * PC_Valor_Hora), "#,##0.00")
  If TotalPCs(IPC) < PC_Minimo Then TotalPCs(IPC) = PC_Minimo
''''  MsgBox "PC:    " & IPC & vbCrLf _
''''       & "Hor:   " & PC_Horas & vbCrLf _
''''       & "Min:   " & PC_Minutos & vbCrLf _
''''       & "Total: " & TotalPCs(IPC) & vbCrLf _
''''       & "T.Ini: " & TiempoInicial
End Sub
