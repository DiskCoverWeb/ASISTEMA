VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJSONScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim dictVars As New Dictionary
Dim plNestCount As Long

Public Function Eval(sJSON As String) As String
   Dim SB As New cStringBuilder
   Dim O As Object
   Dim C As Object
   Dim I As Long
   
   Set O = JSON.parse(sJSON)
   If (JSON.GetParserErrors = "") And Not (O Is Nothing) Then
      For I = 1 To O.Count
         Select Case VarType(O.Item(I))
         Case vbNull
            SB.Append "null"
         Case vbDate
            SB.Append CStr(O.Item(I))
         Case vbString
            SB.Append CStr(O.Item(I))
         Case Else
            Set C = O.Item(I)
            SB.Append ExecCommand(C)
         End Select
      Next
   Else
      MsgBox JSON.GetParserErrors, vbExclamation, "Parser Error"
   End If
   Eval = SB.toString
End Function

Public Function ExecCommand(ByRef obj As Variant) As String
   Dim SB As New cStringBuilder
   
   If plNestCount > 40 Then
      ExecCommand = "ERROR: Nesting level exceeded."
   Else
      plNestCount = plNestCount + 1
      
      Select Case VarType(obj)
         Case vbNull
            SB.Append "null"
         Case vbDate
            SB.Append CStr(obj)
         Case vbString
            SB.Append CStr(obj)
         Case vbObject
            
            Dim I As Long
            Dim J As Long
            Dim this As Object
            Dim key
            Dim paramKeys
            
            If TypeName(obj) = "Dictionary" Then
               Dim sOut As String
               Dim sRet As String
   
               Dim keys
               keys = obj.keys
               For I = 0 To obj.Count - 1
                  sRet = ""
             
                  key = keys(I)
                  If VarType(obj.Item(key)) = vbString Then
                     sRet = obj.Item(key)
                  Else
                     Set this = obj.Item(key)
                  End If
                  
                  ' command implementation
                  Select Case LCase(key)
                  Case "alert":
                     MsgBox ExecCommand(this.Item("message")), vbInformation, ExecCommand(this.Item("title"))
                     
                  Case "input":
                     SB.Append InputBox(ExecCommand(this.Item("prompt")), ExecCommand(this.Item("title")), ExecCommand(this.Item("default")))
                     
                  Case "switch"
                     sOut = ExecCommand(this.Item("default"))
                     sRet = LCase(ExecCommand(this.Item("case")))
                     For J = 0 To this.Item("items").Count - 1
                        If LCase(this.Item("items").Item(J + 1).Item("case")) = sRet Then
                           sOut = ExecCommand(this.Item("items").Item(J + 1).Item("return"))
                           Exit For
                        End If
                     Next
                     SB.Append sOut
                  
                  Case "set":
                     If dictVars.Exists(this.Item("name")) Then
                        dictVars.Item(this.Item("name")) = ExecCommand(this.Item("value"))
                     Else
                        dictVars.Add this.Item("name"), ExecCommand(this.Item("value"))
                     End If
                     
                  Case "get":
                     sRet = ExecCommand(dictVars(CStr(this.Item("name"))))
                     If sRet = "" Then
                        sRet = ExecCommand(this.Item("default"))
                     End If
                     
                     SB.Append sRet
                     
                  Case "if"
                     Dim val1 As String
                     Dim val2 As String
                     Dim bRes As Boolean
                     val1 = ExecCommand(this.Item("value1"))
                     val2 = ExecCommand(this.Item("value2"))
                     
                     bRes = False
                     Select Case LCase(this.Item("type"))
                     Case "eq" ' =
                        If LCase(val1) = LCase(val2) Then
                           bRes = True
                        End If
                        
                     Case "gt" ' >
                        If val1 > val2 Then
                           bRes = True
                        End If
                     
                     Case "lt" ' <
                        If val1 < val2 Then
                           bRes = True
                        End If
                     
                     Case "gte" ' >=
                        If val1 >= val2 Then
                           bRes = True
                        End If
                     
                     Case "lte" ' <=
                        If val1 <= val2 Then
                           bRes = True
                        End If
                     
                     End Select
                     
                     If bRes Then
                        SB.Append ExecCommand(this.Item("true"))
                     Else
                        SB.Append ExecCommand(this.Item("false"))
                     End If
                     
                  Case "return"
                     SB.Append obj.Item(key)
                  
                     
                  Case Else
                     If TypeName(this) = "Dictionary" Then
                        paramKeys = this.keys
                        For J = 0 To this.Count - 1
                           If J > 0 Then
                              sRet = sRet & ","
                           End If
                           sRet = sRet & CStr(this.Item(paramKeys(J)))
                        Next
                     End If
                     
                     
                     SB.Append "<%" & UCase(key) & "(" & sRet & ")%>"
                     
                  End Select
               Next I
               
            ElseIf TypeName(obj) = "Collection" Then
   
               Dim value
               For Each value In obj
                  SB.Append ExecCommand(value)
               Next value
               
            End If
            Set this = Nothing
   
         Case vbBoolean
            If obj Then SB.Append "true" Else SB.Append "false"
         
         Case vbVariant, vbArray, vbArray + vbVariant
         
         Case Else
            SB.Append Replace(obj, ",", ".")
      End Select
      plNestCount = plNestCount - 1
   End If
   
   ExecCommand = SB.toString
   Set SB = Nothing
   
End Function




