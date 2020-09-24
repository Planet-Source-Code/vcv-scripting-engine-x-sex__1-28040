VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSSE_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private scripts()        As clsSSE_Script
Private script_count     As Integer

'* Global variables
Private variables()     As typVariable
Private varCount        As Integer
Private Enum enumVarType
    xString
    xInteger
    xReal
End Enum
Private Type typVariable
    Name    As String
    value   As String
    type    As enumVarType
End Type

Public cChildAlias  As New clsSSE_Alias
Function function_calc(paramlist())
    If UBound(paramlist) < 1 Then function_calc = -1: Exit Function
    
    Dim strEval As String
    strEval = Eval(JoinArrayV(paramlist, " ", 2))
    
    function_calc = strEval
End Function

Function function_dec(paramlist())
    If UBound(paramlist) < 1 Then function_dec = -1: Exit Function
    
    'MsgBox Val(paramlist(1)) & "~" & Val(paramlist(2))
    If UBound(paramlist) >= 2 Then
        function_dec = Val(paramlist(1)) - Val(paramlist(2))
    Else
        function_dec = Val(paramlist(1)) - 1
    End If
End Function

Function function_doevents(paramlist())
    DoEvents
End Function

Function function_eval(paramlist())
    If UBound(paramlist) < 1 Then function_eval = -1: Exit Function
    
    Dim strEval As String
    strEval = Eval(JoinArrayV(paramlist, " ", 2))
    
    function_eval = strEval
End Function

Function function_inc(paramlist())
    If UBound(paramlist) < 1 Then function_inc = -1: Exit Function
    
    If UBound(paramlist) >= 2 Then
        function_inc = Val(paramlist(1)) + Val(paramlist(2))
    Else
        function_inc = Val(paramlist(1)) + 1
    End If
End Function

Function function_instr(paramlist())
    If UBound(paramlist) < 2 Then function_instr = -1: Exit Function
    
    If UBound(paramlist) = 2 Then
        function_instr = InStr(1, paramlist(1), paramlist(2))
    Else
        function_instr = InStr(Val(paramlist(3)), paramlist(1), paramlist(2))
    End If
End Function

Function function_left(paramlist())
    If UBound(paramlist) < 2 Then function_left = -1: Exit Function
    
    If paramlist(1) = "" Or Val(paramlist(2)) < 0 Then Exit Function
    function_left = Left(paramlist(1), Val(paramlist(2)))
End Function
Function function_len(paramlist())
    If UBound(paramlist) < 1 Then function_len = -1: Exit Function
    
    function_len = Len(paramlist(1))
End Function

Function function_mid(paramlist())
    If UBound(paramlist) < 2 Then function_mid = -1: Exit Function
    
    If UBound(paramlist) = 2 Then
        function_mid = Mid(paramlist(1), Val(paramlist(2)))
    Else
        function_mid = Mid(paramlist(1), Val(paramlist(2)), Val(paramlist(3)))
    End If
End Function

Function function_right(paramlist())
    If UBound(paramlist) < 2 Then function_right = -1: Exit Function
    
    If paramlist(1) = "" Or Val(paramlist(2)) < 0 Then Exit Function
    function_right = Right(paramlist(1), Val(paramlist(2)))
    
End Function


Function function_set(paramlist())
    If UBound(paramlist) < 2 Then function_set = -1: Exit Function
    
    'MsgBox paramlist(1) & "~fucker"
    If paramlist(1) Like "*:*" Then
        cChildAlias.SetLocalVar CStr(Mid(paramlist(1), 1, InStr(paramlist(1), ":") - 1)), JoinArrayV(paramlist, " ", 3), Mid(paramlist(1), InStr(paramlist(1), ":") + 1)
    End If
    
    cChildAlias.SetLocalVar CStr(paramlist(1)), JoinArrayV(paramlist, " ", 3)
End Function

Function function_ticks(paramlist())
    
    function_ticks = Timer
    
End Function

Public Function GetGlobalVar(strName As String) As String
    Dim i As Integer
    For i = 1 To varCount
        If variables(i).Name = strName Then
            GetLocalVar = variables(i).value
            Exit Function
        End If
    Next i
    GetLocalVar = ""
End Function


Private Sub AddGlobalVar(strName As String, strValue As String, Optional tvType As Integer = 0)
    varCount = varCount + 1
    ReDim Preserve variables(1 To varCount) As typVariable
    variables(varCount).Name = strName
    variables(varCount).value = strValue
    variables(varCount).type = tvType
End Sub
Private Sub SetGlobalVar(strName As String, strValue As String)
    Dim i As Integer
    
    For i = 1 To varCount
        If variables(i).Name = strName Then
            variables(i).value = strValue
            Exit Sub
        End If
    Next i
    AddGlobalVar strName, strValue
End Sub
Public Sub dev_evalscript(script_id As Integer)
    
    scripts(script_id).dev_evalaliases
    
End Sub


    


Public Function ExecuteAlias(strAliasNameX As String, strParamList())

    On Error GoTo NotBuiltIn
    Dim returnValue
    returnValue = CallByName(Me, "function_" & strAliasNameX, VbMethod, strParamList)
    ExecuteAlias = CStr(returnValue)
    Exit Function

NotBuiltIn:
    Dim i As Integer, j As Integer, returnVal As String
    For i = 1 To script_count
        
        returnValue = scripts(i).ExecuteAlias(strAliasNameX, strParamList)
        If scripts(i).bExecuted = False Then
            '* **
            '* code to generate error??
            '* **
        
            ExecuteAlias = ""
            Exit Function
        End If
        ExecuteAlias = returnValue
    Next i
    
End Function


Function function_input(paramlist())
    If UBound(paramlist) < 1 Then function_input = -1: Exit Function
    
    If UBound(paramlist) = 1 Then
        function_input = InputBox(paramlist(1))
    ElseIf UBound(paramlist) = 2 Then
        function_input = InputBox(paramlist(1), paramlist(2))
    Else
        function_input = InputBox(paramlist(1), paramlist(2), paramlist(3))
    End If
End Function

Function function_msg(paramlist())
    If UBound(paramlist) < 1 Then function_msg = -1: Exit Function
    
    If UBound(paramlist) = 1 Then
        function_msg = MsgBox(paramlist(1))
    ElseIf UBound(paramlist) = 2 Then
        function_msg = MsgBox(paramlist(1), 0, paramlist(2))
    Else
        function_msg = MsgBox(paramlist(1), paramlist(3), paramlist(2))
    End If
    
End Function

Public Sub LoadScript(script_id As Integer, strFileName As String)
    scripts(script_id).LoadScript strFileName, Me
    
End Sub

Public Sub NewScript()
    script_count = script_count + 1
    ReDim Preserve scripts(1 To script_count) As clsSSE_Script
    Set scripts(script_count) = New clsSSE_Script
    Set scripts(script_count).parent = Me
End Sub

Public Function ScriptCount() As Integer
    ScriptCount = script_count
End Function

