Attribute VB_Name = "modSEX_CommonFunctions"
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal mWnd As Long, ByVal aWnd As Long, data As String, parms As String, show As Boolean, nopause As Boolean) As Long

Function CallDLL(strLibraryName As String, functionName As String)
    Dim lb As Long, pa As Long
    'map 'user32' into the address space of the calling process.
    lb = LoadLibrary(strLibraryName)
    pa = GetProcAddress(lb, functionName)
    'CallWindowProc pa, Me.hWnd,
    
End Function

Function FileExists(strFileName As String) As Boolean
    On Error GoTo MakeF
    'If file does Not exist, there will be an Error
    Open strFileName For Input As #1
    Close #1
    'no error, file exists
    FileExists = True
    Exit Function
MakeF:
    'error, file does Not exist
    FileExists = False
    Exit Function

End Function


Function JoinArray(thearray() As String, strDelim As String, start As Integer, Optional endx As Integer = -1) As String
    If endx = -1 Then endx = UBound(thearray) + 1
    Dim i As Integer, result As String
    
    For i = start - 1 To endx - 1
        If i = endx - 1 Then
            result = result & thearray(i)
        Else
            result = result & thearray(i) & strDelim
        End If
    Next i
    JoinArray = result
End Function

Function JoinArrayV(thearray(), strDelim As String, start As Integer, Optional endx As Integer = -1) As String
    If endx = -1 Then endx = UBound(thearray) + 1
    Dim i As Integer, result As String
    
    For i = start - 1 To endx - 1
        If i = endx - 1 Then
            result = result & thearray(i)
        Else
            result = result & thearray(i) & strDelim
        End If
    Next i
    JoinArrayV = result
End Function

Function TrimLeft(strText As String) As String
    Dim i As Integer
    For i = 1 To Len(strText)
        If Mid(strText, i, 1) <> " " And Mid(strText, i, 1) <> Chr(9) Then
            TrimLeft = Right(strText, Len(strText) - (i - 1))
            Exit Function
        End If
    Next i
End Function


