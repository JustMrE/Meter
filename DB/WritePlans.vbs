Option Explicit
Dim subjectName, level1Name, level2Name, val, day
'Выполняем проверку
'If NamedPipeExists() = False Then 
'    Call EndWriting()
'end if   
'MsgBox "ok!" 
subjectName = "с.н. ПС Алматинского энергоузла Общее"
level1Name = "прием"
level2Name = "ручное"
day = "26"
val = "123"
'Передаем в счетчики
call SendMsgToMeter(subjectName, level1Name, level2Name, day, val)
call EndWriting()
'err:

'Function NamedPipeExists() 
'    Dim pipePath
'    MsgBox "Checking!"
'    pipePath = "\\.\pipe\MeterServer"
'    On Error GoTo Next
'    Set file = CreateObject("Scripting.FileSystemObject").CreateTextFile(pipePath)
'    file.WriteLine "check"
'    file.Close
'    return True
'    err:
'    MsgBox "Meter not opened!"
'    return False
'End Function 

Sub SendMsgToMeter(subjectName, level1Name, level2Name, day, val)
    Dim fso
    Dim file
    Dim pipePath
    Dim i, msg

    Dim json
    Set json = CreateObject("Scripting.Dictionary")

    json.Add "subjectName", subjectName
    json.Add "level1Name", level1Name
    json.Add "level2Name", level2Name
    json.Add "cod", ""
    json.Add "day", day
    json.Add "value", val
    msg = ToJson(json)

    i = 0
    pipePath = "\\.\pipe\MeterServer"

    Do
    i = i + 1
        On Error Resume Next
        Set file = CreateObject("Scripting.FileSystemObject").CreateTextFile(pipePath)
        On Error GoTo 0
    If i > 50 Then Exit Sub
    Loop Until Not file Is Nothing

    file.WriteLine msg
    file.Close
End Sub

Function ToJson(ByVal dict)
    Dim key, result, value

    result = "{"
    For Each key In dict.Keys
        result = result & IIf(Len(result) > 1, ",", "")

        If TypeName(dict(key)) = "Dictionary" Then
            value = ToJson(dict(key))
            ToJson = value
        Else
            value = Chr(34) & dict(key) & Chr(34)
        End If

        result = result & Chr(34) & key & Chr(34) & ":" & value & ""
    Next
    result = result & "}"

    ToJson = result
End Function

Sub EndWriting()
    WScript.Quit
End sub

Function IIf(expr, truepart, falsepart)
IIf = falsepart
if expr then IIf = truepart
End Function