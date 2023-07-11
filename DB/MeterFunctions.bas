Attribute VB_Name = "MeterFunctions"
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'��� �������� ������ � �������� ���������� ����� �������� ���� �������� �� ���� ��. � ������ ��������� ��������:
'   If NamedPipeExists = False Then GoTo con
'����� ����� ���������� ������
'� ����� ��������� �����
'   SendMsgToMeter(subjectName, level1Name, level2Name, cod, day, val)
'   ���:
'       subjectName - �������� �������� � ��������� (� ���� ������ ��� string)
'       level1Name - ���� �� ��������� �����/������/������/����. � ��������� � �������� ������ ���� ��� ����, ����� �� ���������
'       level2Name - ���� �� ��������� �����������/�������/������. � ��������� � �������� ������ ���� ��� ����, ����� �� ���������
'       cod - ���� ������������ ���� ������ �������� ����� ������� ��� �����. � �������� � ��������� ������ ���� ���� � �������� ���, ����� �� ���������
'       day - ���� �� ������� ������������ ������. ������������ � ����������
'       val - ������������ �������� (������ ���� ������, ��� ���������� ���������� �������� �� ���������)
'
'��� ��������� �������� ����� ���������� ������ � �����
'����� ������ ������ ��� ���� ��������� 1 - �� ���������, 2 - ������ ����� �� ����
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'
'1 - �� ���������
'-----------------------------------------------------------------------------------------------------------------------
'Sub WriteSubjectTest()
'Dim subjectName$, level1Name$, level2Name$, val$, day$
'
'��������� ��������
'If NamedPipeExists = False Then GoTo con
'
'subjectName = "�.�. �� ������������ ���������� �����"
'level1Name = "�����"
'level2Name = "������"
'day = "26"
'val = "123"
'
'�������� � ��������
'SendMsgToMeter subjectName:=subjectName, level1Name:=level1Name, level2Name:=level2Name, day:=day, val:=val
'
'con:
'End Sub
'-----------------------------------------------------------------------------------------------------------------------
'
'
'
'2 - ������ ����� �� ���� � �����
'-----------------------------------------------------------------------------------------------------------------------
'Sub WritePlansTest()
'Dim i%, r As Range, cod$, val$, day%, c%, j%, aci%, lsi%
'
'��������� ��������
'If NamedPipeExists = False Then GoTo con
'
'day = 26
'For i = 1 To 763
'    Set r = Range(Cells(i, 27), Cells(i, 27))
'    If Not r.value = "" Then
'        cod = r.value
'        val = r.Offset(0, -1).value
'
'        �������� � ��������
'        SendMsgToMeter cod:=cod, day:=day, val:=val
'    End If
'Next
'
'
'con:
'End Sub
'-----------------------------------------------------------------------------------------------------------------------


Function NamedPipeExists() As Boolean
    Dim pipePath
    pipePath = "\\.\pipe\MeterServer"

    On Error GoTo err
    Set file = CreateObject("Scripting.FileSystemObject").CreateTextFile(pipePath)
    file.WriteLine "check"
    file.Close
    NamedPipeExists = True
    GoTo con

    err:
    MsgBox "Meter not opened!"
    NamedPipeExists = False
    con:
End Function

Sub SendMsgToMeter(Optional subjectName As String = "", Optional level1Name As String = "", Optional level2Name As String = "", Optional cod As String = "", Optional day As String = "", Optional val As String = "")
    Dim fso As Object
    Dim file As Object
    Dim pipePath As String
    Dim i%, msg$

    Dim json As Object
    Set json = CreateObject("Scripting.Dictionary")

    json.Add "subjectName", subjectName
    json.Add "level1Name", level1Name
    json.Add "level2Name", level2Name
    json.Add "cod", cod
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

Function ToJson(ByVal dict As Object) As String
    Dim key As Variant, result As String, value As String

    result = "{"
    For Each key In dict.Keys
        result = result & call IIf(Len(result) > 1, ",", "")

        If TypeName(dict(key)) = "Dictionary" Then
            value = ToJson(dict(key))
            ToJson = value
        Else
            value = Chr(34) & dict(key) & Chr(34)
        End If

        result = result & Chr(34) & key & Chr(34) & ":" & value & ""
    Next key
    result = result & "}"

    ToJson = result
End Function

Function IIf(expr, truepart, falsepart)
IIf = falsepart
if expr then IIf = truepart
End Function
