VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MeterWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim jsonData$, dataCollection As Collection
Dim DestinationPath As String

Private Sub Class_Initialize()
'    DestinationPath = "D:\jsons\"
    DestinationPath = "X:\MeterWorker"
    Set dataCollection = New Collection
End Sub

Private Sub Class_Terminate()

End Sub
'subjectName - �������� �������� � ��������� (������ ��������� ���� � ����)
'level1Name - �����\������\������ � ��������� � �������� ������ ���� ������ ����
'level2Name - ������\�����������\������� � ��������� � �������� subjectName � ���� level1Name ������ ���� ������ ����
Public Function AddToJson(Optional subjectName As String = "", Optional level1Name As String = "", Optional level2Name As String = "", Optional cod As String = "", Optional day As String = "", Optional val As String = "", Optional size As String = "")

Dim i%
If val = "" Then val = 0
    
Dim jsonObject As Object
Set jsonObject = CreateObject("Scripting.Dictionary")
    
jsonObject.Add "subjectName", subjectName
jsonObject.Add "level1Name", level1Name
jsonObject.Add "level2Name", level2Name
jsonObject.Add "cod", cod
jsonObject.Add "day", day
jsonObject.Add "value", val

dataCollection.Add jsonObject
End Function

Public Function WriteData()
GenerateJsonData
WriteToFile
End Function

Private Function GenerateJsonData()
Dim key As Variant, result As String, value As String
Dim obj
jsonData = "["

For Each obj In dataCollection
    jsonData = jsonData & ToJson(obj) & ","
Next

jsonData = jsonData & "]"
End Function

Private Function WriteToFile()
Dim TempPath As String
Dim filePath As String
Dim FileNumber As Integer
Dim stream As Object
Dim dat$, filename$
dat = format(DateTime.Now, "ddMMyy_HHmmss")
filename = "\temp_file " & dat & ".json"
    ' �������� ���� � ��������� �����
    TempPath = Environ("TEMP")
    
    ' ������� ���������� ��� ����� �� ��������� �����
    filePath = TempPath & filename
    
    
    Set stream = CreateObject("ADODB.Stream")
    
    ' ����������� �����
    With stream
        .Type = 2 ' ������������� ��� ������ ��� �����
        .Charset = "utf-8" ' ������������� ��������� ��� UTF-8
        .Open
        .WriteText jsonData
        .SaveToFile filePath, 2 ' ��������� � ����, ������������� ��� ���� �� ����������
        .Close
    End With
    
    ' ����������� ������
    Set stream = Nothing
    
'    ' ��������� ���� ��� ������
'    FileNumber = FreeFile
'    Open filePath For Output As #FileNumber
'
'    ' ���������� ������ � ����
'    Print #FileNumber, jsonData
'
'    ' ��������� ����
'    Close #FileNumber
    
    ' ���������� ���� � ������ ����������
    Name filePath As DestinationPath & filename
    
'    MsgBox "���� ������� ������� � ��������� � D:\"
    
    jsonData = ""
End Function

Sub ���������������������������()
    Dim filePath As String
    Dim textToWrite As String
    Dim stream As Object
    
    ' ��������� ���� � �����
    filePath = "C:\path\to\your\file.txt"
    
    ' ����� ��� ������
    textToWrite = "������ ������ � �������������� ���������: ������, ���!"
    
    ' ������� ������ ADODB.Stream
    Set stream = CreateObject("ADODB.Stream")
    
    ' ����������� �����
    With stream
        .Type = 2 ' ������������� ��� ������ ��� �����
        .Charset = "utf-8" ' ������������� ��������� ��� UTF-8
        .Open
        .WriteText textToWrite
        .SaveToFile filePath, 2 ' ��������� � ����, ������������� ��� ���� �� ����������
        .Close
    End With
    
    ' ����������� ������
    Set stream = Nothing
    
    MsgBox "����� ������� ������� � ����!"
End Sub

Private Function ToJson(ByVal dict As Object) As String
    Dim key As Variant, result As String, value As String

    result = "{"
    For Each key In dict.Keys
        result = result & IIf(Len(result) > 1, ",", "")

        If TypeName(dict(key)) = "Dictionary" Then
            value = ToJson(dict(key))
            ToJson = value
        Else
'            If key = "cod" Or key = "day" Then
            value = Chr(34) & dict(key) & Chr(34)
        End If

        result = result & Chr(34) & key & Chr(34) & ":" & value & ""
    Next key
    result = result & "}"

    ToJson = result
End Function
