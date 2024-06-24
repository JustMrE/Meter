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
'subjectName - название субъекта в счетчиках (должно совпадать точь в точь)
'level1Name - прием\отдача\сальдо в счетчиках у субъекта должно быть данное поле
'level2Name - ручное\оперативное\счетчик в счетчиках у субъекта subjectName в поле level1Name должно быть данное поле
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
    ' Получаем путь к временной папке
    TempPath = Environ("TEMP")
    
    ' Создаем уникальное имя файла во временной папке
    filePath = TempPath & filename
    
    
    Set stream = CreateObject("ADODB.Stream")
    
    ' Настраиваем поток
    With stream
        .Type = 2 ' Устанавливаем тип потока как текст
        .Charset = "utf-8" ' Устанавливаем кодировку как UTF-8
        .Open
        .WriteText jsonData
        .SaveToFile filePath, 2 ' Сохраняем в файл, перезаписывая его если он существует
        .Close
    End With
    
    ' Освобождаем объект
    Set stream = Nothing
    
'    ' Открываем файл для записи
'    FileNumber = FreeFile
'    Open filePath For Output As #FileNumber
'
'    ' Записываем данные в файл
'    Print #FileNumber, jsonData
'
'    ' Закрываем файл
'    Close #FileNumber
    
    ' Перемещаем файл в другую директорию
    Name filePath As DestinationPath & filename
    
'    MsgBox "Файл успешно записан и перемещен в D:\"
    
    jsonData = ""
End Function

Sub ЗаписьКириллическихСимволов()
    Dim filePath As String
    Dim textToWrite As String
    Dim stream As Object
    
    ' Указываем путь к файлу
    filePath = "C:\path\to\your\file.txt"
    
    ' Текст для записи
    textToWrite = "Пример текста с кириллическими символами: Привет, мир!"
    
    ' Создаем объект ADODB.Stream
    Set stream = CreateObject("ADODB.Stream")
    
    ' Настраиваем поток
    With stream
        .Type = 2 ' Устанавливаем тип потока как текст
        .Charset = "utf-8" ' Устанавливаем кодировку как UTF-8
        .Open
        .WriteText textToWrite
        .SaveToFile filePath, 2 ' Сохраняем в файл, перезаписывая его если он существует
        .Close
    End With
    
    ' Освобождаем объект
    Set stream = Nothing
    
    MsgBox "Текст успешно записан в файл!"
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

