Option Explicit
	Public mdict, buf, tidict, tidict1, buf2(500,2), kod, vOIK(500, 25), vrArr(100, 2)
		Set mdict = CreateObject("Scripting.Dictionary") 'задание словаря месяцев
			mdict.CompareMode = 1
			'Dim marr1, marr2 
			'marr1 = Array("января" , "февраля" , "марта" , "апреля" , "мая" , "июня" , "июля" , "августа" , "сентября" , "октября" , "ноября" , "декабря")
			'marr2 = Array("01" , "02" , "03" , "04" , "05" , "06" , "07" , "08" , "09" , "10" , "11" , "12")
			'For i = 1 To 12
			'	mdict.Add marr1(i - 1), marr2(i - 1)
			'Next
			mdict.Add "января","01"
			mdict.Add "февраля","02"
			mdict.Add "марта","03"
			mdict.Add "апреля","04"
			mdict.Add "мая","05"
			mdict.Add "июня","06"
			mdict.Add "июля","07"
			mdict.Add "августа","08"
			mdict.Add "сентября","09"
			mdict.Add "октября","10"
			mdict.Add "ноября","11"
			mdict.Add "декабря","12"
	'_____________________________________________________________________________________________________________________________________________________________
		Set tidict = CreateObject("Scripting.Dictionary")
		Set tidict1 = CreateObject("Scripting.Dictionary")
	Public  i, md1, ExcelApp,ExcelApp1,ExcelApp2,ExcelApp3, filName(3), ZZcount, ii, jj, ss, razn, KODTI, sheetName
	public flag
	dim if1, if2, if3, if4, if5
	dim usl1, b1, b2
	flag = false
	b1 = false
	b2 = false
	usl1 = false
	Dim dd, kk,ttimer, ttimer1, otv, FSO, F, path, pass
	
	ttimer = Timer
	
	Set ExcelApp3 = GetObject(,"Excel.Application")
	If ExcelApp3.WorkBooks.Count > 2 Then
		ExcelApp3.ActiveWorkbook.Close
	Else
		ExcelApp3.Quit
	End If
	
	Set FSO = CreateObject("Scripting.FileSystemObject")
	F = FSO.GetFile(Wscript.ScriptFullName)
	'path = FSO.GetParentFolderName(F)
	path = "D:\NURAKHMANOV\Desktop\2020"

	MsgBox "начата запись плана!"
	
	Set ExcelApp = CreateObject("Excel.Application")
	Set ExcelApp1 = CreateObject("Excel.Application")
	ExcelApp.WorkBooks.Open path & "\89.xls"
	ExcelApp.Visible = true
	ExcelApp1.WorkBooks.Open path & "\Счетчики 2020.xlsm"
	ExcelApp1.DisplayAlerts = false
	
	'Открываем кодти и переносим в массив диапазон A1:C100
	Set ExcelApp2 = CreateObject("Excel.Application")
		ExcelApp2.WorkBooks.Open path & "\Коды ТИ.xlsx"
		ExcelApp2.Windows("Коды ТИ.xlsx").Activate
		ExcelApp2.Sheets(1).Select
		ExcelApp2.Visible=false
		KODTI = ExcelApp2.Sheets(1).Range("A1:C500") 'элементы в кодти заносим в массив
		ExcelApp2.ActiveWorkbook.Close
		
	'For i = 1 To  UBound(KODTI)
	'	If KODTI(i, 2) <> "" Then
	'		tidict.Add KODTI(i, 2), 0
	'		if KODTI(i, 3)<>"" Then
	'			tidict1.Add KODTI(i, 2), KODTI(i, 3)
	'		end if
	'	End If
	'Next
	i = 1
	dim l
	l = 0
	While (i <= UBound(KODTI) And KODTI(i, 1) <> "")
		If KODTI(i, 2) <> "" Then
			tidict.Add KODTI(i, 2), 0
			
			vrArr(l, 1) = KODTI(i, 2)
			
			l = l + 1
			if KODTI(i, 3)<>"" Then
				tidict1.Add KODTI(i, 2), KODTI(i, 3)
			end if
		End If
		i = i + 1
	Wend
	
	ExcelApp.Windows("89.xls").Activate
	If ExcelApp.ActiveSheet.Cells(1, "A").Value = "" then
		dd = 1
	else
		dd = ExcelApp.ActiveSheet.Cells(1, "A").Value
		if Len(dd) > 2 then
			pass = left(dd, 4)
			dd = right(dd, len(dd) - 4)
			if pass = "2562" then
				flag = true
			end if
		end if
	end if
	
	For i = 1 To dd
		if (dd >= 2) then 'Or dd = 3) Then
			ExcelApp.Sheets(i).Activate
			sheetName = ExcelApp.ActiveSheet.Name
		else 
			sheetName = ExcelApp.ActiveSheet.Name
		end if
		
		if not(i = 1) then erase buf
		Erase vOIK
		
		Call OpredData()
		
		buf = ExcelApp.Sheets(sheetName).Range("B4:AA700")
		l = 0
		For ii = 1 To UBound(buf)
			If tidict.Exists(buf(ii,26)) Then
				tidict.Item(buf(ii,26)) = buf(ii,25)
			End If
			If tidict1.Exists(buf(ii,26)) Then
				for jj = 1 to 24
					vOIK(ii, jj) = buf(ii, jj)
				next
				vOIK(ii, 25) = tidict1.Item(buf(ii,26))
			End If
		Next
		
		
		For l = 0 to UBound(vrArr)
			If tidict.Exists(vrArr(l, 1)) Then
				vrArr(l, 2) = tidict.item(vrArr(l, 1))
			end if
		next
		
		ExcelApp1.Windows("Счетчики 2020.xlsm").Activate
		ExcelApp1.Sheets("ТЭПм").Select
		ExcelApp1.Visible = false
		jj = 6
		
		Call ProvDataZapisi()
		
		if flag = false then
			if (razn < 0) then b1 = true
			if (razn > -5) then b2 = true
			usl1 = b1 and b2
		else
			if (razn < 0) then b1 = true
			usl1 = b1
		end if
		
		'режим администратора
		if razn >= 2 and flag = true then
			razn = 1
		end if
		
		ExcelApp1.run "WriteToLog", "usl1 = " & usl1
		
		Select Case true 'razn
			Case (razn = 0) 'повторная запись на последнее число
				if flag = true then
					otv = 6
				else
					otv = MsgBox ("Производится повторная запись на " & md1, vbYesNo)
				end if
				If otv = 6 Then
					ExcelApp1.run "WriteToLogEmty"
					ExcelApp1.run "WritePlansToLog", "Производится повторная запись на " & md1
					ExcelApp1.run "wrPlan", vrArr, CStr(md1)
				End If
			Case (razn = 1) 'первичная запись на последнее число
				
				ExcelApp1.run "WriteToLogEmty"
				ExcelApp1.run "WritePlansToLog", "Производится первичная запись на " & md1
				ExcelApp1.run "wrPlan", vrArr, CStr(md1)
			Case (razn >= 2) ' 2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50
				MsgBox "План должен записываться на " & CDate(ExcelApp1.Sheets("ТЭПм").Cells(7, 1).Value) + 1 & ", а Вы пытаетесь ввести данные на " & md1
				Call endThisMacro()
			Case (usl1) '-1, -2, -3, -4,  'повторная запись на предыдущие дни
				if flag = true then
					otv = 6
				else
					otv = MsgBox ("Производится повторная запись на " & md1, vbYesNo)
				end if
				If otv = 6 Then
					ExcelApp1.run "WriteToLogEmty"
					ExcelApp1.run "WritePlansToLog", "Производится повторная запись на " & md1
					ExcelApp1.run "wrPlan", vrArr, CStr(md1)
				End If
			Case (razn = -999)
				MsgBox "Превышен лимит плановых записей без факта, не более пяти"
				Call endThisMacro()
			End Select
	Next
	ExcelApp1.Sheets("ПС").Select
	MsgBox Timer - ttimer & vbCrLf & "Готово!"
	call endThisMacro()

	
Public Sub OpredData() 
	Dim tekmon, d, g, mes, wd1, k, WshShell
	k = 6
	Do While k = 6
		ExcelApp.Windows("89.xls").Activate

		tekmon = ExcelApp.Sheets(sheetName).Cells(1, 13).Value
		d = ExcelApp.Sheets(sheetName).Cells(1, 12).Value
		g = ExcelApp.Sheets(sheetName).Cells(1, 15).Value
		
		if d<10 then d="0" & d 
		If mdict.Exists(tekmon) Then
			mes = mdict.Item(tekmon)
			k = 7
			md1 = d & "." & mes & "." & Mid(g,3) ' dd.mm.yy
			'filName(i) = d & "-" & mes & "-" & Mid(g,3) & "_план.xls"' dd.mm.yy
		Else
			k = MsgBox ("Неправильно записано наименование месяца в таблице " & ExcelApp.Sheets(sheetName).Name & ". Исправьте и нажмите да. При нажатии нет запись не будет произведена" ,vbYesNo + vbSystemModal)
			Set WshShell = WScript.CreateObject("WScript.Shell")
			WshShell.SendKeys "{ENTER}"
			Set WshShell = Nothing
			ExcelApp.Sheets(sheetName).Cells(1, 1).Activate
			if k = 7 Then Call endThisMacro()
		End If
	Loop
End Sub

Public Sub ProvDataZapisi()
	Dim i, mm
	i=0 
	mm = ""
	' считаем количество записей на листе "тэп" без факта
	if flag = false then
		Do While mm =""
			i=i+1
			mm = ExcelApp1.Sheets("ТЭПм").Cells(6 + i, 3).Value
			if i > 50 then
				exit do
			end if
		Loop
		i = i - 1
	end if
	razn = 0
	if i > 5 Then 
		razn = -999
	ElseIf i <= 5 Then
		razn = CDate(md1) - CDate(ExcelApp1.Sheets("ТЭПм").Cells(7, 1).Value)
		If razn = 0 Then 
			ss = 1 
			Exit Sub
		End If
		If razn = 1 Then 
			ss = 1
			Exit Sub
		End If
		If razn = 2 Then 
			Exit Sub
		End If
		If razn < 0 Then
			razn = CDate(md1) - CDate(ExcelApp1.Sheets("ТЭПм").Cells(7, 1).Value)
			ss = 1 - razn
		End If
	End If	
End Sub

sub endThisMacro()
		If Not(ExcelApp.ActiveWorkbook Is Nothing) Then ExcelApp.ActiveWorkbook.Close True
			ExcelApp1.DisplayAlerts = False
			If Not(ExcelApp1.ActiveWorkbook Is Nothing) Then ExcelApp1.ActiveWorkbook.Close True
			ExcelApp1.DisplayAlerts = True
		If Not(ExcelApp2.ActiveWorkbook Is Nothing) Then ExcelApp2.ActiveWorkbook.Close
	WScript.Quit
end sub