MsgBox "Выбирете текстовый файл", vbOkOnly+vbInformation  +524288, "Открытие файла"
Set wShell=CreateObject("WScript.Shell")
Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
sFileSelected = oExec.StdOut.ReadLine

  Set fso = CreateObject("Scripting.FileSystemObject") 						'Создаем объект доступа к файловой системе компьютера.
  
  Set file = fso.OpenTextFile(sFileSelected, 1) 							'Открываем файл на чтение
  OldNotes = file.ReadLine 													'Читаем файл
  file.Close 																'Закрываем файл
 
  sFileSelected = Left(sFileSelected,Len(sFileSelected)-3) & "gcode"  		'Переименовываем имя файла на *.gcode
  Set file = fso.OpenTextFile(sFileSelected, 2, True) 						'Открываем файл на запись, если нету - создаем
  
  Pos1 = 1 																	'Начальная позиция
	 do																		'Начало цикла
		 Pos2 = InStr(Pos1, OldNotes, "(") 									'Ищем открытие скобки
		 tmp1 = Mid(OldNotes, Pos1, Pos2-Pos1) 								'Берем текст от начальной позиции до открытой скобки
		 Pos1 = Pos2 														'Сдвигаем начальную позицию
		 Pos2 = InStr(Pos1, OldNotes, ")")									'Ищем закрытие скобки
		 tmp2 = Mid(OldNotes, Pos1+1, Pos2-Pos1-1)							'Берем текст от начальной позиции до закрытой скобки
		 Pos1 = Pos2+2 														'Сдвигаем начальную позицию
		    
		 if tmp1 = "P" then 												'Если нота P
		  	file.WriteLine("G4 P" &  Delay(tmp2))							'Делаем запись паузы
		 else
		 	file.WriteLine("M300 P" & Delay(tmp2) & " S" & ToMHz(tmp1)) 	'Делаем запись ноты
		 end if
	 loop while Pos2<len(OldNotes)-1 										'Повторяем до конца файла
	 
  file.Close																'Закрываем файл
  
  MsgBox "Создан файл: " & sFileSelected, vbOkOnly+vbInformation  +524288, "Сохранение файла"

Function Delay (a)															'Функция определения длительности ноты
   Select Case a
   		Case "1/1."
		Delay = 2400
   		Case "1/1"
		Delay = 1600
		Case "1/2."
		Delay = 1200
		Case "1/2"
		Delay = 800
		Case "1/4."
		Delay = 600
		Case "1/4"
		Delay = 400
		Case "1/8."
		Delay = 300
   		Case "1/8"
		Delay = 200
		Case "1/16."
		Delay = 150
		Case "1/16"
		Delay = 100
		Case "1/32."
		Delay = 75
		Case "1/32"
		Delay = 50
   	End Select
End Function
  
Function ToMHz (a)															'Функция определения частоты ноты
   Select Case a
		Case "C8"
		MHz = 4186
		Case "B7"
		MHz = 3951
		Case "Ais7"
		MHz = 3729
		Case "A7"
		MHz = 3520
		Case "Gis7"
		MHz = 3322
		Case "G7"
		MHz = 3136
		Case "Fis7"
		MHz = 2960
		Case "F7"
		MHz = 2794
		Case "E7"
		MHz = 2637
		Case "Dis7"
		MHz = 2489
		Case "D7"
		MHz = 2349
		Case "Cis7"
		MHz = 2217
		Case "C7"
		MHz = 2093
		Case "B6"
		MHz = 1976
		Case "Ais6"
		MHz = 1865
		Case "A6"
		MHz = 1760
		Case "Gis6"
		MHz = 1661
		Case "G6"
		MHz = 1568
		Case "Fis6"
		MHz = 1480
		Case "F6"
		MHz = 1397
		Case "E6"
		MHz = 1319
		Case "Dis6"
		MHz = 1245
		Case "D6"
		MHz = 1175
		Case "Cis6"
		MHz = 1109
		Case "C6"
		MHz = 1047
		Case "B5"
		MHz = 988
		Case "Ais5"
		MHz = 932
		Case "A5"
		MHz = 880
		Case "Gis5"
		MHz = 831
		Case "G5"
		MHz = 784
		Case "Fis5"
		MHz = 740
		Case "F5"
		MHz = 698
		Case "E5"
		MHz = 659
		Case "Dis5"
		MHz = 622
		Case "D5"
		MHz = 587
		Case "Cis5"
		MHz = 554
		Case "C5"
		MHz = 523
		Case "B4"
		MHz = 494
		Case "Ais4"
		MHz = 466
		Case "A4"
		MHz = 440
		Case "Gis4"
		MHz = 415
		Case "G4"
		MHz = 392
		Case "Fis4"
		MHz = 370
		Case "F4"
		MHz = 349
		Case "E4"
		MHz = 330
		Case "Dis4"
		MHz = 311
		Case "D4"
		MHz = 294
		Case "Cis4"
		MHz = 277
		Case "C4"
		MHz = 262
		Case "B3"
		MHz = 247
		Case "Ais3"
		MHz = 233
		Case "A3"
		MHz = 220
		Case "Gis3"
		MHz = 208
		Case "G3"
		MHz = 196
		Case "Fis3"
		MHz = 185
		Case "F3"
		MHz = 175
		Case "E3"
		MHz = 165
		Case "Dis3"
		MHz = 156
		Case "D3"
		MHz = 147
		Case "Cis3"
		MHz = 139
		Case "C3"
		MHz = 131
		Case "B2"
		MHz = 123
		Case "Ais2"
		MHz = 117
		Case "A2"
		MHz = 110
		Case "Gis2"
		MHz = 104
		Case "G2"
		MHz = 98
		Case "Fis2"
		MHz = 92
		Case "F2"
		MHz = 87
		Case "E2"
		MHz = 82
		Case "Dis2"
		MHz = 78
		Case "D2"
		MHz = 73
		Case "Cis2"
		MHz = 69
		Case "C2"
		MHz = 65
		Case "B1"
		MHz = 62
		Case "Ais1"
		MHz = 58
		Case "A1"
		MHz = 55
		Case "Gis1"
		MHz = 52
		Case "G1"
		MHz = 49
		Case "Fis1"
		MHz = 46
		Case "F1"
		MHz = 44
		Case "E1"
		MHz = 41
		Case "Dis1"
		MHz = 39
		Case "D1"
		MHz = 37
		Case "Cis1"
		MHz = 35
		Case "C1"
		MHz = 33
		Case "B0"
		MHz = 31
		Case "Ais0"
		MHz = 29
		Case "A0"
		MHz = 28
	End Select
	
ToMHz = MHz
End Function