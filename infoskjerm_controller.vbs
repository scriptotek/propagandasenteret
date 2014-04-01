' @ rev. 8 (2014-03-11)
' <infoskjerm_controller.vbs>
' Dan Michael Heggø <d.m.heggo@ub.uio.no> (2012)
' 
' Scriptet kjører i en loop (Sub MainLoop), som sjekker en gitt mappe 
' (baseFolder) for powerpoint-filer (ppt, pps, pptx, ppsx) hvert 
' 5. sekund. Den nyeste filen startes, mens evt. eldre filer flyttes til 
' en arkivmappe (archiveFolder).
'
' Hva hvis powerpoint-filen som skal åpnes er i bruk?
'   Hvis filen er i bruk, venter scriptet til den er lukket. 
'   Det er ikke egentlig noe teknisk i veien for å starte en fil som er i bruk 
'   i read-only modus, eller lage en kopi av den, men hvis noen holder på å 
'   jobbe med filen, kan det jo godt være den ikke befinner seg i en tilstand 
'   som er egnet for fremvisning. Derfor venter vi.
'
' Hva kan gå galt?
'   Hvis Powerpoint skulle finne på å kræsje, vil scriptet starte programmet på
'   nytt, men det er da viktig at ikke en dialogboks blokkerer systemet:
'    - for å skru av "Windows is checking for a solution…", se <http://tinyurl.com/btfc6fl>
'    - for å skru av "auto recovery": File > Powerpoint options > Save og fjern
'      avkryssing for "Save autorecover information every ..."
' Låser scriptet den aktive powerpoint-filen?
'   Nei, scriptet lagrer en midlertidig kopi, som den kjører istedet for 
'   originalfilen. Denne legges i scriptFolder, skjules, og startes i read-only
'   modus (hvorfor ikke?)
'
' Hvordan avslutte scriptet?
'   Scriptet kan avsluttes ved å opprette en fil "killscript" i scriptFolder 
'   (uten filendelse). Da vil scriptet rydde opp alle midlertidige filer, lukke 
'   powerpoint og seg selv. Hvis scriptet avsluttes på andre måter, vil det 
'   uansett rydde neste gang det kjører. 
' 
Option Explicit 	' Stans også ved manglende variabeldeklarasjoner, siden 
					' dette kan innføre feil som er vanskelig å spore

Const ForReading = 1
Const ForWriting = 2
' Vi holder antallet globale variabler til et minimum:
Dim baseFolder, scriptFolder, archiveFolder, logFile, pptPattern, pptPattern2


Dim aapningstiderEnabled = False
Dim aapningstider(6,3)
For i = 0 to 6   ' alle dager
	aapningstider(i, 0) = 8    ' Åpner klokka 08
	aapningstider(i, 1) = 0    ' (minutter, men disse ignoreres)
	aapningstider(i, 2) = 22   ' Stenger klokka 22
	aapningstider(i, 3) = 0    ' (minutter, men disse ignoreres)
Next

Const SCRIPT_CLOSING = "Scriptet avslutter"
baseFolder = "C:\SHOW\"       ' mappen som skal sjekkes
'baseFolder = "M:\realfagsbiblioteket\vbs-test\"   ' for testing
scriptFolder = "script\"      ' undermappe av baseFolder
archiveFolder = "arkiv\"      ' undermappe av baseFolder
logFile = "log.txt"           ' havner i scriptFolder

' regexp for å matche powerpointfiler: skipp filer som starter med tilde, disse er temp-filer
pptPattern = "^[^~].*\.pp(t|s|tx|sx)$"
pptPattern2 = "^[^~].*\.pp(s|sx)$"  ' eksluderer show-filer

Sub LogMsg(msg)
	' Skriver ut en melding til skjerm og til loggfil
	Dim fso, objTextFile
	Const ForAppending = 8

	Set fso = CreateObject("Scripting.FileSystemObject")
		
	WScript.Echo("[" & Now() & "] " & msg)
	Set objTextFile = fso.OpenTextFile(baseFolder & scriptFolder & logFile, ForAppending, True)
	objTextFile.WriteLine("[" & Now() & "] " & msg)
	objTextFile.Close
End Sub

Sub Welcome
	' Skriver ut velkomstmelding når scriptet starter
	WScript.Echo(" ")
	WScript.Echo("__      __  _ _                                        ")
	Wait(0.1)
	WScript.Echo("\ \    / / | | |                                       ")
	Wait(0.1)
	WScript.Echo(" \ \  / /__| | | _____  _ __ ___  _ __ ___   ___ _ __  ")
	Wait(0.1)
	WScript.Echo("  \ \/ / _ \ | |/ / _ \| '_ ` _ \| '_ ` _ \ / _ \ '_ \ ")
	Wait(0.1)
	WScript.Echo("   \  /  __/ |   < (_) | | | | | | | | | | |  __/ | | |")
	Wait(0.1)
	WScript.Echo("    \/ \___|_|_|\_\___/|_| |_| |_|_| |_| |_|\___|_| |_|")
	Wait(0.1)
	WScript.Echo(" ")
	Wait(0.4)
   	LogMsg("-[ Starter ]------------------------------------------------------------------------------------")


	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")	
	If fso.FileExists(baseFolder & scriptFolder & "specialpage") = True Then
		LogMsg("Fant <specialpage>-fil. Sletter...")
		Call fso.DeleteFile(baseFolder & scriptFolder & "specialpage")
		LogMsg("------------------------------------------------------------------------------------------")
	End If
	
	Dim filepath : filepath = baseFolder & scriptFolder & "Nå vises.txt"
	Dim file : Set file = fso.OpenTextFile(filepath, ForWriting, True)
	file.WriteLine("Starter")
	file.Close
	
End Sub


Sub Wait(seconds)
	' Venter i <seconds> sekunder
	Dim i
	'Dim out
	'Set out = WScript.StdOut
	If seconds > 0.5 Then
	'out.Write("Sleeping "&seconds&" seconds")
		For i = 1 To seconds*2 Step 1
			WScript.Sleep 500
		'out.Write(".")
		Next
	'out.WriteBlankLines(1)
	Else
		WScript.Sleep seconds*1000
	End If
End Sub

Function GetNewestFile()
	' Returnerer filnavnet til den sist endrede eller opprettede 
	' filen i mappa <baseFolder> som matcher <pattern>
	Dim re, fso, file, filename, filedate, newestdate
	
	Set re = New RegExp
	re.IgnoreCase = True
	re.Pattern = pptPattern
	
	Set fso = CreateObject("Scripting.FileSystemObject")	
	For Each file in fso.GetFolder(baseFolder).Files
		If re.Test(file.Name) Then
			' Filer som er kopiert kan ha DateCreated nyere enn DateLastModified. 
			' Vi bruker den av datoene som er nyest
			If file.DateCreated > file.DateLastModified Then
				newestdate = file.DateCreated
			Else
				newestdate = file.DateLastModified
			End If
			If IsEmpty(filedate) Or newestdate > filedate Then
				filedate = newestdate
				filename = file.Name
			End If
		End If
	Next

	GetNewestFile = filename ' The VBScript-way of returning a value

End Function

Sub ArchiveOldFiles(currentFileName)
	' Arkiverer alle powerpoint-filer bortsett fra currentFileName
	Dim re, fso, file, j, newname
	
	Set re = New RegExp
	re.IgnoreCase = True
	re.Pattern = pptPattern
	
	Set fso = CreateObject("Scripting.FileSystemObject")	
	For Each file in fso.GetFolder(baseFolder).Files
		If re.Test(file.Name) And file.Name <> currentFileName Then
			newname = file.Name
			j = 1
			Do While fso.FileExists(baseFolder & archiveFolder & newname)
				newname = fso.GetBaseName(file.Name) & "_" & j & "." & fso.GetExtensionName(file.Name)
				j = j + 1
			Loop
			LogMsg("Arkiverer <" & file.Name & "> som <" & newname & ">")
			On Error Resume Next   ' Catch errors instead of exiting
			Call fso.MoveFile(baseFolder & file.Name, baseFolder & archiveFolder & newname)
			If Err.Number <> 0 Then
				LogMsg(" -> Filen er i bruk")
			End If
			Err.Clear
			On Error GoTo 0        ' Back to strict error handling
		End If
	Next
	
End Sub

Function IsWriteAccessible(sFilePath)
    ' Fra: http://stackoverflow.com/questions/12300678/how-can-i-determine-if-a-file-is-locked-using-vbs
    ' Strategy: Attempt to open the specified file in 'append' mode.
    ' Does not appear to change the 'modified' date on the file.
    ' Works with binary files as well as text files.
    IsWriteAccessible = False
	Const ForAppending = 8
    Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")
    Dim nErr : nErr = 0
    Dim sDesc : sDesc = ""

    On Error Resume Next    ' Catch errors instead of exiting

    Dim oFile : Set oFile = oFso.OpenTextFile(sFilePath, ForAppending)
    If Err.Number = 0 Then
        oFile.Close
        If Err Then
            nErr = Err.Number
            sDesc = Err.Description
        Else
            IsWriteAccessible = True
        End if
    Else
        Select Case Err.Number
            Case 70
                ' Permission denied because:
                ' - file is open by another process
                ' - read-only bit is set on file, *or*
                ' - NTFS Access Control List settings (ACLs) on file
                '   prevents access

            Case Else
                ' 52 - Bad file name or number
                ' 53 - File not found
                ' 76 - Path not found

                nErr = Err.Number
                sDesc = Err.Description
        End Select
    End If

    Set oFile = Nothing
    Set oFso = Nothing
    On Error GoTo 0    ' Back to strict error handling

    If nErr Then
        Err.Raise nErr, , sDesc
    End If
End Function

Function WaitForFileReady(filename)
	' Venter til filen er ulåst, men maks <maxwaittime> sekunder. 
	' Returnerer True hvis fila er klar til bruk, eller False ellers.
	Dim timestep : timestep = 5 ' seconds between each check
	Dim maxwaittime : maxwaittime = 60 ' seconds
	Dim waittime : waittime = 0
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	Dim fileReady : fileReady = False

	LogMsg("Sjekker om <" & filename & "> er klar")
	Do
		if fso.FileExists(baseFolder & filename) = False Then
			LogMsg("Filen eksisterer ikke lenger. Avbryter")
			Exit Do
		End If

		fileReady = IsWriteAccessible(baseFolder & filename)
		If not fileReady Then
			LogMsg("Venter på at filen skal bli klar...")
		End If
		Wait(timestep)
		waittime = waittime + timestep
		If waittime >= maxwaittime Then
			LogMsg("Har ventet " & maxwaittime & " sekunder. Sjekker på nytt")
			Exit Do
		End If
	Loop Until fileReady

	WaitForFileReady = fileReady
End Function

Sub CloseAllShows()
	' Lukker alle åpne filer, men ikke selve programmet
	Dim app, presentation
	Set app = CreateObject("PowerPoint.Application")
	app.Visible = True ' Merkelig sak, men må visst settes
	'LogMsg("Lukker alle show")
	Do
		For Each presentation in app.Presentations
			LogMsg("Lukker <" & presentation.Name & ">")
			Call presentation.Close()
		Next
	Loop Until app.Presentations.Count = 0
End Sub

Function StartShow(filename)
	' Lukker alle eksisterende show, og starter angitt show. 
	' * Vi lager en midlertidig kopi av showet, som vi starter istedet for 
	'   originalfilen, for å unngå at vi låser originalfilen.
	' * Hvis originalfilen er låst, antar vi at noen holder på å jobbe med den,
	'   og venter med å kopiere den til filen har blitt låst opp.
	' Returnerer True hvis showet kunne startes, False ellers

	Const ppAdvanceOnTime = 2   ' Run according to timings (not clicks)
	Const ppShowTypeKiosk = 3   ' Run in "Kiosk" mode (fullscreen)
	Const ppAdvanceTime = 10     ' Show each slide for 10 seconds
	
	Dim sh, app, presentation, slideShowWindow, fileReady, fso, newname
	
	If WaitForFileReady(filename) = False Then
		' Filen er fortsatt låst. Vi kan like godt ta en ny sjekk i mappen 
		' mens vi venter, i tilfelle det har dukket opp en ny fil der.
		StartShow = False ' return state
		Exit Function
	End If
	
	Call CloseAllShows()
	Call Cleanup(False)
	
	' Lag kopi av showet:
	Set fso = CreateObject("Scripting.FileSystemObject")
	newname = "_active." + fso.GetExtensionName(baseFolder & filename)
	'if fso.FileExists(baseFolder & scriptFolder & newname) Then
	'	fso.DeleteFile(baseFolder & scriptFolder & newname)
	'End If
	
	LogMsg("Kopierer <" & filename & "> til <" & newname & ">")
	On Error Resume Next   ' Catch errors instead of exiting
	Call fso.CopyFile(baseFolder & filename, baseFolder & scriptFolder & newname)	
	fso.GetFile(baseFolder & scriptFolder & newname).Attributes = 2 ' Hidden file. 
	If Err.Number <> 0 Then
		LogMsg(" -> Kopiering mislykket (#:" & Err.Number & ": " & Err.Description & ")")
		Err.Clear
		On Error GoTo 0        ' Back to strict error handling
		StartShow = False ' return state
		Exit Function
	End If
	On Error GoTo 0        ' Back to strict error handling

	' Start kopien:	
	Set app = CreateObject("PowerPoint.Application")
	app.Visible = True ' Merkelig sak, men må visst settes
	LogMsg("Starter <" & filename & "> som <" & newname & ">")
	Set sh = CreateObject("WScript.Shell")
	Call sh.AppActivate(app.Caption)
	Set presentation = app.Presentations.Open(baseFolder & scriptFolder & newname, True)

	LogMsg("Showet er lastet inn")
	
	' Apply powerpoint settings
	presentation.Slides.Range.SlideShowTransition.AdvanceOnTime = TRUE
	presentation.SlideShowSettings.AdvanceMode = ppAdvanceOnTime 
	'presentation.SlideShowSettings.ShowType = ppShowTypeKiosk
	'presentation.Slides.Range.SlideShowTransition.AdvanceTime = ppAdvanceTime
	presentation.SlideShowSettings.LoopUntilStopped = True

	'presentation.Saved = True
	'presentation.SlideShowSettings.ShowType = ppShowTypeKiosk
	Set slideShowWindow = presentation.SlideShowSettings.Run()
	slideShowWindow.Activate
	
	Call ArchiveOldFiles(filename)	
	StartShow = True ' return state
	
	'presentation.SlideShowWindow.View.PointerType = ppSlideShowPointerAlwaysHidden

	'objPresentation.Slides.Range.SlideShowTransition.AdvanceOnTime = TRUE
	'objPresentation.SlideShowSettings.AdvanceMode = ppAdvanceOnTime 

	'objPresentation.SlideShowSettings.StartingSlide = 1
	'objPresentation.SlideShowSettings.EndingSlide = objPresentation.Slides.Count
	'objPresentation.Slides.Range.SlideShowTransition.AdvanceTime = 8

	'Set objSlideShow = objPresentation.SlideShowSettings.Run.View
	'objPresentation.SlideShowWindow.View.PointerType = ppSlideShowPointerAlwaysHidden
       
	'Do Until objSlideShow.State = ppSlideShowDone
	'    If Err <> 0 Then
	'        Exit Do
	'    End If
	'Loop
	'objPresentation.Saved = True
	'objPresentation.Close
	'objPPT.Quit
	
End Function

Sub UpdateStatus(showname)
	' Lager en tekstfil som viser hvilket show som kjøres,
	' for å gi en form for tilbakemelding til brukeren
	Dim app, fso, file, showfile, filepath, objTextFile, newestdate
	Dim slideidx : slideidx = -1
	newestdate = "n/a"
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	If showname <> SCRIPT_CLOSING Then
	
		Set app = CreateObject("PowerPoint.Application")
		app.Visible = True ' Merkelig sak, men må visst settes
		If app.SlideShowWindows.Count = 1 Then
			' showname = app.SlideShowWindows(1).Presentation.Name
			' Bruker heller oppgitt navn, siden vi jobber med en kopi

			Set showfile = fso.GetFile(baseFolder & showname)
			If showfile.DateCreated > showfile.DateLastModified Then
				newestdate = showfile.DateCreated
			Else
				newestdate = showfile.DateLastModified
			End If
		
		slideidx = app.SlideShowWindows(1).View.CurrentShowPosition

		Elseif app.SlideShowWindows.Count = 0 Then
			showname = "ingen presentasjoner vises nå"
		Else
			showname = "mer enn én presentasjon (noe er galt!)"
		End If
	
	End If
		
	filepath = baseFolder & scriptFolder & "Nå vises.txt"
	Set file = fso.OpenTextFile(filepath, ForWriting, True)
	file.WriteLine(showname)
	file.WriteLine(newestdate)
	file.WriteLine(slideidx)
	file.Close

End Sub

Sub Cleanup(closePowerPoint)
	' Lukker åpne show og sletter midlertidige filer
	Dim app, re1, fso, file
	LogMsg("Rydder og sletter midlertidige filer")
	
	Set app = CreateObject("PowerPoint.Application")
	app.Visible = True
	Call CloseAllShows()
	if closePowerPoint = True Then
		Call UpdateStatus(SCRIPT_CLOSING)	
	Else
		Call UpdateStatus("")
	End If
	
	Set fso = CreateObject("Scripting.FileSystemObject")	
	If fso.FolderExists(baseFolder) = False Then
		WScript.Echo "Mappen " & baseFolder & " eksisterer ikke!"
		WScript.Quit
	End If	
	If fso.FolderExists(baseFolder & scriptFolder) = False Then
		fso.CreateFolder(baseFolder & scriptFolder)
	End If
	If fso.FolderExists(baseFolder & archiveFolder) = False Then
		fso.CreateFolder(baseFolder & archiveFolder)
	End If
	
	Set re1 = New RegExp
	re1.Pattern = "^_active.pp(t|s|tx|sx)$"
	For Each file in fso.GetFolder(baseFolder & scriptFolder).Files
		If re1.Test(file.Name) Then
			'LogMsg("Sletter <" & file.Name & ">")
			fso.DeleteFile(baseFolder & scriptFolder & file.Name)
		End If
	Next	
	
	If closePowerPoint = True Then
		Call app.Quit()
	End If
End Sub

Sub OpenIE2(url)
	Dim wshell : Set wshell = CreateObject("Wscript.Shell")
	Dim shell : Set shell = CreateObject("Shell.Application")
	Dim window, wtitle
	wshell.Run("IEXPLORE.EXE -k """ & url & """")	
	Dim winFound : winFound = False
	Do 
		For Each window in shell.Windows
			WScript.echo window.LocationUrl
			If left(window.LocationUrl, Len(url)) = url Then
				wtitle = window.Document.Title & " - " & window.Name
				WScript.Echo "Found " & wtitle
				winFound = True
			End If
		Next
		WScript.Sleep 500
	Loop While winFound = False	
End Sub

Sub CloseIE()
	Dim sa : Set sa = CreateObject("Shell.Application")
	Dim window
	Dim closed
	Do
		closed = 0
		For Each window in sa.Windows
			If window.Name = "Windows Internet Explorer" Then
				Call window.Quit()
				closed = closed + 1
			End If
		Next
	Loop While closed > 0
End Sub

Sub MainLoop
	Dim prevfilename, prevdate, currentfile, currentfilename, currentdate, fso
	Dim cday, date1, hourOpen, hourClose, isCountingDown, isSleeping, hourNow, minNow, secNow, specialPageStarted
	isCountingDown = 0
	isSleeping = 0
	specialPageStarted = False
		
	Set fso = CreateObject("Scripting.FileSystemObject")
	Do
		'LogMsg("Loop")
		If fso.FileExists(baseFolder & scriptFolder & "killscript") = True Then
			LogMsg("-----------------------------------------------------------------")
			LogMsg("Fant <killscript>-fil. Avslutter...")
			Call fso.DeleteFile(baseFolder & scriptFolder & "killscript")
			
			filepath = baseFolder & scriptFolder & "Nå vises.txt"
			Set file = fso.OpenTextFile(filepath, ForWriting, True)
			file.WriteLine("Lukker Powerpoint")
			file.Close

			Exit Do
		End If
		
		If specialPageStarted Then
			' Pass
		Else

			If fso.FileExists(baseFolder & scriptFolder & "specialpage") = True Then
				LogMsg("-----------------------------------------------------------------")
				LogMsg("Fant <specialpage>-fil. Avslutter...")
				specialPageStarted = True
				
				Dim objTextFile : Set objTextFile = fso.OpenTextFile(baseFolder & scriptFolder & "specialpage", ForReading, True)
				Dim url : url = objTextFile.Readline
				objTextFile.Close

				Call OpenIE2(url)
			End If
			
			
			date1 = Now()
			'date1 = CDate("19.11.2012 21:48:41")  ' for å teste ulike datoer
			hourNow = Hour(date1)
			minNow = Minute(date1)
			secNow = Second(date1)

			cday = WeekDay(date1) - 2
			If cday = -1 Then
				cday = 6
			End If
			' 0-4: ukedager, 5-6: lørdag, søndag

			hourOpen = aapningstider(cday, 0)
			hourClose = aapningstider(cday, 2)
			' TODO: Legge til støtte for minutter også :)

			If (aapningstiderEnabled) And ((hourClose-hourNow <= 0) Or (hourNow-hourOpen < -1) Or ((hourNow-hourOpen = -1) And (60-minNow > 10))) Then	' 10 min før åpning starter vi igjen
				' Biblioteket er stengt

				If isSleeping = 0 Then
					LogMsg("Går i sovemodus")
					isCountingDown = 0
					isSleeping = 1
					Call CloseIE()
					Wait(1.0)
					Call OpenIE2("http://biblionaut.net/propaganda/natta/")
				End If

			Elseif (aapningstiderEnabled) And ((hourClose-hourNow = 1) And (60-minNow < 20)) Then	' 20 minutter til stenging: nedtelling
				' Vi nærmer oss stengetid
				
				If isCountingDown = 0 Then
					LogMsg("Starter nedtelling til stenging")
					isSleeping = 0
					isCountingDown = 1
					Call CloseIE()
					Call OpenIE2("http://biblionaut.net/propaganda/stenging/")
					'Wait(5)
					'Call CloseIE()
				End If
			
			Else

				If isSleeping = 1 Then
					LogMsg("God morgen!")
					LogMsg("Åpningstider i dag: " + hourOpen + "-" + hourClose)
					isSleeping = 0
					isCountingDown = 0
					'Call CloseIE()
					Exit Do			' strengt tatt ikke nødvendig, men kanskje like greit å starte på nytt en gang i døgnet uansett?
				End If
							
			
			currentfilename = GetNewestFile()
			If IsEmpty(currentfilename) Then
				LogMsg("Fant ingen powerpoints i " & baseFolder)
			Else
				Set currentfile = fso.getFile(baseFolder + currentfilename)
				If IsEmpty(prevfilename) Then
					If StartShow(currentfile.Name) = True Then
						prevfilename = currentfile.Name
						prevdate = currentfile.DateLastModified
					End If
				Elseif prevfilename <> currentfile.Name Then
					LogMsg("-----------------------------------------------------------------")
					LogMsg("Fant nytt show: <" & currentfile.Name & ">")
					If StartShow(currentfile.Name) = True Then
						prevfilename = currentfile.Name
						prevdate = currentfile.DateLastModified
					End If
				Elseif prevdate <> currentfile.DateLastModified Then
					LogMsg("-----------------------------------------------------------------")
					LogMsg("Showet har blitt endret: <" & currentfile.Name & ">")
					If StartShow(currentfile.Name) = True Then
						prevfilename = currentfile.Name
						prevdate = currentfile.DateLastModified
					End If
				End If
			End If
			Call UpdateStatus(currentfilename)
			End If
		End If
		Wait(5)
	Loop
End Sub

Call CloseIE()
Call Welcome()
Call Cleanup(False)
Call MainLoop()
Call CloseIE()
Call Cleanup(True)
