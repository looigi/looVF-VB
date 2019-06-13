Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.IO

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://looVF.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class looVF
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaSuccessivoMultimedia(idTipologia As String, Categoria As String) As String
		Dim Db As New GestioneDB
		Dim Ritorno As String = ""
		Dim Sql As String

		If Db.LeggeImpostazioniDiBase() = True Then
			Dim ConnSQL As Object = Db.ApreDB()
			Dim Rec As Object = CreateObject("ADODB.Recordset")
			Dim idCategoria As String = ""

			If Categoria <> "" And Categoria <> "Tutto" Then
				Sql = "Select * From Categorie Where idTipologia=" & idTipologia & " And Categoria='" & Categoria & "'"
				Rec = Db.LeggeQuery(ConnSQL, Sql)
				If Not Rec.Eof Then
					idCategoria = Rec("idCategoria").Value
				Else
					Return "ERROR: Categoria non trovata"
				End If
				Rec.Close
			End If

			Sql = "Select Count(*) From Dati Where idTipologia=" & idTipologia
			If Categoria <> "" And Categoria <> "Tutto" Then
				Sql &= " And idCategoria=" & idCategoria
			End If
			Rec = Db.LeggeQuery(ConnSQL, Sql)
			Dim Quante As Long = Rec(0).Value
			Rec.Close

			Static x As Random = New Random()
			Dim y As Long = x.Next(Quante)
			Dim Inizio As Long = 0

			If Categoria <> "" And Categoria <> "Tutto" Then
				Sql = "Select Min(Progressivo) From Dati Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria
				Rec = Db.LeggeQuery(ConnSQL, Sql)
				If Rec(0).Value Is DBNull.Value Then
					Inizio = 0
				Else
					Inizio = Rec(0).Value - 1
				End If
				Rec.Close
			End If

			Ritorno = (Inizio + y).ToString
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaCategorie() As String
		Dim Db As New GestioneDB
		Dim Ritorno As String = ""
		Dim Sql As String

		If Db.LeggeImpostazioniDiBase() = True Then
			Dim ConnSQL As Object = Db.ApreDB()
			Dim Rec As Object = CreateObject("ADODB.Recordset")

			Sql = "Select * From Categorie"
			Rec = Db.LeggeQuery(ConnSQL, Sql)
			Do Until Rec.Eof
				Ritorno &= Rec("idCategoria").Value & ";" & Rec("idTipologia").Value & ";" & Rec("Categoria").Value.ToString.Replace(";", "***PV***") & ";" & Rec("Protetta").Value & ";§"

				Rec.MoveNext
			Loop
			Rec.Close()
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaPermessi(Device As String, User As String, IMEI As String, IMSI As String) As String
		Dim Db As New GestioneDB
		Dim Ritorno As String = ""
		Dim Sql As String

		If Db.LeggeImpostazioniDiBase() = True Then
			Dim ConnSQL As Object = Db.ApreDB()
			Dim Rec As Object = CreateObject("ADODB.Recordset")
			Dim sDevice As String = Device
			Dim sUser As String = User

			sDevice = sDevice.Replace("***AND***", "&")
			sDevice = sDevice.Replace("***PI***", "?")

			sUser = sUser.Replace("***AND***", "&")
			sUser = sUser.Replace("***PI***", "?")

			Sql = "Select * From Permessi Where Device='" & sDevice & "' And Utente='" & sUser & "'" '  And IMEI='" & IMEI & "' And IMSI='" & IMSI & "'"
			Rec = Db.LeggeQuery(ConnSQL, Sql)
			If Not Rec.Eof Then
				Ritorno = Rec("Amministratore").Value

				Rec.Close()
			Else
				Dim id As Integer

				Sql = "Select Max(idUtente)+1 From Permessi"
				Rec = Db.LeggeQuery(ConnSQL, Sql)
				If Rec(0).Value Is DBNull.Value Then
					id = 1
				Else
					id = Rec(0).Value
				End If
				Rec.Close

				Sql = "Insert Into Permessi Values (" & id & ", '" & sDevice & "', '" & sUser & "', '" & IMEI & "', '" & IMSI & "', 'N')"
				Db.EsegueSql(ConnSQL, Sql)

				Ritorno = "N"
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaQuantiFiles(idTipologia As String) As String
		Dim Db As New GestioneDB
		Dim Quanti As String = ""
		Dim Sql As String

		If Db.LeggeImpostazioniDiBase() = True Then
			Dim ConnSQL As Object = Db.ApreDB()
			Dim Rec As Object = CreateObject("ADODB.Recordset")

			Sql = "Select Count(*) From Dati Where idTipologia=" & idTipologia
			Rec = Db.LeggeQuery(ConnSQL, Sql)
			Quanti = Rec(0).Value
			Rec.Close()
		End If

		Return Quanti
	End Function

	<WebMethod()>
	Public Function RitornaMultimediaDaId(idTipologia As String, idCategoria As String, idMultimedia As String) As String
		Dim Db As New GestioneDB
		Dim Ritorno As String = ""
		Dim Sql As String

		If Db.LeggeImpostazioniDiBase() = True Then
			Dim ConnSQL As Object = Db.ApreDB()
			Dim Rec As Object = CreateObject("ADODB.Recordset")

			Sql = "Select Dati.NomeFile, Dati.Dimensioni, Dati.Data, Dati.idCategoria, Categorie.Categoria, Categorie.Percorso, Categorie.LetteraDisco From Dati " &
				"Left Join Categorie On Dati.idCategoria=Categorie.idCategoria And Dati.idTipologia=Categorie.idTipologia " &
				"Where Dati.idTipologia=" & idTipologia & " And Dati.idCategoria=" & idCategoria & " And Progressivo=" & idMultimedia
			Rec = Db.LeggeQuery(ConnSQL, Sql)
			If Not Rec.Eof Then
				Dim Thumb As String = ""

				If idTipologia = "2" Then
					Dim Conversione As String = "" & Rec("LetteraDisco").Value.ToString
					Dim PathOriginale As String = Rec("Percorso").Value.ToString
					If Conversione <> "" Then
						Dim cc() As String = Conversione.Split("*")
						PathOriginale = PathOriginale.Replace(cc(0), cc(1))
					End If
					If Right(PathOriginale, 1) <> "\" Then
						PathOriginale &= "\"
					End If

					' Return "" & Rec("Categoria").Value & " - " & PathOriginale & " - " & Rec("NomeFile").value

					Thumb = CreaThumbDaVideo("" & Rec("Categoria").Value, PathOriginale, PathOriginale & Rec("NomeFile").value)
				End If

				Ritorno = Thumb & "§" & Rec("NomeFile").Value.ToString.Replace(";", "***PV***") & ";" & Rec("Dimensioni").Value & ";" & Rec("Data").Value & ";" & Rec("idCategoria").Value & ";" & idMultimedia.ToString & ";"
			Else
				Ritorno = "ERROR: Nessun file rilevato"
			End If
			Rec.Close()
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaImmaginiPerGriglia(QuanteImm As String, Categoria As String) As String
		Dim Db As New GestioneDB
		Dim Ritorno As String = ""
		Dim Sql As String

		If Db.LeggeImpostazioniDiBase() = True Then
			Dim ConnSQL As Object = Db.ApreDB()
			Dim Rec As Object = CreateObject("ADODB.Recordset")

			Dim Inizio As Long = 0
			Dim idCategoria As String = ""

			If Categoria <> "" And Categoria <> "Tutto" Then
				Sql = "Select * From Categorie Where idTipologia=1 And Categoria='" & Categoria & "'"
				Rec = Db.LeggeQuery(ConnSQL, Sql)
				If Not Rec.Eof Then
					idCategoria = Rec("idCategoria").Value
				Else
					Return "ERROR: Categoria non trovata"
				End If
				Rec.Close
			End If

			Sql = "Select Count(*) From Dati Where idTipologia=1"
			If Categoria <> "" And Categoria <> "Tutto" Then
				Sql &= " And idCategoria=" & idCategoria
			End If
			Rec = Db.LeggeQuery(ConnSQL, Sql)
			Dim Quante As Long = Rec(0).Value
			Rec.Close

			If Categoria <> "" And Categoria <> "Tutto" Then
				Sql = "Select Min(Progressivo) From Dati Where idTipologia=1 And idCategoria=" & idCategoria
				Rec = Db.LeggeQuery(ConnSQL, Sql)
				Inizio = Rec(0).Value - 1
				Rec.Close
			End If

			For i As Integer = 1 To Val(QuanteImm)
				Randomize()
				Static x As Random = New Random()
				Dim y As Long = x.Next(Quante)
				Dim idMultimedia As Long = Inizio + y

				Sql = "Select Dati.NomeFile, Dati.Dimensioni, Dati.Data, Dati.idCategoria, Categorie.Categoria, Categorie.Percorso From Dati " &
					"Left Join Categorie On Dati.idCategoria=Categorie.idCategoria And Dati.idTipologia=Categorie.idTipologia " &
					"Where Dati.idTipologia=1 And Progressivo=" & idMultimedia
				Rec = Db.LeggeQuery(ConnSQL, Sql)
				If Not Rec.Eof Then
					Ritorno &= Rec("NomeFile").Value.ToString.Replace(";", "***PV***") & ";" & Rec("Dimensioni").Value & ";" & Rec("Data").Value & ";" & Rec("idCategoria").Value & ";" & idMultimedia.ToString & ";§"
				Else
					Ritorno = "ERROR: Nessun file rilevato"
					Exit For
				End If
				Rec.Close()
			Next
		End If

		Return Ritorno
	End Function

	Private Function CreaThumbDaVideo(Categoria As String, Percorso As String, Video As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim PathBase As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PercorsoThumbs.txt")

		If Strings.Right(PathBase, 1) <> "\" Then
			PathBase &= "\"
		End If

		Dim ritorno As String = "111->" & PathBase & "*" & Categoria & "*" & Percorso & "*" & Video & "*" & vbCrLf

		Dim OutPut As String = PathBase & Categoria & "\"
		Dim Nome As String = gf.TornaNomeFileDaPath(Video)
		'Dim Nome As String = Video
		Dim Estensione As String = gf.TornaEstensioneFileDaPath(Nome)

		Dim Cartella As String = Video.Replace(Nome, "")

		ritorno &= "222->" & OutPut & "*" & Nome & "*" & Estensione & "*" & Video & "*" & Cartella & "*" & vbCrLf

		' OutPut &= Cartella

		gf.CreaDirectoryDaPercorso(OutPut)
		OutPut &= Nome.Replace(Estensione, "") & ".jpg"
		' Video = Percorso & "\" & Video

		If Not File.Exists(OutPut) Then
			Dim processoFFMpeg As Process = New Process()
			Dim pi As ProcessStartInfo = New ProcessStartInfo()
			pi.Arguments = "-i """ & Video & """ -vframes 1 -an -s 1024x768 -ss 5 """ & OutPut & """"

			' Return pi.Arguments

			pi.FileName = Server.MapPath(".") & "\ffmpeg.exe"
			' gf.CreaAggiornaFile(Server.MapPath(".") & "\Buttami.txt", pi.Arguments)
			pi.WindowStyle = ProcessWindowStyle.Normal
			processoFFMpeg.StartInfo = pi
			processoFFMpeg.Start()
			processoFFMpeg.WaitForExit()
		End If

		Return OutPut.Replace(PathBase, "")
	End Function

	<WebMethod()>
	Public Function RitornaFiles() As String
		Dim gf As New GestioneFilesDirectory

		'If Aggiorna = "N" Then
		'	Dim Ok As Boolean = False

		'	gf.ScansionaDirectorySingola(Server.MapPath(".") & "\Temp\")
		'	Dim Filetti() As String = gf.RitornaFilesRilevati
		'	Dim qFiles As Integer = gf.RitornaQuantiFilesRilevati
		'	Dim MaxDatella As Date = "01/01/1970 00:00:00"
		'	For i As Integer = 1 To qFiles
		'		Dim Nome As String = gf.TornaNomeFileDaPath(Filetti(i))
		'		Dim Campi() As String = Nome.Split("_")
		'		Nome = Campi(0) & "/" & Campi(1) & "/" & Campi(2) & " " & Campi(3) & ":" & Campi(4) & ":" & Campi(5)
		'		Dim Datella As Date = Nome.Replace(".txt", "")
		'		If Datella > MaxDatella Then
		'			MaxDatella = Datella
		'			Ok = True
		'		End If
		'	Next

		'	If Ok Then
		'		Dim RitornoAggiorna As String = Server.MapPath(".") & "\Temp\" & MaxDatella.ToString.Replace(":", "_").Replace(" ", "_").Replace("\", "_").Replace("/", "_") & ".txt"

		'		Return RitornoAggiorna
		'	End If
		'End If

		Dim sPathsVideo As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PercorsiVideo.txt")
		Dim sPathsImm As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PercorsiImmagini.txt")
		Dim PathVideo() As String = sPathsVideo.Split("§")
		Dim PathImmagini() As String = sPathsImm.Split("§")
		'Dim FilesVideo As List(Of StrutturaFiles) = New List(Of StrutturaFiles)
		'Dim FilesImmagini As List(Of StrutturaFiles) = New List(Of StrutturaFiles)
		Dim Conta As Long
		Dim Db As New GestioneDB
		Dim Sql As String

		If Db.LeggeImpostazioniDiBase() = True Then
			Dim ConnSQL As Object = Db.ApreDB()
			Dim idCategoria As Integer = 0

			Db.EsegueSql(ConnSQL, "Delete From Categorie")
			Db.EsegueSql(ConnSQL, "Delete From Dati")

			Conta = 0
			For Each p As String In PathVideo
				If p.Trim <> "" Then
					Dim pp() As String = p.Split(";")

					idCategoria += 1
					Sql = "Insert Into Categorie Values (" & idCategoria & ", 2, '" & pp(0).Replace("'", "''") & "', '" & pp(1).Replace("'", "''") & "', '" & pp(2) & "', '" & pp(3) & "')"
					Dim sRitorno As String = Db.EsegueSql(ConnSQL, Sql)

					If Strings.Right(pp(1), 1) <> "\" Then
						pp(1) &= "\"
					End If
					gf.ScansionaDirectorySingola(pp(1))
					Dim qFiles As Integer = gf.RitornaQuantiFilesRilevati
					Dim Files() As String = gf.RitornaFilesRilevati
					For i As Integer = 1 To qFiles
						'Dim sf As New StrutturaFiles

						'sf.Categoria = Conta
						'sf.NomeFile = Files(i).Replace(pp(1), "").Replace(";", "**PV***").Replace("§", "***COSO***")
						'sf.DimensioniFile = FileLen(Files(i))
						'sf.DataFile = FileDateTime(Files(i))

						'FilesVideo.Add(sf)
						Conta += 1
						Sql = "Insert Into Dati Values (" &
							" " & Conta & ", " &
							"2, " &
							" " & idCategoria & ", " &
							"'" & Files(i).Replace(pp(1), "").Replace("'", "''") & "', " &
							" " & FileLen(Files(i)) & ", " &
							"'" & FileDateTime(Files(i)) & "' " &
							")"
						sRitorno = Db.EsegueSql(ConnSQL, Sql)
						If sRitorno <> "" Then
							Return "ERROR:" & sRitorno
						End If
					Next
				End If
			Next

			Conta = 0
			idCategoria = 0
			For Each p As String In PathImmagini
				If p.Trim <> "" Then
					Dim pp() As String = p.Split(";")

					idCategoria += 1
					Sql = "Insert Into Categorie Values (" & idCategoria & ", 1, '" & pp(0).Replace("'", "''") & "', '" & pp(1).Replace("'", "''") & "', '" & pp(2) & "', '" & pp(3) & "')"
					Dim sRitorno As String = Db.EsegueSql(ConnSQL, Sql)

					If Strings.Right(pp(1), 1) <> "\" Then
						pp(1) &= "\"
					End If
					gf.ScansionaDirectorySingola(pp(1))
					Dim qFiles As Integer = gf.RitornaQuantiFilesRilevati
					Dim Files() As String = gf.RitornaFilesRilevati
					' Conta += 1
					For i As Integer = 1 To qFiles
						'Dim sf As New StrutturaFiles

						'sf.Categoria = Conta
						'sf.NomeFile = Files(i).Replace(pp(1), "").Replace(";", "**PV***").Replace("§", "***COSO***")
						'sf.DimensioniFile = FileLen(Files(i))
						'sf.DataFile = FileDateTime(Files(i))

						'FilesImmagini.Add(sf)
						Conta += 1
						Sql = "Insert Into Dati Values (" &
							" " & Conta & ", " &
							"1, " &
							" " & idCategoria & ", " &
							"'" & Files(i).Replace(pp(1), "").Replace("'", "''") & "', " &
							" " & FileLen(Files(i)) & ", " &
							"'" & FileDateTime(Files(i)) & "' " &
							")"
						sRitorno = Db.EsegueSql(ConnSQL, Sql)
						If sRitorno <> "" Then
							Return "ERROR: " & sRitorno
						End If
					Next
				End If
			Next

			'gf.CreaDirectoryDaPercorso(Server.MapPath(".") & "\Temp\")

			'Dim NomeFile As String = Server.MapPath(".") & "\Temp\" & Now.ToString.Replace(":", "_").Replace(" ", "_").Replace("\", "_").Replace("/", "_") & ".txt"
			'gf.ApreFileDiTestoPerScrittura(NomeFile)

			'For Each p As String In PathVideo
			'	If p.Trim <> "" Then
			'		Dim pp() As String = p.Split(";")
			'		Dim Stringa As String = "CategoriaVideo;" & pp(0).Replace(";", "**PV***").Replace("§", "***COSO***") & ";§"
			'		gf.ScriveTestoSuFileAperto(Stringa)
			'	End If
			'Next

			'For Each p As String In PathImmagini
			'	If p.Trim <> "" Then
			'		Dim pp() As String = p.Split(";")
			'		Dim Stringa As String = "CategoriaImmagini;" & pp(0).Replace(";", "**PV***").Replace("§", "***COSO***") & ";§"
			'		gf.ScriveTestoSuFileAperto(Stringa)
			'	End If
			'Next

			'For Each sf As StrutturaFiles In FilesVideo
			'Dim Stringa As String = "Video;" & sf.Categoria & ";" & sf.NomeFile & ";" & sf.DimensioniFile.ToString & ";" & sf.DataFile & ";§"
			'gf.ScriveTestoSuFileAperto(Stringa)
			'Next

			'For Each sf As StrutturaFiles In FilesImmagini
			'Dim Stringa As String = "Pic;" & sf.Categoria & ";" & sf.NomeFile & ";" & sf.DimensioniFile.ToString & ";" & sf.DataFile & ";§"
			'gf.ScriveTestoSuFileAperto(Stringa)
			'Next
			'gf.ChiudeFileDiTestoDopoScrittura()

			' Dim Ritorno As String = gf.LeggeFileIntero(NomeFile)

			gf = Nothing

			'Dim Ritorno As String = NomeFile
		End If


		Return "*"
	End Function

	<WebMethod()>
	Public Function EffettuaRicerca(idTipologia As String, Categoria As String, Ricerca As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim Ritorno As String = ""
		Dim Db As New GestioneDB
		Dim Rec As Object = CreateObject("ADODB.Recordset")
		Dim Sql As String

		If Db.LeggeImpostazioniDiBase() = True Then
			Dim ConnSQL As Object = Db.ApreDB()

			Dim idCategoria As String = ""

			If Categoria <> "" And Categoria <> "Tutto" Then
				Sql = "Select * From Categorie Where idTipologia=" & idTipologia & " And Categoria='" & Categoria & "'"
				Rec = Db.LeggeQuery(ConnSQL, Sql)
				If Not Rec.Eof Then
					idCategoria = Rec("idCategoria").Value
				Else
					Return "ERROR: Categoria non trovata"
				End If
				Rec.Close
			End If

			Sql = "Select Top 30 Dati.Progressivo, Dati.NomeFile, Dati.Dimensioni, Dati.Data, Dati.idCategoria, Categorie.Categoria, Categorie.Percorso From Dati " &
				"Left Join Categorie On Dati.idCategoria=Categorie.idCategoria And Dati.idTipologia=Categorie.idTipologia " &
				"Where Dati.idTipologia=" & idTipologia & " And Dati.NomeFile Like '%" & Ricerca & "%' And Dati.idCategoria = " & idCategoria
			Rec = Db.LeggeQuery(ConnSQL, Sql)
			If Not Rec.eof Then
				Do Until Rec.eof
					Dim Thumb As String = ""

					If idTipologia = "2" Then
						Thumb = CreaThumbDaVideo(Rec("Categoria").Value, Rec("Percorso").Value, Rec("NomeFile").value)
					End If

					Ritorno &= Thumb & ";" & Rec("NomeFile").Value.ToString.Replace(";", "***PV***") & ";" & Rec("Dimensioni").Value & ";" & Rec("Data").Value & ";" & Rec("idCategoria").Value & ";" & Rec("Progressivo").Value.ToString & ";" & idTipologia & ";§"

					Rec.MoveNext
				Loop

				Rec.Close()
			Else
				Ritorno = "ERROR: Nessun file rilevato"
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function TornaNumeroImmaginePerSfondo() As String
		Dim gf As New GestioneFilesDirectory
		Dim Path As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PercorsiSfondi.txt")
		Dim Immagine As String = ""

		Try
			MkDir(Server.MapPath(".") & "\Sfondi")
		Catch ex As Exception

		End Try

		If QuanteImmaginiSfondi = 0 Then
			If Strings.Right(Path, 1) <> "\" Then
				Path &= "\"
			End If
			gf.ScansionaDirectorySingola(Path)
			Dim filetti() As String = gf.RitornaFilesRilevati
			QuanteImmaginiSfondi = gf.RitornaQuantiFilesRilevati

			For i As Integer = 1 To QuanteImmaginiSfondi
				ListaImmagini.Add(filetti(i))
			Next
		End If

		Dim Minuti As Integer = (Now.Minute \ 3) * 3

		'If Minuti > 0 And Minuti < 15 Then
		'	Minuti = 0
		'Else
		'	If Minuti >= 15 And Minuti < 30 Then
		'		Minuti = 15
		'	Else
		'		If Minuti >= 30 And Minuti < 45 Then
		'			Minuti = 30
		'		Else
		'			If Minuti >= 45 Then
		'				Minuti = 45
		'			End If
		'		End If
		'	End If
		'End If

		Dim NomeFile As String = Server.MapPath(".") & "\Sfondi\Sfondo_" & Minuti & ".txt"

		If File.Exists(NomeFile) Then
			Immagine = gf.LeggeFileIntero(NomeFile)
		Else
			For Each foundFile As String In My.Computer.FileSystem.GetFiles(Server.MapPath(".") & "\Sfondi")
				Dim conta As Integer = 0

				While File.Exists(foundFile)
					File.Delete(foundFile)
					Threading.Thread.Sleep(1000)
					conta += 1
					If conta > 10 Then
						Exit While
					End If
				End While
			Next

			Static x As Random = New Random()
			Dim y As Long = x.Next(QuanteImmaginiSfondi)

			Immagine = y.ToString & ";" & ListaImmagini.Item(y).Replace(Path, "") & ";"

			gf.CreaAggiornaFile(NomeFile, Immagine)
		End If


		Return Immagine
	End Function

End Class