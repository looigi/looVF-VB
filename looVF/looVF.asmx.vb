﻿Imports System.Web.Services
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
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Sql As String

		'Dim gf As New GestioneFilesDirectory
		'Dim NomeFile As String = Server.MapPath(".") & "\Log\LogRitorno.txt"
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim idCategoria As String = ""

			'gf.ScriveTestoSuFileAperto(NomeFile, idTipologia & "-" & Categoria)

			If Categoria <> "" And Categoria <> "Tutto" Then
				Sql = "Select * From Categorie Where idTipologia=" & idTipologia & " And Categoria='" & Categoria & "'"
				'gf.ScriveTestoSuFileAperto(NomeFile, Sql)
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				If Rec.Eof = False Then
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
			'gf.ScriveTestoSuFileAperto(NomeFile, Sql)
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			Dim Quante As Long = Rec(0).Value
			Rec.Close
			'gf.ScriveTestoSuFileAperto(NomeFile, Quante)

			Static x As Random = New Random()
			Dim y As Long = x.Next(Quante)
			Dim Inizio As Long = 0

			If Categoria <> "" And Categoria <> "Tutto" Then
				Sql = "Select Min(Progressivo) From Dati Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria
				'gf.ScriveTestoSuFileAperto(NomeFile, Sql)
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				If Rec(0).Value Is DBNull.Value Then
					Inizio = 0
				Else
					Inizio = Rec(0).Value - 1
				End If
				Rec.Close
			Else
				idCategoria = -1
			End If

			Ritorno = (Inizio + y).ToString & ";" & idCategoria
		End If
		'gf.ScriveTestoSuFileAperto(NomeFile, Ritorno)

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaCategoriaDaID(id As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Sql As String

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object

			Sql = "Select B.idCategoria, B.Categoria From Dati A " &
				"Left Join Categorie B On A.idCategoria = B.idCategoria And A.idTipologia = B.idTipologia " &
				"Where Progressivo=" & id
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			If Rec.Eof = False Then
				Ritorno &= Rec("idCategoria").Value & ";" & Rec("Categoria").Value
			End If
			Rec.Close()
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaCategorie() As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Sql As String

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object

			Sql = "Select * From Categorie"
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
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
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Sql As String

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim sDevice As String = Device
			Dim sUser As String = User

			sDevice = sDevice.Replace("***And***", "&")
			sDevice = sDevice.Replace("***PI***", "?")

			sUser = sUser.Replace("***And***", "&")
			sUser = sUser.Replace("***PI***", "?")

			Sql = "Select * From Permessi Where Device='" & sDevice & "' And Utente='" & sUser & "'" '  And IMEI='" & IMEI & "' And IMSI='" & IMSI & "'"
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			If Rec.Eof = False Then
				Ritorno = Rec("Amministratore").Value

				Rec.Close()
			Else
				Dim id As Integer

				Sql = "Select Max(idUtente)+1 From Permessi"
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				If Rec(0).Value Is DBNull.Value Then
					id = 1
				Else
					id = Rec(0).Value
				End If
				Rec.Close

				Sql = "Insert Into permessi Values (" & id & ", '" & sDevice & "', '" & sUser & "', '" & IMEI & "', '" & IMSI & "', 'N')"
				Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

				Ritorno = "N"
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaQuantiFiles(idTipologia As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Quanti As String = ""
		Dim Sql As String

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object

			Sql = "Select Count(*) From Dati Where idTipologia=" & idTipologia
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			Quanti = Rec(0).Value
			Rec.Close()
		End If

		Return Quanti
	End Function

	<WebMethod()>
	Public Function RitornaQuantiFilesPerVB(idTipologia As String, idCategoria As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Quanti As String = ""
		Dim Sql As String

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object

			If idCategoria = -1 Then
				Sql = "Select Count(*) From Dati Where idTipologia=" & idTipologia
			Else
				Sql = "Select Count(*) From Dati Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria
			End If
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			Quanti = Rec(0).Value
			Rec.Close()
		End If

		Return Quanti
	End Function

	<WebMethod()>
	Public Function RitornaMultimediaDaId(idTipologia As String, idCategoria As String, idMultimedia As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Sql As String

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object

			Sql = "Select Dati.NomeFile, Dati.Dimensioni, Dati.Data, Dati.idCategoria, Categorie.Categoria, Categorie.Percorso, Categorie.LetteraDisco From Dati " &
				"Left Join Categorie On Dati.idCategoria=Categorie.idCategoria And Dati.idTipologia=Categorie.idTipologia " &
				"Where Dati.idTipologia=" & idTipologia & " And Dati.idCategoria=" & idCategoria & " And Progressivo=" & idMultimedia
			' Dim gf As New GestioneFilesDirectory
			' Dim NomeFile As String = Server.MapPath(".") & "\Log\LogRitorno.txt"
			' gf.ScriveTestoSuFileAperto(NomeFile, Sql)
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			If Rec.Eof = False Then
				Dim Thumb As String = ""

				If idTipologia = "2" Then
					Dim Conversione As String = "" & Rec("LetteraDisco").Value.ToString
					Dim PathOriginale As String = Rec("Percorso").Value.ToString
					If Conversione <> "" And Conversione <> "--" Then
						Dim cc() As String = Conversione.Split("*")
						PathOriginale = PathOriginale.Replace(cc(0), cc(1))
					End If
					If Right(PathOriginale, 1) <> "\" Then
						PathOriginale &= "\"
					End If

					' Return "" & Rec("Categoria").Value & " - " & PathOriginale & " - " & Rec("NomeFile").value

					Thumb = CreaThumbDaVideo("" & Rec("Categoria").Value, PathOriginale, PathOriginale & Rec("NomeFile").value, Conversione)
				End If

				Ritorno = Thumb & "§" & Rec("NomeFile").Value.ToString.Replace(";", "***PV***") & ";" & Rec("Dimensioni").Value & ";" & Rec("Data").Value & ";" & Rec("idCategoria").Value & ";" & idMultimedia.ToString & ";"
			Else
				Ritorno = "ERROR: Nessun file rilevato"
			End If
			'gf.ScriveTestoSuFileAperto(NomeFile, Ritorno)
			Rec.Close()
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaImmaginiPerGriglia(QuanteImm As String, Categoria As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Sql As String

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object

			Dim Inizio As Long = 0
			Dim idCategoria As String = ""

			If Categoria <> "" And Categoria <> "Tutto" Then
				Sql = "Select * From Categorie Where idTipologia=1 And Categoria='" & Categoria & "'"
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				If Rec.Eof = False Then
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
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			Dim Quante As Long = Rec(0).Value
			Rec.Close

			If Categoria <> "" And Categoria <> "Tutto" Then
				Sql = "Select Min(Progressivo) From Dati Where idTipologia=1 And idCategoria=" & idCategoria
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
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
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				If Rec.Eof = False Then
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

	Private Function CreaThumbDaVideo(Categoria As String, Percorso As String, Video As String, Conversione As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim PathBase As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PercorsoThumbs.txt")
		Dim Barra As String = "\"

		If TipoDB = "SQLSERVER" Then
			Barra = "\"
		Else
			Barra = "/"
		End If

		If Strings.Right(PathBase, 1) <> barra Then
			PathBase &= barra
		End If

		Dim ritorno As String = "" ' "111->" & PathBase & "*" & Categoria & "*" & Percorso & "*" & Video & "*" & vbCrLf

		Dim OutPut As String = PathBase & Categoria & Barra
		Dim Nome As String = gf.TornaNomeFileDaPath(Video)
		'Dim Nome As String = Video
		Dim Estensione As String = gf.TornaEstensioneFileDaPath(Nome)

		Dim Cartella As String = Video.Replace(Nome, "")

		' ritorno &= "222->" & OutPut & "*" & Nome & "*" & Estensione & "*" & Video & "*" & Cartella & "*" & vbCrLf

		' OutPut &= Cartella

		gf.CreaDirectoryDaPercorso(OutPut)
		OutPut &= Nome.Replace(Estensione, "") & ".jpg"
		' Video = Percorso & "\" & Video

		If Not gf.EsisteFile(OutPut) Then
			Dim processoFFMpeg As Process = New Process()
			Dim pi As ProcessStartInfo = New ProcessStartInfo()

			If Conversione <> "" And Conversione.Contains("*") Then
				Dim cc() As String = Conversione.Split("*")
				Video = Video.Replace(cc(1), cc(0))
				OutPut = OutPut.Replace(cc(1), cc(0))
			End If

			If TipoDB <> "SQLSERVER" Then
				Video = Video.Replace("\", "/")
				Video = Video.Replace("//", "/")
				Video = Video.Replace("/\", "/")

				OutPut = OutPut.Replace("\", "/")
				OutPut = OutPut.Replace("//", "/")
				OutPut = OutPut.Replace("/\", "/")
			End If

			pi.Arguments = "-i """ & Video & """ -vframes 1 -an -s 1024x768 -ss 5 """ & OutPut & """"

			' Return pi.Arguments

			Dim Comando As String
			If TipoDB <> "SQLSERVER" Then
				Comando = "ffmpeg"
			Else
				Comando = Server.MapPath(".") & "\ffmpeg.exe"
			End If

			pi.FileName = Comando
			' gf.ScriveTestoSuFileAperto(Server.MapPath(".") & "\Buttami.txt", pi.Arguments)
			pi.WindowStyle = ProcessWindowStyle.Normal
			processoFFMpeg.StartInfo = pi
			processoFFMpeg.StartInfo.UseShellExecute = False
			processoFFMpeg.StartInfo.RedirectStandardOutput = True
			processoFFMpeg.StartInfo.RedirectStandardError = True
			processoFFMpeg.Start()

			Dim OutPutP As String = processoFFMpeg.StandardOutput.ReadToEnd()
			ritorno = OutPutP & "****"
			Dim Err As String = processoFFMpeg.StandardError.ReadToEnd()
			ritorno &= Err & "*****"

			' Return ritorno

			processoFFMpeg.WaitForExit()
		End If

		Return OutPut.Replace(PathBase, "")
	End Function

	<WebMethod()>
	Public Function RitornaFiles() As String
		If StaLeggendoImmagini Then
			Return "ERROR: Sto già caricando immagini"
		End If

		StaLeggendoImmagini = True

		Dim gf As New GestioneFilesDirectory

		Dim NomeFileLog As String = Server.MapPath(".") & "\Logs\Log_" & dataAttuale() & ".txt"
		gf.ApreFileDiTestoPerScrittura(NomeFileLog)

		Try
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
			Dim Db As New clsGestioneDB(TipoDB)
			Dim Sql As String
			Dim Barra As String = "\"

			If TipoDB = "SQLSERVER" Then
				Barra = "\"
			Else
				Barra = "/"
			End If

			Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
			If ConnessioneSQL <> "" Then
				Dim idCategoria As Integer = 0

				Try
					gf.CreaDirectoryDaPercorso(Server.MapPath(".") & "\Logs")
				Catch ex As Exception

				End Try

				Db.EsegueSql(Server.MapPath("."), "Delete From Categorie", ConnessioneSQL)
				Db.EsegueSql(Server.MapPath("."), "Delete From Dati", ConnessioneSQL)

				gf.ScriveTestoSuFileAperto(dataAttuale() & " - VIDEO")
				gf.ScriveTestoSuFileAperto(dataAttuale() & " -----------------------------------------------------------")
				Conta = 0
				For Each p As String In PathVideo
					If p.Trim <> "" Then
						gf.ScriveTestoSuFileAperto(dataAttuale() & " - Elaborazione video: " & p)

						Dim pp() As String = p.Split(";")

						idCategoria += 1
						Sql = "Insert Into categorie Values (" & idCategoria & ", 2, '" & pp(0).Replace("'", "''") & "', '" & pp(1).Replace("'", "''") & "', '" & pp(2) & "', '" & pp(3) & "')"
						Dim sRitorno As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

						If Strings.Right(pp(1), 1) <> Barra Then
							pp(1) &= Barra
						End If
						gf.ScriveTestoSuFileAperto(dataAttuale() & " - Scansione cartella: " & pp(1))
						gf.ScansionaDirectorySingola(pp(1))
						Dim qFiles As Integer = gf.RitornaQuantiFilesRilevati
						Dim Files() As String = gf.RitornaFilesRilevati
						gf.ScriveTestoSuFileAperto(dataAttuale() & " - Numero files rilevati: " & qFiles)

						For i As Integer = 1 To qFiles
							If (i / 1000 = Int(i / 1000)) Then
								gf.ScriveTestoSuFileAperto(dataAttuale() & " - Scrittura: " & i & "/" & qFiles)
							End If
							'Dim sf As New StrutturaFiles

							'sf.Categoria = Conta
							'sf.NomeFile = Files(i).Replace(pp(1), "").Replace(";", "**PV***").Replace("§", "***COSO***")
							'sf.DimensioniFile = FileLen(Files(i))
							'sf.DataFile = FileDateTime(Files(i))

							'FilesVideo.Add(sf)
							Conta += 1

							Dim SoloNome As String = Files(i).Replace(pp(1), "").Replace("'", "''")
							If TipoDB = "SQLSERVER" Then
								SoloNome = SoloNome.Replace("\", "/")
							End If
							Dim Dime As String = gf.TornaDimensioneFile(Files(i))
							Dim Datella As String = gf.TornaDataDiCreazione(Files(i))

							Sql = "Insert Into dati Values (" &
								" " & Conta & ", " &
								"2, " &
								" " & idCategoria & ", " &
								"'" & SoloNome & "', " &
								" " & Dime & ", " &
								"'" & Datella & "' " &
								")"
							sRitorno = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
							If sRitorno <> "OK" Then
								gf.ScriveTestoSuFileAperto(dataAttuale() & " - Errore: " & sRitorno)

								StaLeggendoImmagini = False

								gf.ChiudeFileDiTestoDopoScrittura()

								Return "ERROR: " & sRitorno & " -> " & Sql
							End If
						Next
						gf.ScriveTestoSuFileAperto(dataAttuale() & " - Fine Scrittura: " & qFiles & "/" & qFiles)
					End If
				Next
				gf.ScriveTestoSuFileAperto(dataAttuale() & " -----------------------------------------------------------")

				gf.ScriveTestoSuFileAperto("")
				gf.ScriveTestoSuFileAperto("")
				gf.ScriveTestoSuFileAperto(dataAttuale() & " - IMMAGINI")
				gf.ScriveTestoSuFileAperto(dataAttuale() & " -----------------------------------------------------------")
				Conta = 0
				idCategoria = 0
				For Each p As String In PathImmagini
					If p.Trim <> "" Then
						gf.ScriveTestoSuFileAperto(dataAttuale() & " - Elaborazione immagini: " & p)
						Dim pp() As String = p.Split(";")

						idCategoria += 1
						Sql = "Insert Into categorie Values (" & idCategoria & ", 1, '" & pp(0).Replace("'", "''") & "', '" & pp(1).Replace("'", "''") & "', '" & pp(2) & "', '" & pp(3) & "')"
						Dim sRitorno As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

						If Strings.Right(pp(1), 1) <> Barra Then
							pp(1) &= Barra
						End If
						gf.ScriveTestoSuFileAperto(dataAttuale() & " - Scansione cartella: " & pp(1))
						gf.ScansionaDirectorySingola(pp(1))
						Dim qFiles As Integer = gf.RitornaQuantiFilesRilevati
						Dim Files() As String = gf.RitornaFilesRilevati
						gf.ScriveTestoSuFileAperto(dataAttuale() & " - Numero files: " & qFiles)
						' Conta += 1
						For i As Integer = 1 To qFiles
							If (i / 1000 = Int(i / 1000)) Then
								gf.ScriveTestoSuFileAperto(dataAttuale() & " - Scrittura: " & i & "/" & qFiles)
							End If
							'Dim sf As New StrutturaFiles

							'sf.Categoria = Conta
							'sf.NomeFile = Files(i).Replace(pp(1), "").Replace(";", "**PV***").Replace("§", "***COSO***")
							'sf.DimensioniFile = FileLen(Files(i))
							'sf.DataFile = FileDateTime(Files(i))

							'FilesImmagini.Add(sf)
							Dim SoloNome As String = Files(i).Replace(pp(1), "").Replace("'", "''")
							If TipoDB = "SQLSERVER" Then
								SoloNome = SoloNome.Replace("\", "/")
							End If
							Dim Dime As String = gf.TornaDimensioneFile(Files(i))
							Dim Datella As String = gf.TornaDataDiCreazione(Files(i))

							Conta += 1
							Sql = "Insert Into dati Values (" &
								" " & Conta & ", " &
								"1, " &
								" " & idCategoria & ", " &
								"'" & SoloNome & "', " &
								" " & Dime & ", " &
								"'" & Datella & "' " &
								")"
							sRitorno = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
							If sRitorno <> "OK" Then
								gf.ScriveTestoSuFileAperto(dataAttuale() & " - Errore: " & sRitorno)
								StaLeggendoImmagini = False

								gf.ChiudeFileDiTestoDopoScrittura()

								Return "ERROR: " & sRitorno
							End If
						Next
						gf.ScriveTestoSuFileAperto(dataAttuale() & " - Fine Scrittura: " & qFiles & "/" & qFiles)
					End If
				Next
				gf.ScriveTestoSuFileAperto(dataAttuale() & " -----------------------------------------------------------")

				gf.ScriveTestoSuFileAperto("")
				gf.ScriveTestoSuFileAperto(dataAttuale() & " - RIEPILOGO")
				If TipoDB = "SQLSERVER" Then
					Sql = "Select Tipologia, Categoria, Isnull(Count(*),0) As Quanti From Dati A " &
						"Left Join Categorie B On A.idCategoria = B.idCategoria " &
						"Left Join Tipologie C On A.idTipologia = C.idTipologia " &
						"Group By Tipologia, Categoria " &
						"Order By 1,2"
				Else
					Sql = "Select Tipologia, Categoria, COALESCE(Count(*),0) As Quanti From Dati A " &
						"Left Join Categorie B On A.idCategoria = B.idCategoria " &
						"Left Join Tipologie C On A.idTipologia = C.idTipologia " &
						"Group By Tipologia, Categoria " &
						"Order By 1,2"
				End If
				Dim Rec As Object
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				Do Until Rec.Eof
					Dim Tipologia As String = Rec("Tipologia").Value
					Dim Categoria As String = Rec("Categoria").Value
					Dim Quanti As String = Rec("Quanti").Value

					gf.ScriveTestoSuFileAperto(dataAttuale() & " - " & Tipologia & ": " & Categoria & " -> Files " & Quanti)

					Rec.MoveNext
				Loop
				Rec.Close
				gf.ScriveTestoSuFileAperto(dataAttuale() & " -----------------------------------------------------------")

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
		Catch ex As Exception
			gf.ScriveTestoSuFileAperto(dataAttuale() & ex.Message)
		End Try

		StaLeggendoImmagini = False

		gf.ChiudeFileDiTestoDopoScrittura()

		Return "*"
	End Function

	<WebMethod()>
	Public Function EffettuaRicerca(idTipologia As String, Categoria As String, Ricerca As String, Quante As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim Ritorno As String = ""
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Rec As Object
		Dim Sql As String

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then

			Dim idCategoria As String = ""

			If Categoria <> "" And Categoria <> "Tutto" Then
				Sql = "Select * From Categorie Where idTipologia=" & idTipologia & " And Categoria='" & Categoria & "'"
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				If Rec.Eof = False Then
					idCategoria = Rec("idCategoria").Value
				Else
					Return "ERROR: Categoria non trovata"
				End If
				Rec.Close
			End If

			If Quante = -1 Then Quante = 9999

			Dim Barra As String = "\"

			If TipoDB = "SQLSERVER" Then
				Barra = "\"
			Else
				Barra = "/"
			End If

			If TipoDB = "SQLSERVER" Then
				Sql = "Select Top " & Quante & " Dati.Progressivo, Dati.NomeFile, Dati.Dimensioni, Dati.Data, Dati.idCategoria, Categorie.Categoria, Categorie.Percorso, Categorie.LetteraDisco From Dati " &
					"Left Join Categorie On Dati.idCategoria=Categorie.idCategoria And Dati.idTipologia=Categorie.idTipologia " &
					"Where Dati.idTipologia=" & idTipologia & " And Dati.NomeFile Like '%" & Ricerca & "%' And Dati.idCategoria = " & idCategoria & " And Dati.NomeFile Not Like '%.nomedia%' "
			Else
				Sql = "Select Dati.Progressivo, Dati.NomeFile, Dati.Dimensioni, Dati.Data, Dati.idCategoria, Categorie.Categoria, Categorie.Percorso, Categorie.LetteraDisco From Dati " &
					"Left Join Categorie On Dati.idCategoria=Categorie.idCategoria And Dati.idTipologia=Categorie.idTipologia " &
					"Where Dati.idTipologia=" & idTipologia & " And Dati.NomeFile Like '%" & Ricerca & "%' And Dati.idCategoria = " & idCategoria & " And Dati.NomeFile Not Like '%.nomedia%' Limit " & Quante
			End If
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			If Rec.Eof = False Then
				Do Until Rec.eof
					Dim Thumb As String = ""

					If idTipologia = "2" Then
						Dim Conversione As String = "" & Rec("LetteraDisco").Value.ToString
						Dim PathOriginale As String = Rec("Percorso").Value.ToString

						'If Conversione <> "" Then
						'	Dim cc() As String = Conversione.Split("*")
						'	PathOriginale = PathOriginale.Replace(cc(0), cc(1))
						'End If
						If Right(PathOriginale, 1) <> Barra Then
							PathOriginale &= Barra
						End If

						' Return "" & Rec("Categoria").Value & " - " & PathOriginale & " - " & Rec("NomeFile").value

						Thumb = CreaThumbDaVideo("" & Rec("Categoria").Value, PathOriginale, PathOriginale & Rec("NomeFile").value, Conversione)
						' Thumb = CreaThumbDaVideo(Rec("Categoria").Value, Rec("Percorso").Value, Rec("NomeFile").value)
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

		If ContatoreRiletturaImmagini = 10 Then
			ContatoreRiletturaImmagini = 0
			QuanteImmaginiSfondi = 0
			ListaImmagini = New List(Of String)
		End If

		Dim Barra As String = "\"

		If TipoDB = "SQLSERVER" Then
			Barra = "\"
		Else
			Barra = "/"
		End If

		If QuanteImmaginiSfondi = 0 Then
			If Strings.Right(Path, 1) <> Barra Then
				Path &= Barra
			End If
			gf.ScansionaDirectorySingola(Path)
			Dim filetti() As String = gf.RitornaFilesRilevati
			QuanteImmaginiSfondi = gf.RitornaQuantiFilesRilevati

			For i As Integer = 1 To QuanteImmaginiSfondi
				ListaImmagini.Add(filetti(i))
			Next
		Else
			ContatoreRiletturaImmagini += 1
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

		If gf.EsisteFile(NomeFile) Then
			Immagine = gf.LeggeFileIntero(NomeFile)
		Else
			gf.ApreFileDiTestoPerScrittura(NomeFile)

			For Each foundFile As String In My.Computer.FileSystem.GetFiles(Server.MapPath(".") & "\Sfondi")
				Dim conta As Integer = 0

				While gf.EsisteFile(foundFile)
					gf.EliminaFileFisico(foundFile)
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

			gf.ScriveTestoSuFileAperto(Immagine)
			gf.ChiudeFileDiTestoDopoScrittura()
		End If

		Return Immagine
	End Function

	<WebMethod()>
	Public Function EliminaMultimediaDaId(idTipologia As String, idCategoria As String, idMultimedia As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Sql As String
		Dim Barra As String = "\"

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object

			If TipoDB = "SQLSERVER" Then
				Barra = "\"
			Else
				Barra = "/"
			End If

			Sql = "Select Dati.NomeFile, Dati.Dimensioni, Dati.Data, Dati.idCategoria, Categorie.Categoria, Categorie.Percorso, Categorie.LetteraDisco From Dati " &
				"Left Join Categorie On Dati.idCategoria=Categorie.idCategoria And Dati.idTipologia=Categorie.idTipologia " &
				"Where Dati.idTipologia=" & idTipologia & " And Dati.idCategoria=" & idCategoria & " And Progressivo=" & idMultimedia
			' Dim gf As New GestioneFilesDirectory
			' Dim NomeFile As String = Server.MapPath(".") & "\Log\LogRitorno.txt"
			' gf.ScriveTestoSuFileAperto(NomeFile, Sql)
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			If Rec.Eof = False Then
				Dim Thumb As String = ""

				If idTipologia = "2" Then
					Dim Conversione As String = "" & Rec("LetteraDisco").Value.ToString
					Dim PathOriginale As String = Rec("Percorso").Value.ToString
					If Conversione <> "" And Conversione <> "--" Then
						Dim cc() As String = Conversione.Split("*")
						PathOriginale = PathOriginale.Replace(cc(0), cc(1))
					End If
					If Right(PathOriginale, 1) <> Barra Then
						PathOriginale &= Barra
					End If

					Dim filetto As String = PathOriginale & Rec("NomeFile").value

					Dim gf As New GestioneFilesDirectory
					Ritorno = gf.EliminaFileFisico(filetto)
					If Ritorno = "" Then
						Ritorno = "*"
					End If
				End If
				' Ritorno = Thumb & "§" & Rec("NomeFile").Value.ToString.Replace(";", "***PV***") & ";" & Rec("Dimensioni").Value & ";" & Rec("Data").Value & ";" & Rec("idCategoria").Value & ";" & idMultimedia.ToString & ";"
			Else
				Ritorno = "ERROR: Nessun file rilevato"
			End If
			'gf.ScriveTestoSuFileAperto(NomeFile, Ritorno)
			Rec.Close()
		End If

		Return Ritorno
	End Function

End Class