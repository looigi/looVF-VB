Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.IO
Imports System.Threading
Imports System.Timers
Imports System.Web.Script.Serialization
Imports looVF.GestioneImmagini

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://looVF.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class looVF
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function ImpostaSuccessivoMultimedia(idTipologia As String, idMultimedia As String) As String
		Dim Ritorno As String = "*"
		Dim gf As New GestioneFilesDirectory

		If idTipologia = 1 Then
			'gf.EliminaFileFisico(Server.MapPath(".") & "/idImmagine.txt")
			'gf.CreaAggiornaFile(Server.MapPath(".") & "/idImmagine.txt", idMultimedia)
			UltimoMultimediaImm = idMultimedia
		Else
			'gf.EliminaFileFisico(Server.MapPath(".") & "/idVideo.txt")
			'gf.CreaAggiornaFile(Server.MapPath(".") & "/idVideo.txt", idMultimedia)
			UltimoMultimediaVid = idMultimedia
		End If

		Return idMultimedia
	End Function

	Private Function RitornaSuccessivoMultimediaPerPreferiti(db As clsGestioneDB, ConnessioneSQL As String, idTipologia As String, Categoria As String, Filtro As String,
															 Random As String, NomeFileLog As String, idCategoria As String) As String
		Dim Ritorno As String = ""
		Dim QuanteRighePreferiti As Long = 0
		Dim Sql As String = ""
		Dim NomeTabella As String = ""
		Dim Altro As String = ""
		Dim Rec As Object

		If Filtro <> "" Then
			Altro = " And Upper(B.NomeFile) Like '%" & Filtro.ToUpper & "%'"
		End If

		If Categoria = "Preferiti" Then
			NomeTabella = "Preferiti"
		Else
			NomeTabella = "PreferitiProt"
		End If

		ScriveLogGlobale(NomeFileLog, "Ricerca Successivo per preferiti: idTipologia " & idTipologia & " idCategoria " & idCategoria & " Categoria " & Categoria & " Filtro " & Filtro & " Random " & Random)

		Sql = "Select Coalesce(Count(*), 0) As Quante From " & NomeTabella & " A " &
			"Left Join Dati B On A.idTipologia=B.idTipologia And A.idCategoria=B.idCategoria And A.Progressivo=B.Progressivo " &
			"Where (B.Eliminata = 'N' Or B.Eliminata = 'n') And A.idTipologia=" & idTipologia & " " & Altro
		If idCategoria <> "" Then
			Sql &= " And A.idCategoria=" & idCategoria
		End If
		Rec = db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
		QuanteRighePreferiti = Rec("Quante").Value
		Rec.Close
		ScriveLogGlobale(NomeFileLog, "Quante righe preferiti: " & QuanteRighePreferiti)

		Dim Ultimo As Integer = -1

		If idTipologia = 1 Then
			Ultimo = UltimoMultimediaImm
		Else
			Ultimo = UltimoMultimediaVid
		End If
		ScriveLogGlobale(NomeFileLog, "Ultimo MM impostato: " & Ultimo)

		Static x As Random = New Random()

		Dim y As Long = -1

		If Random = "S" Or Random = "" Then
			y = x.Next(QuanteRighePreferiti)
			ScriveLogGlobale(NomeFileLog, "Valore Random per Random 'S': " & y & "/" & QuanteRighePreferiti)
		Else
			If Ultimo <> -1 Then
				y = Ultimo + 1
				If y > QuanteRighePreferiti Then
					y = 0
				End If
				ScriveLogGlobale(NomeFileLog, "Valore Sequenziale per Random 'N': " & y & "/" & QuanteRighePreferiti)
			Else
				y = x.Next(QuanteRighePreferiti)
				ScriveLogGlobale(NomeFileLog, "Valore Random per Random 'N': " & y & "/" & QuanteRighePreferiti & " Ultimo = -1")
			End If
		End If

		If y = -1 Then
			Return "ERROR: Non riesco a impostare il valore. Quante righe: " & QuanteRighePreferiti
		End If

		If idTipologia = 1 Then
			UltimoMultimediaImm = y
		Else
			UltimoMultimediaVid = y
		End If

		Dim NumeroImmagine As Long
		Dim idCategoriaRilevato As Integer
		Dim CategoriaRilevata As String
		Dim Altro2 As String = ""

		If Filtro <> "" Then
			Altro2 = " Where Upper(C.NomeFile) Like '%" & Filtro.ToUpper & "%'"
		End If

		Sql = "Select idTipologia, idCategoria, Progressivo, Categoria, NomeFile From ( " &
			"SELECT ROW_NUMBER() OVER(Order BY idTipologia, idCategoria, progressivo) As Numero, A.idTipologia, A.idCategoria, A.Progressivo, B.Categoria, C.NomeFile " &
			"FROM " & NomeTabella & " A " &
			"Left Join categorie B On A.idtipologia = B.idtipologia And A.idCategoria = B.idcategoria " &
			"Left Join Dati C On A.idtipologia = C.idtipologia And A.idCategoria = C.idcategoria And A.Progressivo = C.Progressivo  " &
			" " & Altro2 & " " &
			") As a Where Numero=" & y
		'"where idtipologia=" & idTipologia & " " & ' " and idcategoria=" & idCategoria & " " &
		ScriveLogGlobale(NomeFileLog, Sql)
		Rec = db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
		If Rec.Eof = False Then
			NumeroImmagine = Rec("Progressivo").Value
			idCategoriaRilevato = Rec("idCategoria").Value
			CategoriaRilevata = Rec("Categoria").Value

			Ritorno = NumeroImmagine.ToString & ";" & idCategoriaRilevato & ";" & QuanteRighePreferiti & ";" & CategoriaRilevata & ";" & Ultimo & ";" & y
		Else
			Ritorno = "ERROR: Nessun multimedia rilevato"
		End If
		Rec.Close

		ScriveLogGlobale(NomeFileLog, "Ritorno: " & Ritorno)

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaSuccessivoMultimediaNuovo(idTipologia As String, Categoria As String, Filtro As String, Random As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Sql As String
		Dim NuovaRicerca As String = idTipologia & ";" & Categoria & ";" & Filtro
		Dim NomeFileLog As String = Server.MapPath(".") & "/Logs/RitornaMultimediaSuccessivo.txt"

		ScriveLogGlobale(NomeFileLog, "-----------------------------------------")
		ScriveLogGlobale(NomeFileLog, "Inizio")

		Dim gf As New GestioneFilesDirectory
		'Dim NomeFile As String = Server.MapPath(".") & "\Log\LogRitorno.txt"
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim idCategoria As String = ""
			Dim Altro As String = ""

			ScriveLogGlobale(NomeFileLog, "Connessione Aperta")
			'gf.ScriveTestoSuFileAperto(NomeFile, idTipologia & "-" & Categoria)

			Dim Quante As Long = 0
			Dim Indici As New List(Of Integer)
			Dim Categorie As New List(Of Integer)
			Dim NomeCategorie As New List(Of String)

			ScriveLogGlobale(NomeFileLog, "Categoria: " & Categoria)

			If Filtro <> "" Then
				Altro = " And Upper(NomeFile) Like '%" & Filtro.ToUpper & "%'"
			End If

			If Categoria <> "" And Categoria <> "Tutto" And Not Categoria.Contains("Preferiti") Then
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

			If Categoria = "Preferiti" Or Categoria = "Preferiti Prot" Then
				ScriveLogGlobale(NomeFileLog, "Categoria uguale a Preferiti: " & Categoria & ". Vado alla funzione adatta")

				Ritorno = RitornaSuccessivoMultimediaPerPreferiti(Db, ConnessioneSQL, idTipologia, Categoria, Filtro, Random, NomeFileLog, idCategoria)

				Return Ritorno
			End If

			ScriveLogGlobale(NomeFileLog, "IDCategoria: " & idCategoria)

			If Filtro <> "" Then
				ScriveLogGlobale(NomeFileLog, "Filtro impostato: " & Filtro)

				Quante = 0
				Sql = "Select * From Dati Where (Eliminata = 'N' Or Eliminata = 'n') And idTipologia=" & idTipologia & " " & Altro
				If idCategoria <> "" Then
					Sql &= " And idCategoria=" & idCategoria
				End If
				'gf.ScriveTestoSuFileAperto(NomeFile, Sql)
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				Do Until Rec.Eof
					Indici.Add(Rec("Progressivo").Value)

					Quante += 1
					Rec.MoveNext
				Loop
				Rec.Close

				ScriveLogGlobale(NomeFileLog, "Indici Rilevati: " & Indici.Count - 1)
			Else
				ScriveLogGlobale(NomeFileLog, "Filtro NON impostato")
				ScriveLogGlobale(NomeFileLog, "Nuova Ricerca: " & NuovaRicerca & " Vecchia Ricerca: " & VecchiaRicerca)

				If NuovaRicerca <> VecchiaRicerca Then
					ScriveLogGlobale(NomeFileLog, "Nuova Ricerca diversa da Vecchia Ricerca")

					Sql = "Select Count(*) From Dati Where (Eliminata = 'N' Or Eliminata = 'n') And idTipologia=" & idTipologia
					If idCategoria <> "" Then
						Sql &= " And idCategoria=" & idCategoria
					End If
					'gf.ScriveTestoSuFileAperto(NomeFile, Sql)
					Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
					Quante = Rec(0).Value
					Rec.Close
					VecchioQuante = Quante
					'gf.ScriveTestoSuFileAperto(NomeFile, Quante)
					VecchiaRicerca = NuovaRicerca
				Else
					Quante = VecchioQuante
					ScriveLogGlobale(NomeFileLog, "Nuova Ricerca uguale Vecchia Ricerca")
				End If
				ScriveLogGlobale(NomeFileLog, "Numero files rilevati: " & Quante)
			End If
			'Else
			'	ScriveLogGlobale(NomeFileLog, "Categoria uguale a Preferiti")

			'	If Filtro <> "" Then
			'		Altro = " And Upper(NomeFile) Like '%" & Filtro.ToUpper & "%'"
			'	End If

			'	Quante = 0

			'	Dim NomeTabella As String = ""

			'	If Categoria = "Preferiti" Then
			'		NomeTabella = "Preferiti"
			'	Else
			'		NomeTabella = "PreferitiProt"
			'	End If

			'	Sql = "Select A.*, B.NomeFile, C.Categoria From " & NomeTabella & " A " &
			'		"Left Join Dati B On A.idTipologia=B.idTipologia And A.idCategoria=B.idCategoria And A.Progressivo=B.Progressivo " &
			'		"Left Join Categorie C On A.idTipologia=C.idTipologia And A.idCategoria=C.idCategoria " &
			'		"Where (B.Eliminata = 'N' Or B.Eliminata = 'n') And A.idTipologia=" & idTipologia & " " & Altro
			'	'gf.ScriveTestoSuFileAperto(NomeFile, Sql)
			'	Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			'	Do Until Rec.Eof
			'		Indici.Add(Rec("Progressivo").Value)
			'		Categorie.Add(Rec("idCategoria").Value)
			'		NomeCategorie.Add(Rec("Categoria").Value)

			'		Quante += 1
			'		Rec.MoveNext
			'	Loop
			'	Rec.Close

			'	ScriveLogGlobale(NomeFileLog, "Numero files rilevati: " & Quante & " Indici: " & Indici.Count - 1 & " Categorie: " & Categorie.Count - 1 & " Nome Categorie: " & NomeCategorie.Count - 1)
			'End If

			Dim Ultimo As Integer = -1

			If idTipologia = 1 Then
				Ultimo = UltimoMultimediaImm
				'If gf.EsisteFile(Server.MapPath(".") & "/idImmagine.txt") Then
				'	Ultimo = Val(gf.LeggeFileIntero(Server.MapPath(".") & "/idImmagine.txt").Trim)
				'End If
			Else
				Ultimo = UltimoMultimediaVid
				'If gf.EsisteFile(Server.MapPath(".") & "/idVideo.txt") Then
				'	Ultimo = Val(gf.LeggeFileIntero(Server.MapPath(".") & "/idVideo.txt").Trim)
				'End If
			End If
			'Return Ultimo
			ScriveLogGlobale(NomeFileLog, "Ultimo MM impostato: " & Ultimo)

			Static x As Random = New Random()

			Dim y As Long = -1

			If Random = "S" Or Random = "" Then
				y = x.Next(Quante)
				ScriveLogGlobale(NomeFileLog, "Valore Random per Random 'S': " & y & "/" & Quante)
			Else
				If Ultimo <> -1 Then
					y = Ultimo + 1
					If y > Quante Then
						y = 0
					End If
					ScriveLogGlobale(NomeFileLog, "Valore Sequenziale per Random 'N': " & y & "/" & Quante)
				Else
					y = x.Next(Quante)
					ScriveLogGlobale(NomeFileLog, "Valore Random per Random 'N': " & y & "/" & Quante & " Ultimo = -1")
				End If
			End If

			If idTipologia = 1 Then
				UltimoMultimediaImm = y
				'gf.EliminaFileFisico(Server.MapPath(".") & "/idImmagine.txt")
				'gf.CreaAggiornaFile(Server.MapPath(".") & "/idImmagine.txt", y)
			Else
				UltimoMultimediaVid = y
				'gf.EliminaFileFisico(Server.MapPath(".") & "/idVideo.txt")
				'gf.CreaAggiornaFile(Server.MapPath(".") & "/idVideo.txt", y)
			End If

			'Return "Random: " & Random & ";Ultimo Multimedia: " & UltimoMultimedia & ";Y: " & y & ";Quante: " & Quante
			Dim Inizio As Long = 0

			'If idCategoria <> "" Then
			If Filtro = "" Then '  And Categoria <> "Preferiti" And Categoria <> "Preferiti Prot"
				ScriveLogGlobale(NomeFileLog, "Filtro nullo")

				Sql = "Select Coalesce(Min(Progressivo),0) From Dati Where (Eliminata = 'N' Or Eliminata = 'n') And idTipologia=" & idTipologia & " " & Altro & " "
				If idCategoria <> "" Then
					Sql &= "And idCategoria=" & idCategoria
				End If
				'gf.ScriveTestoSuFileAperto(NomeFile, Sql)
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				Inizio = (Rec(0).Value - 1)
				Rec.Close

				ScriveLogGlobale(NomeFileLog, "Inizio: " & Inizio)
			Else
				ScriveLogGlobale(NomeFileLog, "Filtro NON Nullo (" & Filtro & ")")

				If Indici.Count > y Then
					Inizio = Indici(y)
				Else
					Inizio = -1
				End If
				'y = 0

				'If Categoria <> "Preferiti" And Categoria <> "Preferiti Prot" Then
				'End If

				ScriveLogGlobale(NomeFileLog, "Inizio: " & Inizio & " Indici: " & Indici.Count - 1)
			End If

			If Random = "N" Then ' And (Categoria = "Preferiti" Or Categoria = "Preferiti Prot") Then
				Inizio = 0

				ScriveLogGlobale(NomeFileLog, "Random No e NON Preferiti, azzero l'inizio")
			End If

			'If Categoria = "Preferiti" Or Categoria = "Preferiti Prot" Then
			'	idCategoria = Categorie.Item(y)
			'	Categoria = NomeCategorie.Item(y)

			'	Sql = "Select * From Dati Where (Eliminata = 'N' Or Eliminata = 'n') And idTipologia=" & idTipologia & " And idCategoria=" & idCategoria
			'	Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			'	idCategoria = Rec("idCategoria").Value
			'	Rec.Close

			'	Dim NomeTabella As String = ""

			'	If Categoria = "Preferiti" Then
			'		NomeTabella = "Preferiti"
			'	Else
			'		NomeTabella = "PreferitiProt"
			'	End If

			'	Sql = "Select Progressivo From ( " &
			'		"SELECT ROW_NUMBER() OVER(Order BY idTipologia, idCategoria, progressivo) As Numero, Progressivo " &
			'		"FROM " & NomeTabella & " " &
			'		"where idtipologia=" & idTipologia & " " & ' " and idcategoria=" & idCategoria & " " &
			'		") As a Where Numero=" & (Inizio + y)
			'	ScriveLogGlobale(NomeFileLog, Sql)
			'	Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			'	Inizio = Rec("Progressivo").Value
			'	Rec.Close

			'	ScriveLogGlobale(NomeFileLog, "Categoria Preferiti. Imposto idCategoria e Categoria: " & idCategoria & " - " & Categoria & ". Valore riga: " & Inizio)
			'End If

			'Else
			'	idCategoria = -1
			'End If
			If idCategoria = "" Then
				ScriveLogGlobale(NomeFileLog, "Categoria Nulla. La ricerco per idTipologia " & idTipologia & " e idMultimedia " & Inizio + y)

				Sql = "Select idCategoria From Dati Where (Eliminata = 'N' Or Eliminata = 'n') And idTipologia=" & idTipologia & " And Progressivo=" & Inizio + y
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				idCategoria = Rec("idCategoria").Value
				Rec.Close

				ScriveLogGlobale(NomeFileLog, "Categoria Rilevata: " & idCategoria)
			End If

			Dim Quanto As Integer

			Quanto = (Inizio + y)
			'Return Inizio & ";" & y
			ScriveLogGlobale(NomeFileLog, "Id Multimedia ritornato: " & Quanto & " - Inizio: " & Inizio & " Y: " & y)

			Ritorno = Quanto.ToString & ";" & idCategoria & ";" & Quante & ";" & Categoria & ";" & Ultimo & ";" & Inizio
			ScriveLogGlobale(NomeFileLog, "Ritorno: " & Ritorno)
		End If
		'gf.ScriveTestoSuFileAperto(NomeFile, Ritorno)

		Return Ritorno
	End Function

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
	Public Function ImpostaPreferito(idTipologia As String, idCategoria As String, idMultimedia As String, SiNo As String, Protetto As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Sql As String

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim NomeTabella As String = "Preferiti"

			If Protetto = "S" Then
				NomeTabella = "PreferitiProt"
			End If

			If SiNo = "S" Then
				Sql = "Insert Into " & NomeTabella & " Values (" & idCategoria & ", " & idTipologia & ", " & idMultimedia & ")"
			Else
				Sql = "Delete From " & NomeTabella & " Where idcategoria=" & idCategoria & " And idtipologia=" & idTipologia & " And progressivo=" & idMultimedia
			End If

			Ritorno = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL)
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaMultimediaDaId(idTipologia As String, idCategoria As String, idMultimedia As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Sql As String

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim Rec2 As Object

			Sql = "Select Dati.NomeFile, Dati.Dimensioni, Dati.Data, Dati.idCategoria, Categorie.Categoria, Categorie.Percorso, Categorie.LetteraDisco, Coalesce(preferiti.idCategoria, -1) As Preferito, Coalesce(preferitiprot.idCategoria, -1) As PreferitoProt From Dati " &
				"Left Join Categorie On Dati.idCategoria=Categorie.idCategoria And Dati.idTipologia=Categorie.idTipologia " &
				"Left Join preferiti On Dati.idCategoria=preferiti.idCategoria And Dati.idTipologia=preferiti.idTipologia And Dati.progressivo=preferiti.progressivo " &
				"Left Join preferitiprot On Dati.idCategoria=preferitiprot.idCategoria And Dati.idTipologia=preferitiprot.idTipologia And Dati.progressivo=preferitiprot.progressivo " &
				"Where Dati.idTipologia=" & idTipologia & " And Dati.idCategoria=" & idCategoria & " And dati.Progressivo=" & idMultimedia
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

				Dim Preferito As String = IIf(Val(Rec("Preferito").Value) > -1, "S", "N")
				If Preferito = "N" Then
					Preferito = IIf(Val(Rec("PreferitoProt").Value) > -1, "S", "N")
				End If

				Dim DatiHash As String = ""

				If idTipologia = 1 Then
					Sql = "Select * From informazioniimmagini Where idCategoria=" & idCategoria & " And idMultimedia=" & idMultimedia
					Rec2 = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
					If Rec2.Eof = False Then
						DatiHash = Rec2("Hash").Value & ";" & Rec2("Punti").Value & ";" & Rec2("Width").Value & ";" & Rec2("Height").Value & ";" & Rec2("DataOra").Value
					Else
						'If StaEffettuandoConversioneAutomaticaI = False Then
						DatiHash = CalcolaHashImmagine(idCategoria, idMultimedia, "S", -1)
						If DatiHash.Contains("ERROR:") Then
							DatiHash = ""
						End If
						'End If
					End If
					Rec2.Close
				Else
					Sql = "Select * From InformazioniVideo Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria & " And idMultimedia=" & idMultimedia
					Rec2 = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL, False)
					If Rec2.Eof = False Then
						DatiHash = Rec2("Jsone").Value
					Else
						'If StaEffettuandoConversioneAutomatica = False Then
						'	DatiHash = ConverteVideo(idTipologia, idCategoria, idMultimedia, "N")
						'	If DatiHash.Contains("ERROR:") Then
						'		DatiHash = ""
						'	End If
						'End If
					End If
					Rec2.Close
				End If

				Ritorno = Thumb & "§" & Rec("NomeFile").Value.ToString.Replace(";", "***PV***") & ";" & Rec("Dimensioni").Value & ";" & Rec("Data").Value & ";" & Rec("idCategoria").Value & ";" & idMultimedia.ToString & ";" & Preferito & ";" & DatiHash
			Else
				Ritorno = "ERROR: Nessun file rilevato"
			End If
			'gf.ScriveTestoSuFileAperto(NomeFile, Ritorno)
			Rec.Close()
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function StaEseguendoRefresh() As String
		Dim StaConvertendoVideo As String = "N"
		Dim DeveRinominareVideoConvertito As String = ""
		Dim NomeFile1 As String = Server.MapPath(".") & "/Logs/ffmpegout.txt"
		Dim NomeFile2 As String = Server.MapPath(".") & "/Logs/FinitaConversione.txt"
		Dim gf As New GestioneFilesDirectory

		If gf.EsisteFile(NomeFile1) Then
			StaConvertendoVideo = "S"
		End If
		If gf.EsisteFile(NomeFile2) Then
			DeveRinominareVideoConvertito = gf.LeggeFileIntero(NomeFile2)
		End If

		Dim Conv As String = IIf(StaEffettuandoConversioneAutomaticaFinale = True, "S", "N")
		Dim ConvI As String = IIf(StaEffettuandoConversioneAutomaticaFinaleI = True, "S", "N")

		If StaLeggendoImmagini Then
			Return "Sto caricando multimedia;" & StaConvertendoVideo & ";" & DeveRinominareVideoConvertito & ";" & Conv & ";" & ConvI
		Else
			Return "NON Sto caricando multimedia;" & StaConvertendoVideo & ";" & DeveRinominareVideoConvertito & ";" & Conv & ";" & ConvI
		End If
	End Function

	<WebMethod()>
	Public Function ContaRigheDuranteRefresh() As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Sql As String

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object

			Sql = "Select COALESCE(Count(*),0) As QuanteRighe From Dati"
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			If Rec.Eof = False Then
				Ritorno = Rec("QuanteRighe").Value
			Else
				Ritorno = "ERROR: Nessuna riga rilevata"
			End If
			Rec.Close
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaMultimediaPerGriglia(QuanteImm As String, Categoria As String, idTipologia As String, Filtro As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Sql As String

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object

			Dim Inizio As Long = 0
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

			Dim Altro As String = ""

			If Filtro <> "" Then
				Altro = " And Upper(Dati.NomeFile) Like '%" & Filtro.ToUpper & "%' "
			End If

			Sql = "Select Count(*) From Dati Where Dati.idTipologia=" & idTipologia & " " & Altro
			If Categoria <> "" And Categoria <> "Tutto" Then
				Sql &= " And Dati.idCategoria=" & idCategoria
			End If
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			Dim Quante As Long = Rec(0).Value
			Rec.Close

			If Categoria <> "" And Categoria <> "Tutto" And Filtro = "" Then
				Sql = "Select Min(Progressivo) From Dati Where Dati.idTipologia=" & idTipologia & " And Dati.idCategoria=" & idCategoria ' & Altro
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				Inizio = Rec(0).Value - 1
				Rec.Close
			End If

			For i As Integer = 1 To Val(QuanteImm)
				Randomize()
				Static x As Random = New Random()
				Dim idMultimedia As Long

				If Filtro = "" Then
					Dim y As Long = x.Next(Quante)
					idMultimedia = Inizio + y

					Sql = "Select Dati.NomeFile, Dati.Dimensioni, Dati.Data, Dati.idCategoria, Categorie.Categoria, Categorie.Percorso, Categorie.LetteraDisco From Dati " &
						"Left Join Categorie On Dati.idCategoria=Categorie.idCategoria And Dati.idTipologia=Categorie.idTipologia " &
						"Where Dati.idTipologia=" & idTipologia & " And Progressivo=" & idMultimedia
				Else
					Dim y As Long = x.Next(Quante)

					Sql = "Select Numero, NomeFile, Dimensioni, Data, idTipologia, idCategoria, Categoria, Progressivo, Percorso, LetteraDisco From ( " &
						"Select ROW_NUMBER() OVER(Order BY Dati.idTipologia, Dati.idCategoria, Dati.progressivo) As Numero, Dati.idTipologia, Dati.NomeFile, Dati.Dimensioni, Dati.Data, Dati.idCategoria, Categorie.Categoria, " &
						"Categorie.Percorso, Categorie.LetteraDisco, Dati.Progressivo From Dati " &
						"Left Join Categorie On Dati.idCategoria=Categorie.idCategoria And Dati.idTipologia=Categorie.idTipologia " &
						"Where Dati.idTipologia=" & idTipologia & " And Dati.idCategoria=" & idCategoria & " " & Altro & " And Categorie.Percorso Is Not Null " & ' And Progressivo=" & idMultimedia
						") As A Where Numero=" & y
				End If
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)

				If Rec.Eof = False Then
					If Filtro <> "" Then
						idMultimedia = Inizio + Val(Rec("Progressivo").Value)
					End If

					If idTipologia = 2 Then
						Dim Conversione As String = "" & Rec("LetteraDisco").Value.ToString
						Dim PathOriginale As String = Rec("Percorso").Value.ToString
						If Conversione <> "" And Conversione <> "--" Then
							Dim cc() As String = Conversione.Split("*")
							PathOriginale = PathOriginale.Replace(cc(0), cc(1))
						End If
						If Right(PathOriginale, 1) <> "\" Then
							PathOriginale &= "\"
						End If

						Dim Thumb As String = CreaThumbDaVideo("" & Rec("Categoria").Value, PathOriginale, PathOriginale & Rec("NomeFile").value, Conversione)
						Ritorno &= Thumb & ";" & Rec("NomeFile").Value.ToString.Replace(";", "***PV***") & ";" & Rec("Dimensioni").Value & ";" & Rec("Data").Value & ";" & Rec("idCategoria").Value & ";" & idMultimedia.ToString & ";§"
					Else
						Ritorno &= ";" & Rec("NomeFile").Value.ToString.Replace(";", "***PV***") & ";" & Rec("Dimensioni").Value & ";" & Rec("Data").Value & ";" & Rec("idCategoria").Value & ";" & idMultimedia.ToString & ";§"
					End If
				Else
					Ritorno = "ERROR: Nessun file rilevato"
					Exit For
				End If

				Rec.Close()
			Next
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

	'<WebMethod()>
	'Public Function EliminaMultimedia(idTipologia As String, idCategoria As String, idMultimedia As String) As String
	'	Dim gf As New GestioneFilesDirectory
	'	Dim Barra As String = "\"

	'	If TipoDB = "SQLSERVER" Then
	'		Barra = "\"
	'	Else
	'		Barra = "/"
	'	End If

	'	Dim Db As New clsGestioneDB(TipoDB)
	'	Dim Ritorno As String = ""
	'	Dim Sql As String
	'	Dim PathVideoInput As String = ""

	'	Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
	'	If ConnessioneSQL <> "" Then
	'		Dim Rec As Object
	'		Sql = "Select B.Categoria, B.Percorso, A.NomeFile From Dati A " &
	'			"Left Join Categorie B On A.idTipologia=B.idTipologia And A.idCategoria=B.idCategoria " &
	'			"Where A.idTipologia=" & idTipologia & " And B.idCategoria=" & idCategoria & " And Progressivo=" & idMultimedia
	'		Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
	'		If Rec.Eof = False Then
	'			PathVideoInput = Rec("Percorso").Value & Barra & Rec("NomeFile").Value
	'		End If
	'	End If

	'	Return Ritorno
	'End Function

	<WebMethod()>
	Public Function RitornaInformazioniConversione() As String
		Dim gf As New GestioneFilesDirectory
		Dim Ritorno As String = ""
		Dim NomeFile As String = Server.MapPath(".") & "/Logs/ffmpegout.txt"

		Ritorno = gf.LeggeFileIntero(NomeFile)

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function BloccaConvertiTuttiIVideo() As String
		StaEffettuandoConversioneAutomatica = False

		Return "*"
	End Function

	<WebMethod()>
	Public Function BloccaRitornaInformazioniImmagini() As String
		StaEffettuandoConversioneAutomaticaI = False

		Return "*"
	End Function

	<WebMethod()>
	Public Function SalvaMultimediaID(Progressivo As String, idTipologia As String, idCategoria As String, idMultimedia As String, Descrizione As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()

		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim Sql As String = ""
			Dim idProgressivo As Integer

			If Progressivo = "" Then
				Sql = "Select Coalesce(Max(Progressivo) + 1, 1) As Massimo From idsalvati"
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL, False)
				idProgressivo = Rec("Massimo").Value
				Rec.Close

				Sql = "Insert Into idsalvati Values (" & idProgressivo & ", " & idTipologia & ", " & idCategoria & ", " & idMultimedia & ", '" & Descrizione.Replace("'", "''") & "')"
			Else
				Sql = "Update idsalvati Set idTipologia=" & idTipologia & ", idCategoria=" & idCategoria & ", idMultimedia=" & idMultimedia & ", Descrizione='" & Descrizione.Replace("'", "''") & "' Where Progressivo=" & Progressivo
			End If

			Dim sRitorno As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
			If sRitorno <> "OK" Then
				Ritorno = "ERROR: " & sRitorno
			Else
				Ritorno = "*"
			End If
		Else
			Ritorno = StringaErrore & ": Connessione non valida"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaMultimediaID(Progressivo As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()

		If ConnessioneSQL <> "" Then
			Dim Sql As String

			Sql = "Delete From idsalvati Where Progressivo=" & Progressivo
			Dim sRitorno As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
			If sRitorno <> "OK" Then
				Ritorno = "ERROR: " & sRitorno
			Else
				Ritorno = "*"
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function CaricaMultimediaID(Progressivo As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()

		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim Sql As String

			Sql = "Select * From idsalvati Where Progressivo=" & Progressivo
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL, False)
			If Rec.Eof Then
				Ritorno = "ERROR: Nessun valore ritornato"
			Else
				Ritorno = Rec("idMultimedia").Value
			End If
			Rec.Close
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function CaricaListaMultimediaID(idTipologia As String, Categoria As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()

		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim Sql As String
			Dim idCategoria As String = ""

			If IsNumeric(Categoria) Then
				idCategoria = Val(Categoria)
			Else
				Sql = "Select * From categorie Where idTipologia=" & idTipologia & " And Categoria='" & Categoria.Replace("'", "''") & "'"
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL, False)
				If Rec.Eof Then
					Ritorno = "ERROR: Nessuna categoria ritornata"
				Else
					idCategoria = Rec("idCategoria").Value
				End If
				Rec.Close
			End If

			If idCategoria <> "" Then
				Sql = "Select * From idsalvati Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria & " Order By Descrizione"
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL, False)
				If Rec.Eof Then
					Ritorno = "ERROR: Nessun valore ritornato"
				Else
					Do Until Rec.Eof
						Ritorno &= Rec("Progressivo").Value & ";" & Rec("idMultimedia").Value & ";" & Rec("Descrizione").Value & "§"

						Rec.MoveNext
					Loop
				End If
				Rec.Close
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ConvertiTuttiIVideo() As String
		timerConv = New Timers.Timer(100)
		AddHandler timerConv.Elapsed, New ElapsedEventHandler(AddressOf ConverteTuttiIVideoThread)
		timerConv.Start()

		Return "*"
	End Function

	Private Sub ConverteTuttiIVideoThread()
		timerConv.Stop()

		Dim gf As New GestioneFilesDirectory
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Inizio As String = ""
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()

		If ConnessioneSQL <> "" Then
			Dim Rec As Object

			StaEffettuandoConversioneAutomatica = True
			StaEffettuandoConversioneAutomaticaFinale = True
			If gf.EsisteFile(Server.MapPath(".") & "/Logs/VideoDaConvertire.txt") Then
				gf.EliminaFileFisico(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")

				Dim Tutti As String = gf.LeggeFileIntero(Server.MapPath(".") & "/Logs/VideoDaConvertire.txt")
				Dim Video() As String = Tutti.Split(vbCrLf)

				For Each v As String In Video
					v = v.Replace(vbCrLf, "").Replace(Chr(13), "").Replace(Chr(10), "")

					'Inizio = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
					'gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
					'gf.ScriveTestoSuFileAperto(Inizio & ": Riga " & v & " Primo carattere: " & Mid(v, 1, 1))
					'gf.ChiudeFileDiTestoDopoScrittura()

					If Mid(v, 1, 1) <> "*" Then
						Dim vv() As String = v.Split(";")
						Dim idCategoria As String = vv(0)
						Dim idMultimedia As String = vv(1)

						Inizio = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
						Dim Esegui As Boolean = False

						'gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
						'gf.ScriveTestoSuFileAperto(Inizio & ": Inizio")
						'gf.ChiudeFileDiTestoDopoScrittura()

						Dim Sql As String = "Select * From InformazioniVideo " &
							"Where idTipologia=2 And idCategoria=" & idCategoria & " And idMultimedia=" & idMultimedia
						Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL, False)
						If Rec.Eof = True Then
							Esegui = True
						Else
							Dim Infos As String = Rec("Jsone").Value.ToString
							If Infos.ToUpper.Contains("H264") Then
								gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
								gf.ScriveTestoSuFileAperto(Inizio & ": Conversione video idCategoria " & idCategoria & " idMultimedia " & idMultimedia & " -> Video già convertito")
								gf.ChiudeFileDiTestoDopoScrittura()

								Esegui = False
							Else
								Esegui = True
							End If
						End If
						Rec.Close

						If Esegui Then
							Inizio = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
							gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
							gf.ScriveTestoSuFileAperto(Inizio & ": Inizio Conversione video idCategoria " & idCategoria & " idMultimedia " & idMultimedia)
							gf.ChiudeFileDiTestoDopoScrittura()

							Dim Rit As String = ConverteVideo(2, idCategoria, idMultimedia, "N")

							Inizio = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

							gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
							gf.ScriveTestoSuFileAperto(Inizio & ": Ritorno informazioni video idCategoria. " & Rit)
							gf.ChiudeFileDiTestoDopoScrittura()

							If Rit.Contains("ERROR:") Then
								'Inizio = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

								'gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
								'gf.ScriveTestoSuFileAperto(Inizio & ": Ritorno informazioni video idCategoria. " & Rit)
								'gf.ChiudeFileDiTestoDopoScrittura()
							Else
								Inizio = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

								gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
								gf.ScriveTestoSuFileAperto(Inizio & ": Ritorno informazioni video idCategoria " & idCategoria & " idMultimedia " & idMultimedia)
								gf.ChiudeFileDiTestoDopoScrittura()

								RitornaInformazioniVideo(2, idCategoria, idMultimedia, "S")

								Inizio = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

								gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
								gf.ScriveTestoSuFileAperto(Inizio & ": Fine Conversione video idCategoria " & idCategoria & " idMultimedia " & idMultimedia & " -> " & Rit)
								gf.ChiudeFileDiTestoDopoScrittura()
							End If
						End If

						If StaEffettuandoConversioneAutomatica = False Then
							StaEffettuandoConversioneAutomaticaFinale = False

							Inizio = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

							gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
							gf.ScriveTestoSuFileAperto(Inizio & ": Conversione Bloccata Da Sito WEB")
							gf.ChiudeFileDiTestoDopoScrittura()

							Exit For
						End If
					Else
						Dim vv() As String = v.Replace("*", "").Split(";")
						Dim idCategoria As String = vv(0)
						Dim idMultimedia As String = vv(1)

						Inizio = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

						gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
						gf.ScriveTestoSuFileAperto(Inizio & ": Skippato da asterisco. idCategoria: " & idCategoria & " idMultimedia: " & idMultimedia)
						gf.ChiudeFileDiTestoDopoScrittura()
					End If

					gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
					gf.ScriveTestoSuFileAperto("---------------------------------------------------------------------------------------")
					gf.ChiudeFileDiTestoDopoScrittura()
				Next

				Ritorno = "*"
			Else
				Ritorno = "ERROR: Prima effettuare un refresh delle informazioni video"
			End If
		End If

		StaEffettuandoConversioneAutomatica = False
		StaEffettuandoConversioneAutomaticaFinale = False

		Inizio = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

		gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
		gf.ScriveTestoSuFileAperto(Inizio & ": " & Ritorno)
		gf.ChiudeFileDiTestoDopoScrittura()
		'Return Ritorno
	End Sub

	<WebMethod()>
	Public Function RefreshTutteInformazioniVideo(ForzaRefresh As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim Barra As String = "\"
			Dim Sql As String
			Dim DaFare As New List(Of String)

			If TipoDB = "SQLSERVER" Then
				Barra = "\"
			Else
				Barra = "/"
			End If

			gf.EliminaFileFisico(Server.MapPath(".") & "/Logs/ConvertiTutto.txt")
			gf.EliminaFileFisico(Server.MapPath(".") & "/Logs/VideoDaConvertire.txt")

			Dim Elaborati As Integer = 0
			Dim Errori As Integer = 0
			Dim Skippati As Integer = 0
			Dim DaConvertire As Integer = 0

			Sql = "Select * From Dati A Where A.idTipologia=2"
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			If Rec.Eof = False Then
				Do Until Rec.Eof
					gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConvertiTutto.txt")
					gf.ScriveTestoSuFileAperto("Elaborazione idCategoria: " & Rec("idCategoria").Value & " idMultimedia: " & Rec("Progressivo").Value & " -> " & Rec("NomeFile").Value)

					If Rec("NomeFile").Value.ToString.Contains("_CONV.mp4") Then
						gf.ScriveTestoSuFileAperto("             Skippo controllo... Già convertito")
						Skippati += 1
					Else
						Ritorno = RitornaInformazioniVideo(2, Rec("idCategoria").Value, Rec("Progressivo").Value, ForzaRefresh)
						If Ritorno.Contains("ERROR:") Then
							gf.ScriveTestoSuFileAperto("             " & Ritorno)
							Errori += 1
						Else
							If Not Ritorno.ToUpper.Contains("H264") Then
								DaFare.Add(Rec("idCategoria").Value & ";" & Rec("Progressivo").Value)
								gf.ScriveTestoSuFileAperto("             Aggiunto alla lista da convertire")
								DaConvertire += 1
							End If
						End If
					End If
					gf.ChiudeFileDiTestoDopoScrittura()
					Elaborati += 1

					Rec.MoveNext
				Loop
				Rec.Close

				gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/VideoDaConvertire.txt")
				For Each s As String In DaFare
					If s <> "" Then
						gf.ScriveTestoSuFileAperto(s)
					End If
				Next
				gf.ChiudeFileDiTestoDopoScrittura()

				Ritorno = "File elaborati: " & Elaborati & " - Skippati: " & Skippati & " - Da Convertire: " & DaConvertire & " - Errori: " & Errori
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function FinisceConversioneVideo(idTipologia As String, idCategoria As String, idMultimedia As String, SoloRitorno As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Sql As String

		Dim NomeNuovo As String = ""
		Dim BytesVecchi As Long
		Dim BytesNuovi As Long

		If SoloRitorno = "N" Then
			Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
			If ConnessioneSQL <> "" Then
				Dim Rec As Object
				Dim PathVideoInput As String = ""
				Dim PathVideoOutput As String = ""
				Dim Barra As String = "\"

				If TipoDB = "SQLSERVER" Then
					Barra = "\"
				Else
					Barra = "/"
				End If

				Sql = "Select B.Categoria, B.Percorso, A.NomeFile From Dati A " &
					"Left Join Categorie B On A.idTipologia=B.idTipologia And A.idCategoria=B.idCategoria " &
					"Where A.idTipologia=" & idTipologia & " And B.idCategoria=" & idCategoria & " And Progressivo=" & idMultimedia
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				If Rec.Eof = False Then
					PathVideoInput = Rec("Percorso").Value & Barra & Rec("NomeFile").Value
					Dim Estensione As String = gf.TornaEstensioneFileDaPath(Rec("NomeFile").Value)
					NomeNuovo = Rec("NomeFile").Value.replace(Estensione, "")
					If Not NomeNuovo.Contains("_CONV") Then
						NomeNuovo &= "_CONV"
					End If
					If Not NomeNuovo.Contains(".mp4") Then
						NomeNuovo &= ".mp4"
					End If
					'For i As Integer = Len(NomeNuovo) To 1 Step -1
					'	If Mid(NomeNuovo, i, 1) = Barra Then
					'		NomeNuovo = Mid(NomeNuovo, i + 1, NomeNuovo.Length)
					'	End If
					'Next
					PathVideoOutput = Rec("Percorso").Value & Barra & Rec("NomeFile").Value.replace(Estensione, "")
					If Not PathVideoOutput.Contains("_CONV") Then
						PathVideoOutput &= "_CONV"
					End If
					If Not PathVideoOutput.Contains(".mp4") Then
						PathVideoOutput &= ".mp4"
					End If

					If DaWebGlobale = True Then
						Dim Inizio2 As String = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
						gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
						gf.ScriveTestoSuFileAperto(Inizio2 & ": PathVideoInput: " & PathVideoInput)
						gf.ScriveTestoSuFileAperto(Inizio2 & ": PathVideoOutput: " & PathVideoOutput)
						gf.ChiudeFileDiTestoDopoScrittura()
					End If

					'NomeNuovo = Rec("Percorso").Value & Barra & NomeNuovo

					Sql = "Update dati Set NomeFile='" & NomeNuovo.Replace("'", "''") & "' Where idCategoria=" & idCategoria & " And idTipologia=" & idTipologia & " And Progressivo=" & idMultimedia
					Dim sRitorno As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
					If sRitorno <> "OK" Then
						Return "ERROR: " & sRitorno
					Else
						If DaWebGlobale = True Then
							Dim Inizio2 As String = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
							gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
							gf.ScriveTestoSuFileAperto(Inizio2 & ": Impostato Nome File su tabella: " & NomeNuovo)
							gf.ChiudeFileDiTestoDopoScrittura()
						End If

						Sql = "Update InformazioniVideo Set Convertito = 'S' " &
							"Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria & " And idMultimedia=" & idMultimedia
						Dim rit As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

						BytesVecchi = gf.TornaDimensioneFile(PathVideoInput)
						BytesNuovi = gf.TornaDimensioneFile(PathVideoOutput)
						gf.EliminaFileFisico(PathVideoInput)

						If DaWebGlobale = True Then
							Dim Inizio2 As String = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
							gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
							gf.ScriveTestoSuFileAperto(Inizio2 & ": Impostato valore convertito su tabella per idTipologia " & idTipologia & " idCategoria " & idCategoria & " idMultimedia " & idMultimedia)
							gf.ChiudeFileDiTestoDopoScrittura()
						End If

						nomeNuovoGlobale = NomeNuovo
						bytesVecchiGlobale = BytesVecchi
						bytesNuoviGlobale = BytesNuovi

						' Dim rit As String = gf.CreaAggiornaFile(Server.MapPath(".") & "/Logs/FiniscoConversione.txt", "RINIOMINA: " & nomeNuovoGlobale & ";" & bytesVecchiGlobale & ";" & bytesNuoviGlobale)

						Dim NomeFile As String = Server.MapPath(".") & "/Logs/ffmpegout.txt"
						gf.EliminaFileFisico(NomeFile)

						gf.EliminaFileFisico(Server.MapPath(".") & "/Logs/FinitaConversione.txt")
					End If
				End If
			End If
		End If

		If DaWebGlobale = True Then
			Dim Inizio2 As String = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
			gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
			gf.ScriveTestoSuFileAperto(Inizio2 & ": " & NomeNuovo & " Bytes Vecchi: " & BytesVecchi & " Bytes Nuovi: " & BytesNuovi)
			gf.ChiudeFileDiTestoDopoScrittura()
		End If

		Return NomeNuovo & ";" & BytesVecchi & ";" & BytesNuovi
	End Function

	<WebMethod()>
	Public Function RitornaInformazioniVideo(idTipologia As String, idCategoria As String, idMultimedia As String, Refresh As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim Barra As String = "\"

		If TipoDB = "SQLSERVER" Then
			Barra = "\"
		Else
			Barra = "/"
		End If

		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Sql As String

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Sql = "Select B.Categoria, B.Percorso, A.NomeFile From Dati A " &
				"Left Join Categorie B On A.idTipologia=B.idTipologia And A.idCategoria=B.idCategoria " &
				"Where A.idTipologia=" & idTipologia & " And B.idCategoria=" & idCategoria & " And Progressivo=" & idMultimedia
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			If Rec.Eof = False Then
				PathVideoInput = Rec("Percorso").Value & Barra & Rec("NomeFile").Value
				If gf.EsisteFile(PathVideoInput) = False Then
					If PathVideoInput.Contains("_CONV") Then
						PathVideoInput = PathVideoInput.Replace("_CONV", "")
					Else
						PathVideoInput = Rec("Percorso").Value & Barra & Rec("NomeFile").Value
						Dim Estensione As String = gf.TornaEstensioneFileDaPath(Rec("NomeFile").Value)
						PathVideoInput = Rec("Percorso").Value & Barra & Rec("NomeFile").Value.replace(Estensione, "") & "_CONV.mp4"
					End If
					If gf.EsisteFile(PathVideoInput) = False Then
						Return "ERROR: Nessun file video fisico rilevato: " & PathVideoInput
					End If
				End If
			Else
				Return "ERROR: Nessun file video rilevato"
			End If
			Rec.Close

			Sql = "Select * From InformazioniVideo " &
				"Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria & " And idMultimedia=" & idMultimedia
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL, False)
			If Rec.Eof = True Or Refresh = "S" Then
				Dim processoFFMpeg2 As Process = New Process()
				Dim pi As ProcessStartInfo = New ProcessStartInfo()
				Dim Comando As String = ""

				Comando = "ffprobe"
				pi.Arguments = "-loglevel 0 -print_format json -show_format -show_streams """ & PathVideoInput & """"

				pi.FileName = Comando
				pi.WindowStyle = ProcessWindowStyle.Normal
				processoFFMpeg2.StartInfo = pi
				processoFFMpeg2.StartInfo.UseShellExecute = False
				processoFFMpeg2.StartInfo.RedirectStandardOutput = True
				processoFFMpeg2.StartInfo.RedirectStandardError = True
				processoFFMpeg2.Start()

				Dim ffReader As StreamReader
				Dim ffReader2 As StreamReader
				Dim strFFOUT As String = "INIZIO"
				Dim strFFOUT2 As String = "INIZIO"

				ffReader = processoFFMpeg2.StandardError
				ffReader2 = processoFFMpeg2.StandardOutput

				Do
					strFFOUT = ffReader.ReadLine
					strFFOUT2 = ffReader2.ReadLine

					'If strFFOUT2 <> "" Then
					Ritorno &= strFFOUT2
					'End If
				Loop Until processoFFMpeg2.HasExited And strFFOUT2 = Nothing Or strFFOUT2 = ""

				If Ritorno = "{" Then Ritorno = "{}"

				Sql = "Delete From InformazioniVideo Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria & " And idMultimedia=" & idMultimedia
				Dim rit As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

				Sql = "Insert Into InformazioniVideo Values (" & idTipologia & ", " & idCategoria & ", " & idMultimedia & ", 'N', '" & Ritorno.Replace(" ", "") & "', 'N')"
				rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

				Ritorno = "N;" & Ritorno
			Else
				Ritorno = Rec("Convertito").Value & ";" & Rec("Jsone").Value & ";"
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ConverteVideo(idTipologia As String, idCategoria As String, idMultimedia As String, DaWeb As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim Barra As String = "\"

		If TipoDB = "SQLSERVER" Then
			Barra = "\"
		Else
			Barra = "/"
		End If

		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Sql As String
		Dim NomeNuovo As String = ""
		Dim BytesVecchi As Long = -1
		Dim BytesNuovi As Long = -1

		idTipologiaGlobale = idTipologia
		idCategoriaGlobale = idCategoria
		idMultimediaGlobale = idMultimedia

		Dim NomeFile As String = Server.MapPath(".") & "/Logs/ffmpegout.txt"
		If gf.EsisteFile(NomeFile) Then
			Return "ERROR: Conversione in corso"
		End If

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim Inizio2 As String = ""

			Sql = "Delete From InformazioniVideo " &
				"Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria & " And idMultimedia=" & idMultimedia

			'Inizio2 = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
			'gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
			'gf.ScriveTestoSuFileAperto(Inizio2 & ": " & Sql)
			'gf.ChiudeFileDiTestoDopoScrittura()

			Dim rit As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

			Sql = "Select B.Categoria, B.Percorso, A.NomeFile From Dati A " &
				"Left Join Categorie B On A.idTipologia=B.idTipologia And A.idCategoria=B.idCategoria " &
				"Where A.idTipologia=" & idTipologia & " And B.idCategoria=" & idCategoria & " And Progressivo=" & idMultimedia

			'Inizio2 = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
			'gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
			'gf.ScriveTestoSuFileAperto(Inizio2 & ": " & Sql)
			'gf.ChiudeFileDiTestoDopoScrittura()

			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			If Rec.Eof = False Then
				PathVideoInput = Rec("Percorso").Value & Barra & Rec("NomeFile").Value
				NomeFileDaConvertire = Rec("NomeFile").Value

				'Inizio2 = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
				'gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
				'gf.ScriveTestoSuFileAperto(Inizio2 & ": Nome File " & Rec("NomeFile").Value)
				'gf.ChiudeFileDiTestoDopoScrittura()

				Dim Estensione As String = gf.TornaEstensioneFileDaPath(Rec("NomeFile").Value)

				'Inizio2 = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
				'gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
				'gf.ScriveTestoSuFileAperto(Inizio2 & ": Estensione " & Estensione)
				'gf.ChiudeFileDiTestoDopoScrittura()

				If Estensione = "" Then
					Inizio2 = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
					gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
					gf.ScriveTestoSuFileAperto(Inizio2 & ": Nome File Non Valido " & Rec("NomeFile").Value)
					gf.ChiudeFileDiTestoDopoScrittura()

					Return "ERROR: Nome file non valido: " & Rec("NomeFile").Value
				End If

				NomeNuovo = Rec("NomeFile").Value.replace(Estensione, "") & "_CONV.mp4"
				NomeNuovo = Rec("NomeFile").Value.replace(Estensione, "") & ".mp4_CONV.mp4"
				For i As Integer = Len(NomeNuovo) To 1 Step -1
					If Mid(NomeNuovo, i, 1) = Barra Then
						NomeNuovo = Mid(NomeNuovo, i + 1, NomeNuovo.Length)
					End If
				Next

				'Inizio2 = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
				'gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
				'gf.ScriveTestoSuFileAperto(Inizio2 & ": Nome Nuovo " & NomeNuovo)
				'gf.ChiudeFileDiTestoDopoScrittura()

				If Not Rec("NomeFile").Value.ToString.Contains("_CONV") Then
					PathVideoOutput = Rec("Percorso").Value & Barra & Rec("NomeFile").Value.replace(Estensione, "") & "_CONV.mp4"
				Else
					PathVideoOutput = Rec("Percorso").Value & Barra & Rec("NomeFile").Value.replace(Estensione, "") & ".mp4"
				End If

				'Inizio2 = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
				'gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
				'gf.ScriveTestoSuFileAperto(Inizio2 & ": Path Video Output " & PathVideoOutput)
				'gf.ChiudeFileDiTestoDopoScrittura()

				Rec.Close

				'Inizio2 = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
				'gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
				'gf.ScriveTestoSuFileAperto(Inizio2 & ": File input: " & PathVideoInput & " - File Output: " & PathVideoOutput & " - NomeNuovo: " & NomeNuovo & " - File esistente: " & gf.EsisteFile(PathVideoInput))
				'gf.ChiudeFileDiTestoDopoScrittura()

				'Return "ERROR: File input: " & PathVideoInput & " - File Output: " & PathVideoOutput & " - NomeNuovo: " & NomeNuovo

				If gf.EsisteFile(PathVideoInput) Then
					gf.EliminaFileFisico(Server.MapPath(".") & "/Logs/ffmpegout.txt")
					gf.EliminaFileFisico(Server.MapPath(".") & "/Logs/frames1.txt")
					gf.EliminaFileFisico(Server.MapPath(".") & "/Logs/frames2.txt")
					' gf.EliminaFileFisico(Server.MapPath(".") & "/FiniscoConversione.txt")
					gf.EliminaFileFisico(PathVideoOutput)

					' Ritorna numero frames
					Dim processoFFMpeg2 As Process = New Process()
					Dim pi As ProcessStartInfo = New ProcessStartInfo()
					Dim Comando As String = ""

					Comando = "ffprobe"
					' 
					pi.Arguments = "-v error -select_streams v:0 -count_frames -show_entries stream=nb_read_frames -print_format csv """ & PathVideoInput & """"

					'If DaWebGlobale = True Then
					Inizio2 = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

					gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
					gf.ScriveTestoSuFileAperto(Inizio2 & ": " & Comando & " " & pi.Arguments)
					gf.ChiudeFileDiTestoDopoScrittura()
					'End If

					pi.FileName = Comando
					pi.WindowStyle = ProcessWindowStyle.Normal
					processoFFMpeg2.StartInfo = pi
					processoFFMpeg2.StartInfo.UseShellExecute = False
					processoFFMpeg2.StartInfo.RedirectStandardOutput = True
					processoFFMpeg2.StartInfo.RedirectStandardError = True
					processoFFMpeg2.Start()

					Dim ffReader As StreamReader
					Dim ffReader2 As StreamReader
					Dim strFFOUT As String = "INIZIO"
					Dim strFFOUT2 As String = "INIZIO"

					NumeroFrames = ""
					ffReader = processoFFMpeg2.StandardError
					ffReader2 = processoFFMpeg2.StandardOutput

					Inizio2 = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

					gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
					gf.ScriveTestoSuFileAperto(Inizio2 & ": Lettura frames")
					gf.ChiudeFileDiTestoDopoScrittura()

					Try
						'If DaWebGlobale = True Then
						'End If

						'Dim Secondi As Integer = 0
						'Dim SecondiAttuali As Integer = Now.Second

						Do
							'Dim SecondiAttuali2 As Integer = Now.Second
							'If SecondiAttuali <> SecondiAttuali2 Then
							'	SecondiAttuali = SecondiAttuali2

							'	Secondi += 1

							'	gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
							'	gf.ScriveTestoSuFileAperto(Inizio2 & ": Lettura frames. Secondi: " & Secondi)
							'	gf.ChiudeFileDiTestoDopoScrittura()

							'	If Secondi > 30 Then
							'		processoFFMpeg2.Kill()
							'		processoFFMpeg2.Dispose()

							'		Return "ERROR: Uscito dal ciclo di attesa per lettura frames. Troppi secondi"
							'	End If
							'End If

							strFFOUT = ffReader.ReadLine
							strFFOUT2 = ffReader2.ReadLine

							gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/frames1.txt")
							gf.ScriveTestoSuFileAperto(strFFOUT)
							gf.ChiudeFileDiTestoDopoScrittura()

							gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/frames2.txt")
							gf.ScriveTestoSuFileAperto(strFFOUT2)
							gf.ChiudeFileDiTestoDopoScrittura()

							'gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
							'gf.ScriveTestoSuFileAperto(Inizio2 & ": Lettura frames 4. Out ---" & strFFOUT & "---")
							'gf.ChiudeFileDiTestoDopoScrittura()

							If strFFOUT <> Nothing And strFFOUT <> "" Then
								If strFFOUT.Contains("moov atom not found") Then
									NumeroFrames = ""

									Inizio2 = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

									gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
									gf.ScriveTestoSuFileAperto(Inizio2 & ": Rilevato moov atom not found")
									gf.ChiudeFileDiTestoDopoScrittura()

									Exit Do
								End If
							End If

							'gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
							'gf.ScriveTestoSuFileAperto(Inizio2 & ": Lettura frames 5. Out " & strFFOUT2)
							'gf.ChiudeFileDiTestoDopoScrittura()

							If strFFOUT2 <> Nothing And strFFOUT2 <> "" Then
								If strFFOUT2.Contains("stream,") Then         'if the strFFOut contains the string
									NumeroFrames = strFFOUT2.Replace("stream,", "")

									Inizio2 = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

									gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
									gf.ScriveTestoSuFileAperto(Inizio2 & ": Rilevato stream 1. Numero Frames: " & NumeroFrames)
									gf.ChiudeFileDiTestoDopoScrittura()

									Exit Do
								End If
							End If

							'gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
							'gf.ScriveTestoSuFileAperto(Inizio2 & ": Lettura frames 6. Out " & strFFOUT2)
							'gf.ChiudeFileDiTestoDopoScrittura()

							If strFFOUT <> Nothing And strFFOUT <> "" Then
								If strFFOUT.Contains("stream,") Then         'if the strFFOut contains the string
									NumeroFrames = strFFOUT.Replace("stream,", "")

									Inizio2 = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

									gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
									gf.ScriveTestoSuFileAperto(Inizio2 & ": Rilevato stream 2. Numero Frames: " & NumeroFrames)
									gf.ChiudeFileDiTestoDopoScrittura()

									Exit Do
								End If
							End If

							''If strFFOUT2.Contains("[/FRAME]") Then         'if the strFFOut contains the string
							''	Exit Do
							''End If
							'gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
							'gf.ScriveTestoSuFileAperto(Inizio2 & ": Lettura frames 7. Out " & strFFOUT)
							'gf.ChiudeFileDiTestoDopoScrittura()
						Loop Until strFFOUT = Nothing Or strFFOUT = ""
					Catch ex As Exception
						NumeroFrames = ""

						Inizio2 = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

						gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
						gf.ScriveTestoSuFileAperto(Inizio2 & ": Errore su lettura frames: " & ex.Message)
						gf.ChiudeFileDiTestoDopoScrittura()

						Return "ERROR: " & ex.Message
					End Try
					' Ritorna numero frames

					'If DaWebGlobale = True Then
					Inizio2 = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

					gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
					gf.ScriveTestoSuFileAperto(Inizio2 & ": Letto numero frames: " & NumeroFrames)
					gf.ChiudeFileDiTestoDopoScrittura()
					'End If

					gf.EliminaFileFisico(Server.MapPath(".") & "/Logs/frames1.txt")
					gf.EliminaFileFisico(Server.MapPath(".") & "/Logs/frames2.txt")

					If NumeroFrames = "" Then
						Return "ERROR: Non riesco a rilevare il numero dei frames"
					End If
					'Return "FRAMES: " & NumeroFrames

					processoFFMpeg = New Process()
					pi = New ProcessStartInfo()

					' -an -i input.mov -vcodec libx264 -pix_fmt yuv420p -profile:v baseline -level 3 output2.mp4
					pi.Arguments = "-an -i """ & PathVideoInput & """ -vcodec libx264 -pix_fmt yuv420p -profile:v baseline -level 3 """ & PathVideoOutput & """" ' > " & Server.MapPath(".") & "/err.txt 2> " & Server.MapPath(".") & "/ffmpegout.txt"

					If TipoDB <> "SQLSERVER" Then
						Comando = "ffmpeg"
					Else
						Comando = Server.MapPath(".") & "\ffmpeg.exe"
					End If

					pi.FileName = Comando
					pi.WindowStyle = ProcessWindowStyle.Normal
					processoFFMpeg.StartInfo = pi
					processoFFMpeg.StartInfo.UseShellExecute = False
					processoFFMpeg.StartInfo.RedirectStandardOutput = True
					processoFFMpeg.StartInfo.RedirectStandardError = True
					processoFFMpeg.Start()

					'Dim OutPutP As String = processoFFMpeg.StandardOutput.ReadToEnd()
					'Ritorno = OutPutP & "****"
					'Dim Err As String = processoFFMpeg.StandardError.ReadToEnd()
					'Ritorno &= Err & "*****"

					' Return ritorno

					'Dim timerLog As Timers.Timer = Nothing

					If DaWeb = "S" Then
						DaWebGlobale = False
						timerLog = New Timers.Timer(100)
						AddHandler timerLog.Elapsed, New ElapsedEventHandler(AddressOf ContinuaConversione)
						timerLog.Start()
					Else
						DaWebGlobale = True

						If DaWebGlobale = True Then
							Inizio2 = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

							gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
							gf.ScriveTestoSuFileAperto(Inizio2 & ": Attesa termine conversione")
							gf.ChiudeFileDiTestoDopoScrittura()
						End If

						ContinuaConversione()

						'Dim Inizio2 As String = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

						'gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
						'gf.ScriveTestoSuFileAperto(Inizio2 & ": Uscita ciclo di attesa fine conversione 2")
						'gf.ChiudeFileDiTestoDopoScrittura()

						If DaWebGlobale = True Then
							Inizio2 = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

							gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
							gf.ScriveTestoSuFileAperto(Inizio2 & ": Terminata conversione")
							gf.ChiudeFileDiTestoDopoScrittura()
						End If

						Ritorno = FinisceConversioneVideo(idTipologia, idCategoria, idMultimedia, "N")

						Return Ritorno
					End If

					' processoFFMpeg.WaitForExit()

					'Sql = "Update dati Set NomeFile='" & NomeNuovo.Replace("'", "''") & "' Where idCategoria=" & idCategoria & " And idTipologia=" & idTipologia & " And Progressivo=" & idMultimedia
					'Dim sRitorno As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
					'If sRitorno <> "OK" Then
					'	Return "ERROR: " & sRitorno
					'Else
					'	BytesVecchi = gf.TornaDimensioneFile(PathVideoInput)
					'	BytesNuovi = gf.TornaDimensioneFile(PathVideoOutput)
					'	gf.EliminaFileFisico(PathVideoInput)
					'End If
				Else
					Inizio2 = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
					gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
					gf.ScriveTestoSuFileAperto(Inizio2 & ": File non esistente: " & PathVideoInput & " - File Output: " & PathVideoOutput & " - NomeNuovo: " & NomeNuovo)
					gf.ChiudeFileDiTestoDopoScrittura()

					Return "ERROR: File non rilevato -> " & PathVideoInput
				End If
			Else
				Return "ERROR: Nessun multimedia rilevato"
			End If
		Else
			Return "ERROR: Connessione non valida"
		End If

		Return "*"
	End Function

	Private Sub ContinuaConversione()
		timerLog.Enabled = False

		Dim gf As New GestioneFilesDirectory
		Dim strFFOUT As String = "INIZIO"
		Dim Inizio As String = Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
		Dim InizioData As DateTime = Now

		Dim ffReader As StreamReader
		Dim currentFramestr As String = ""
		Dim currentFrameInt As Integer = -1

		ffReader = processoFFMpeg.StandardError

		Dim Minuti2 As Integer = Now.Minute

		Try
			Do
				strFFOUT = ffReader.ReadLine
				If strFFOUT.Contains("frame=") Then         'if the strFFOut contains the string
					currentFramestr = Mid(strFFOUT, 7, 6)   'grab the next part after the string 'frame='
					currentFrameInt = CInt(currentFramestr) 'convert the string back to an integer
				End If

				Dim dime1 As Long = gf.TornaDimensioneFile(PathVideoInput)
				Dim dime2 As Long = gf.TornaDimensioneFile(PathVideoOutput)
				Dim sDime1 As String = ""
				Dim sDime2 As String = ""

				If dime1 > 1024 * 1024 * 1024 Then
					sDime1 = (CInt((dime1 / 1024 / 1024 / 1024) * 100) / 100) & " Gb."
				Else

					If dime1 > 1024 * 1024 Then
						sDime1 = (CInt((dime1 / 1024 / 1024) * 100) / 100) & " Mb."
					Else
						If dime1 > 1024 Then
							sDime1 = (CInt((dime1 / 1024) * 100) / 100) & " Kb."
						Else
							sDime1 = dime1 & " B."
						End If
					End If
				End If

				If dime2 > 1024 * 1024 * 1024 Then
					sDime2 = (CInt((dime2 / 1024 / 1024 / 1024) * 100) / 100) & " Gb."
				Else
					If dime2 > 1024 * 1024 Then
						sDime2 = (CInt((dime2 / 1024 / 1024) * 100) / 100) & " Mb."
					Else
						If dime2 > 1024 Then
							sDime2 = (CInt((dime2 / 1024) * 100) / 100) & " Kb."
						Else
							sDime2 = dime2 & " B."
						End If
					End If
				End If

				Dim perc As Single = CInt((dime2 / dime1) * 100) ' / 100
				Dim Scritta As String = ""

				Dim differenza As Integer = DateDiff("s", InizioData, Now)
				Dim Ore As Integer = 0
				Dim Minuti As Integer = 0
				Dim Secondi As Integer = 0
				While differenza > 59
					differenza -= 60
					Minuti += 1
					If Minuti > 59 Then
						Minuti = 0
						Ore += 1
					End If
				End While
				Secondi = differenza
				Dim DataDaStampare As String = Format(Ore, "00") & ":" & Format(Minuti, "00") & ":" & Format(Secondi, "00")

				Scritta = NomeFileDaConvertire & ";"
				Scritta &= "Ora inizio UTC: " & Inizio & " - Tempo Impiegato: " & DataDaStampare & ";"
				If currentFrameInt <> -1 Then
					Dim perc2 As Single = CInt((currentFrameInt / Val(NumeroFrames)) * 100) ' / 100

					Scritta &= "Frame: " & currentFrameInt & "/" & NumeroFrames & " " & perc2 & "%;"
				End If
				'Scritta &= strFFOUT & ";"
				Scritta &= "Dimensione Origine: " & sDime1 & ";"
				Scritta &= "Dimensione Destinazione: " & sDime2 & " " & perc & "%;"
				' Scritta &= "Perc.: " & perc & "%"

				If DaWebGlobale = True Then
					Dim NuoviMinuti As Integer = Now.Minute

					If NuoviMinuti <> Minuti2 Then
						Minuti2 = NuoviMinuti
						If Not Scritta Is Nothing And Scritta <> "" Then
							Dim Inizio2 As String = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

							gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
							gf.ScriveTestoSuFileAperto(Inizio2 & ": " & Scritta.Replace(vbCrLf, "").Trim)
							gf.ChiudeFileDiTestoDopoScrittura()
						End If
					End If
				End If

				Dim rit2 As String = gf.CreaAggiornaFile(Server.MapPath(".") & "/Logs/ffmpegout.txt", Scritta)
				If rit2.Contains(StringaErrore) Then
					Exit Do
				End If

				'If strFFOUT.Contains("[libx264") Then
				'	Exit Do
				'End If
				If currentFrameInt > Val(NumeroFrames) Then
					Exit Do
				End If

				'Dim Fine As Boolean = False
				'If Not processoFFMpeg Is Nothing Then
				'	Fine = processoFFMpeg.HasExited
				'End If
			Loop Until strFFOUT = Nothing Or strFFOUT = ""
		Catch ex As Exception
			Dim rit2 As String = gf.CreaAggiornaFile(Server.MapPath(".") & "/Logs/ffmpegout.txt", ex.Message)
		End Try

		If DaWebGlobale = True Then
			Dim Inizio2 As String = Now.Year & "/" & Format(Now.Month, "00") & "/" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

			gf.ApreFileDiTestoPerScrittura(Server.MapPath(".") & "/Logs/ConversioneAutomatica.txt")
			gf.ScriveTestoSuFileAperto(Inizio2 & ": Uscita ciclo di attesa fine conversione")
			gf.ChiudeFileDiTestoDopoScrittura()
		End If

		Dim Tutto As String = ""

		'gf.EliminaFileFisico(Server.MapPath(".") & "/Logs/ffmpegout.txt")

		If gf.EsisteFile(Server.MapPath(".") & "/Logs/FinitaConversione.txt") Then
			Tutto = gf.LeggeFileIntero(Server.MapPath(".") & "/Logs/FinitaConversione.txt")
			gf.EliminaFileFisico(Server.MapPath(".") & "/Logs/FinitaConversione.txt")
		End If
		Tutto &= idTipologiaGlobale & "*" & idCategoriaGlobale & "*" & idMultimediaGlobale & "§"
		Dim rit As String = gf.CreaAggiornaFile(Server.MapPath(".") & "/Logs/FinitaConversione.txt", Tutto)
		'FinisceConversioneVideo(idTipologiaGlobale, idCategoriaGlobale, idMultimediaGlobale, "N")
		' processoFFMpeg.WaitForExit()
	End Sub

	Private Function CreaThumbDaVideo(Categoria As String, Percorso As String, Video As String, Conversione As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim PathBase As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PercorsoThumbs.txt")
		PathBase = PathBase.Replace(vbCrLf, "")

		Dim Barra As String = "\"

		If TipoDB = "SQLSERVER" Then
			Barra = "\"
		Else
			Barra = "/"
		End If

		If Strings.Right(PathBase, 1) <> Barra Then
			PathBase &= Barra
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

	Public Function SistemaPercorso(pathPassato As String) As String
		Dim pp As String = pathPassato

		pp = pp.Replace(vbCrLf, "").Trim()
		If Strings.Right(pp, 1) = "\" Or Strings.Right(pp, 1) = "/" Then
			pp = Mid(pp, 1, pp.Length - 1)
		End If

		Return pp
	End Function

	Public Sub ScriveLog(MP As String, Squadra As String, NomeFile As String, Cosa As String)
		If Not effettuaLog Then
			Return
		End If

		If Squadra = "" Then
			Squadra = "NessunaSquadra"
		End If

		Dim gf As New GestioneFilesDirectory

		'If nomeFileLogMail = "" Then

		Dim nomeFileLog As String = Server.MapPath(".") & "\" & NomeFile & ".txt"
		gf.CreaDirectoryDaPercorso(nomeFileLog)
		'End If

		Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

		gf.ApreFileDiTestoPerScrittura(nomeFileLog)
		gf.ScriveTestoSuFileAperto(Datella & " " & Cosa)
		gf.ChiudeFileDiTestoDopoScrittura()

		gf = Nothing
	End Sub

	<WebMethod()>
	Public Function RitornaFilesNuovo(idTipologia As String, Categoria As String) As String
		If StaLeggendoImmagini Then
			Return "ERROR: Sto già caricando multimedia"
		End If

		If Categoria = "" Or Categoria.ToUpper = "TUTTE" Then
			RitornaFiles()
		End If

		StaLeggendoImmagini = True

		Dim gf As New GestioneFilesDirectory

		Dim NomeFileLog As String = Server.MapPath(".") & "\Logs\Log_" & dataAttuale() & ".txt"
		gf.ApreFileDiTestoPerScrittura(NomeFileLog)

		Try
			Dim sPathsVideo As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PercorsiVideo.txt")
			Dim sPathsImm As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PercorsiImmagini.txt")
			Dim PathVideo() As String = sPathsVideo.Split("§")
			Dim PathImmagini() As String = sPathsImm.Split("§")
			Dim Conta As Long
			Dim Db As New clsGestioneDB(TipoDB)
			Dim Sql As String
			Dim Barra As String = "\"
			Dim Rec As Object

			If TipoDB = "SQLSERVER" Then
				Barra = "\"
			Else
				Barra = "/"
			End If

			Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
			If ConnessioneSQL <> "" Then
				Dim idCategoria As Integer = -1

				Try
					gf.CreaDirectoryDaPercorso(Server.MapPath(".") & "\Logs")
				Catch ex As Exception

				End Try

				Sql = "Select * From Categorie Where idTipologia=" & idTipologia & " And Categoria='" & Categoria & "'"
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				If Rec.Eof = False Then
					idCategoria = Rec("idCategoria").Value
				Else
					Return "ERROR: Categoria non trovata"
				End If
				Rec.Close

				'Db.EsegueSql(Server.MapPath("."), "Delete From Dati Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria, ConnessioneSQL)

				If idTipologia = "2" Then
					Conta = 0
					gf.ScriveTestoSuFileAperto(dataAttuale() & " - VIDEO")
					gf.ScriveTestoSuFileAperto(dataAttuale() & " -----------------------------------------------------------")
					Try
						For Each p As String In PathVideo
							Dim pp() As String = p.Split(";")
							Dim Nome As String = pp(0)
							'Dim idCategoria As Integer = -1
							'Dim idVecchioCategoria As Integer = 0

							If Nome.ToUpper.Trim = Categoria.ToUpper.Trim Then
								'	If idCategoria = -1 Or idCategoria <> idVecchioCategoria Then
								'		idVecchioCategoria = idCategoria

								'		gf.ScriveTestoSuFileAperto(dataAttuale() & " - Lettura categoria: " & Categoria)
								'		Sql = "Select * From Categorie Where idTipologia=" & idTipologia & " And Categoria='" & Categoria & "'"
								'		Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
								'		If Rec.Eof = False Then
								'			idCategoria = Rec("idCategoria").Value
								'			gf.ScriveTestoSuFileAperto(dataAttuale() & " - Letto id categoria: " & idCategoria)
								'		Else
								'			gf.ScriveTestoSuFileAperto(dataAttuale() & " - Lettura categoria: " & "ERROR: Categoria non trovata")
								'			Return "ERROR: Categoria non trovata"
								'		End If
								'		Rec.Close
								'	End If

								gf.ScriveTestoSuFileAperto(dataAttuale() & " - Elaborazione video: " & p)

								'idCategoria += 1
								'Sql = "Insert Into categorie Values (" & idCategoria & ", 2, '" & pp(0).Replace("'", "''") & "', '" & pp(1).Replace("'", "''") & "', '" & pp(2) & "', '" & pp(3) & "')"
								'Dim sRitorno As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

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

									Dim SoloNome As String = Files(i).Replace(pp(1), "").Replace("'", "''")
									If TipoDB = "SQLSERVER" Then
										SoloNome = SoloNome.Replace("\", "/")
									End If

									Dim SoloNomeConv As String = SoloNome
									Dim este As String = gf.TornaEstensioneFileDaPath(SoloNomeConv)
									SoloNomeConv = SoloNomeConv.Replace(este, "") & "_CONV.mp4"

									Sql = "Select * from dati where idTipologia=2 And idCategoria=" & idCategoria & " And (nomefile='" & SoloNome & "' Or nomefile = '" & SoloNomeConv & "')"
									Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
									If Rec.Eof = True Then
										gf.ScriveTestoSuFileAperto(dataAttuale() & " - Aggiungo file " & Files(i))

										Dim Dime As String = gf.TornaDimensioneFile(Files(i))
										Dim Datella As String = gf.TornaDataDiCreazione(Files(i))

										Conta += 1

										Sql = "Insert Into dati Values (" &
											" " & Conta & ", " &
											"2, " &
											" " & idCategoria & ", " &
											"'" & SoloNome & "', " &
											" " & Dime & ", " &
											"'" & Datella & "', " &
											"'N' " &
											")"
										Dim sRitorno As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
										If sRitorno <> "OK" Then
											gf.ScriveTestoSuFileAperto(dataAttuale() & " - Errore: " & sRitorno)
											gf.ScriveTestoSuFileAperto(dataAttuale() & " - SQL: " & Sql)

											StaLeggendoImmagini = False

											gf.ChiudeFileDiTestoDopoScrittura()

											Return "ERROR: " & sRitorno & " -> " & Sql
										End If
									End If
								Next

								Dim Eliminati As Integer = 0

								Sql = "Select * from dati where idTipologia=2 And idCategoria=" & idCategoria & " and nomefile not like '%_CONV%'"
								Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
								Do Until Rec.Eof
									Dim Nometto As String = Rec("NomeFile").Value.ToString
									Dim Ok As Boolean = False

									For i As Integer = 1 To qFiles
										If Files(i).Contains(Nometto) Then
											Ok = True
											Exit For
										End If
									Next

									If Ok = False Then
										gf.ScriveTestoSuFileAperto(dataAttuale() & " - Rimuovo file " & Nometto & " Progressivo " & Rec("Progressivo").Value)

										Sql = "Update Dati Set Eliminata = 'S' Where idTipologia=2 And idCategoria=" & idCategoria & " And Progressivo=" & Rec("Progressivo").Value
										Dim sRitorno As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
										If sRitorno <> "OK" Then
											gf.ScriveTestoSuFileAperto(dataAttuale() & " - Errore su eliminazione: " & sRitorno)
											gf.ScriveTestoSuFileAperto(dataAttuale() & " - SQL: " & Sql)

											StaLeggendoImmagini = False

											gf.ChiudeFileDiTestoDopoScrittura()

											Return "ERROR: " & sRitorno & " -> " & Sql
										End If

										Eliminati += 1
									End If

									Rec.MoveNext
								Loop
								Rec.Close

								gf.ScriveTestoSuFileAperto(dataAttuale() & " - Fine Scrittura: " & qFiles & "/" & qFiles)
								gf.ScriveTestoSuFileAperto(dataAttuale() & " - Aggiunti: " & Conta)
								gf.ScriveTestoSuFileAperto(dataAttuale() & " - Eliminati: " & Eliminati)
							End If
						Next
					Catch ex As Exception
						gf.ScriveTestoSuFileAperto(dataAttuale() & "ERRORE su elaborazione video: Tipologia: " & idTipologia & " Categoria:" & Categoria & " -> " & ex.Message)
					End Try

					gf.ScriveTestoSuFileAperto(dataAttuale() & " -----------------------------------------------------------")

					gf.ScriveTestoSuFileAperto("")
					gf.ScriveTestoSuFileAperto("")
				End If

				If idTipologia = "1" Then
					gf.ScriveTestoSuFileAperto(dataAttuale() & " - IMMAGINI")
					gf.ScriveTestoSuFileAperto(dataAttuale() & " -----------------------------------------------------------")
					Conta = 0
					'Dim idCategoria As Integer = -1
					'Dim idVecchioCategoria As Integer = 0

					Try
						For Each p As String In PathImmagini
							Dim pp() As String = p.Split(";")
							Dim Nome As String = pp(0)
							Dim Estensione As String = gf.TornaEstensioneFileDaPath(Nome).Replace(".", "").ToUpper

							If Nome.ToUpper.Trim = Categoria.ToUpper.Trim And Not Nome.ToUpper.Contains(".NOMEDIA") And Not Nome.ToUpper.Contains("THUMBS") Then
								'If idCategoria = -1 Or idCategoria <> idVecchioCategoria Then
								'	idVecchioCategoria = idCategoria

								'	gf.ScriveTestoSuFileAperto(dataAttuale() & " - Lettura categoria: " & Categoria)
								'	Sql = "Select * From Categorie Where idTipologia=" & idTipologia & " And Categoria='" & Categoria & "'"
								'	Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
								'	If Rec.Eof = False Then
								'		idCategoria = Rec("idCategoria").Value
								'		gf.ScriveTestoSuFileAperto(dataAttuale() & " - Letto id categoria: " & idCategoria)
								'	Else
								'		gf.ScriveTestoSuFileAperto(dataAttuale() & " - Lettura categoria: " & "ERROR: Categoria non trovata")
								'		Return "ERROR: Categoria non trovata"
								'	End If
								'	Rec.Close
								'End If

								gf.ScriveTestoSuFileAperto(dataAttuale() & " - Elaborazione immagini: " & p)

								'idCategoria += 1
								'Sql = "Insert Into categorie Values (" & idCategoria & ", 1, '" & pp(0).Replace("'", "''") & "', '" & pp(1).Replace("'", "''") & "', '" & pp(2) & "', '" & pp(3) & "')"
								'Dim sRitorno As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

								If Strings.Right(pp(1), 1) <> Barra Then
									pp(1) &= Barra
								End If
								gf.ScriveTestoSuFileAperto(dataAttuale() & " - Scansione cartella: " & pp(1))
								gf.ScansionaDirectorySingola(pp(1))
								Dim qFiles As Integer = gf.RitornaQuantiFilesRilevati
								Dim Files() As String = gf.RitornaFilesRilevati
								gf.ScriveTestoSuFileAperto(dataAttuale() & " - Numero files: " & qFiles)
								Conta = 0
								For i As Integer = 1 To qFiles
									If (i / 1000 = Int(i / 1000)) Then
										gf.ScriveTestoSuFileAperto(dataAttuale() & " - Scrittura: " & i & "/" & qFiles)
									End If

									Dim SoloNome As String = Files(i).Replace(pp(1), "").Replace("'", "''")
									If TipoDB = "SQLSERVER" Then
										SoloNome = SoloNome.Replace("\", "/")
									End If

									Sql = "Select * from dati where idTipologia=1 And idCategoria=" & idCategoria & " And nomefile='" & SoloNome & "'"
									Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
									If Rec.Eof = True Then
										gf.ScriveTestoSuFileAperto(dataAttuale() & " - Aggiungo file " & Files(i))

										Dim Dime As String = gf.TornaDimensioneFile(Files(i))
										Dim Datella As String = gf.TornaDataDiCreazione(Files(i))

										Conta += 1

										Sql = "Insert Into dati Values (" &
											" " & Conta & ", " &
											"1, " &
											" " & idCategoria & ", " &
											"'" & SoloNome & "', " &
											" " & Dime & ", " &
											"'" & Datella & "', " &
											"'N' " &
											")"
										Dim sRitorno As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
										If sRitorno <> "OK" Then
											gf.ScriveTestoSuFileAperto(dataAttuale() & " - Errore: " & sRitorno)
											gf.ScriveTestoSuFileAperto(dataAttuale() & " - SQL: " & Sql)
											StaLeggendoImmagini = False

											gf.ChiudeFileDiTestoDopoScrittura()

											Return "ERROR: " & sRitorno
										End If
									End If
								Next

								Dim Eliminati As Integer = 0

								Sql = "Select * from dati where idTipologia=1 And idCategoria=" & idCategoria
								Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
								Do Until Rec.Eof
									Dim Nometto As String = Rec("NomeFile").Value.ToString
									Dim Ok As Boolean = False

									For i As Integer = 1 To qFiles
										If Files(i).Contains(Nometto) Then
											Ok = True
											Exit For
										End If
									Next

									If Ok = False Then
										gf.ScriveTestoSuFileAperto(dataAttuale() & " - Elimino file " & Nometto & " Progressivo " & Rec("Progressivo").Value)

										Sql = "Update Dati Set Eliminata = 'N' Where idTipologia=1 And idCategoria=" & idCategoria & " And Progressivo=" & Rec("Progressivo").Value
										Dim sRitorno As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
										If sRitorno <> "OK" Then
											gf.ScriveTestoSuFileAperto(dataAttuale() & " - Errore su eliminazione: " & sRitorno)
											gf.ScriveTestoSuFileAperto(dataAttuale() & " - SQL: " & Sql)

											StaLeggendoImmagini = False

											gf.ChiudeFileDiTestoDopoScrittura()

											Return "ERROR: " & sRitorno & " -> " & Sql
										End If
										Eliminati += 1
									End If

									Rec.MoveNext
								Loop
								Rec.Close

								gf.ScriveTestoSuFileAperto(dataAttuale() & " - Fine Scrittura: " & qFiles & "/" & qFiles)
								gf.ScriveTestoSuFileAperto(dataAttuale() & " - Aggiunti: " & Conta)
								gf.ScriveTestoSuFileAperto(dataAttuale() & " - Eliminati: " & Eliminati)
							End If
						Next
					Catch ex As Exception
						gf.ScriveTestoSuFileAperto(dataAttuale() & "ERRORE su elaborazione immagini: Tipologia: " & idTipologia & " Categoria:" & Categoria & " -> " & ex.Message)
					End Try
					gf.ScriveTestoSuFileAperto(dataAttuale() & " -----------------------------------------------------------")
				End If

				gf.ScriveTestoSuFileAperto("")
				gf.ScriveTestoSuFileAperto(dataAttuale() & " - RIEPILOGO")
				Try
					If TipoDB = "SQLSERVER" Then
						Sql = "Select Tipologia, Categoria, Isnull(Count(*),0) As Quanti From Dati A " &
							"Left Join Categorie B On A.idCategoria = B.idCategoria " &
							"Left Join Tipologie C On A.idTipologia = C.idTipologia " &
							"Where A.Eliminata = 'N' Or A.Eliminata = 'n' " &
							"Group By Tipologia, Categoria " &
							"Order By 1,2"
					Else
						Sql = "Select Tipologia, Categoria, COALESCE(Count(*),0) As Quanti From Dati A " &
							"Left Join Categorie B On A.idCategoria = B.idCategoria " &
							"Left Join Tipologie C On A.idTipologia = C.idTipologia " &
							"Where A.Eliminata = 'N' Or A.Eliminata = 'n' " &
							"Group By Tipologia, Categoria " &
							"Order By 1,2"
					End If

					Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
					Do Until Rec.Eof
						Dim Tipologia As String = Rec("Tipologia").Value
						Dim Categoria2 As String = Rec("Categoria").Value
						Dim Quanti As String = Rec("Quanti").Value

						gf.ScriveTestoSuFileAperto(dataAttuale() & " - " & Tipologia & ": " & Categoria2 & " -> Files " & Quanti)

						Rec.MoveNext
					Loop
					Rec.Close
				Catch ex As Exception
					gf.ScriveTestoSuFileAperto(dataAttuale() & " ERRORE Nel riepilogo: " & ex.Message)
				End Try
				gf.ScriveTestoSuFileAperto(dataAttuale() & " -----------------------------------------------------------")
			End If
		Catch ex As Exception
			gf.ScriveTestoSuFileAperto(dataAttuale() & ex.Message)
		End Try

		VecchiaRicerca = ""
		StaLeggendoImmagini = False

		gf.ChiudeFileDiTestoDopoScrittura()

		gf = Nothing

		Return "*"
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
		Dim P As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PercorsiSfondi.txt")
		Dim pp() As String = P.Split(";")
		Dim PathSfondi As String = pp(0).Replace(vbCrLf, "")
		Dim PathSfondiDir As String = pp(1).Replace(vbCrLf, "")

		Dim Immagine As String = ""

		If TipoDB <> "SQLSERVER" Then
			PathSfondi = PathSfondi.Replace("\", "/")
			PathSfondi = PathSfondi.Replace("//", "/")
			PathSfondi = PathSfondi.Replace("/\", "/")

			PathSfondiDir = PathSfondiDir.Replace("\", "/")
			PathSfondiDir = PathSfondiDir.Replace("//", "/")
			PathSfondiDir = PathSfondiDir.Replace("/\", "/")
		End If

		Try
			MkDir(PathSfondi)
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
			'If Strings.Right(PathSfondiDir, 1) <> Barra Then
			'	PathSfondiDir &= Barra
			'End If

			gf.ScansionaDirectorySingola(PathSfondiDir)
			Dim filetti() As String = gf.RitornaFilesRilevati
			QuanteImmaginiSfondi = gf.RitornaQuantiFilesRilevati

			For i As Integer = 1 To QuanteImmaginiSfondi
				ListaImmagini.Add(filetti(i))
			Next
		Else
			ContatoreRiletturaImmagini += 1
		End If

		'Return PathSfondi & " - " & PathSfondiDir & " - " & QuanteImmaginiSfondi

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

		Dim NomeFile As String = PathSfondi & "/Sfondo_" & Minuti & ".txt" ' Server.MapPath(".") & "\Sfondi\Sfondo_" & Minuti & ".txt"

		If TipoDB <> "SQLSERVER" Then
			NomeFile = NomeFile.Replace("\", "/")
			NomeFile = NomeFile.Replace("//", "/")
			NomeFile = NomeFile.Replace("/\", "/")
		End If

		If gf.EsisteFile(NomeFile) Then
			Immagine = gf.LeggeFileIntero(NomeFile)
		Else
			gf.ApreFileDiTestoPerScrittura(NomeFile)

			Dim Perc As String = PathSfondi & "\"

			If TipoDB <> "SQLSERVER" Then
				Perc = Perc.Replace("\", "/")
				Perc = Perc.Replace("//", "/")
				Perc = Perc.Replace("/\", "/")
			End If

			For Each foundFile As String In My.Computer.FileSystem.GetFiles(Perc)
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

			Immagine = y.ToString & ";" & ListaImmagini.Item(y).Replace(PathSfondiDir, "") & ";"

			gf.ScriveTestoSuFileAperto(Immagine)
			gf.ChiudeFileDiTestoDopoScrittura()
		End If

		Return Immagine
	End Function

	<WebMethod()>
	Public Function EliminaMultimediaDaId(idTipologia As String, idCategoria As String, idMultimedia As String) As String
		Dim Ritorno As String = ""
		Dim Sql As String
		Dim Barra As String = "\"
		Dim Db As New clsGestioneDB(TipoDB)
		Dim NomeFileLog As String = Server.MapPath(".") & "/Logs/EliminazioneImmagine.txt"

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
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			If Rec.Eof = False Then
				Dim Conversione As String = "" & Rec("LetteraDisco").Value.ToString
				Dim PathOriginale As String = Rec("Percorso").Value.ToString
				If Conversione <> "" And Conversione <> "--" Then
					Dim cc() As String = Conversione.Split("*")
					PathOriginale = PathOriginale.Replace(cc(0), cc(1))
				End If
				If Right(PathOriginale, 1) <> Barra Then
					PathOriginale &= Barra
				End If

				Dim gf As New GestioneFilesDirectory

				Dim filetto As String = PathOriginale & Rec("NomeFile").value
				Dim CartellaIntermedia As String = PathOriginale
				If CartellaIntermedia.Contains("CartelleCondivise") Then
					CartellaIntermedia = Mid(CartellaIntermedia, CartellaIntermedia.IndexOf("CartelleCondivise") + 18, CartellaIntermedia.Length)
				End If
				Dim daCopiare As String = Server.MapPath(".") & "BackupCancellazioni" & CartellaIntermedia & Rec("NomeFile").value
				gf.CreaDirectoryDaPercorso(daCopiare)
				gf.ImpostaAttributiFile(Server.MapPath(".") & "BackupCancellazioni" & CartellaIntermedia, FileAttribute.Normal)
				ScriveLogGlobale(NomeFileLog, "Impostati attributi cartella di backup: " & Server.MapPath(".") & "BackupCancellazioni" & CartellaIntermedia)

				ScriveLogGlobale(NomeFileLog, "-----------------------------------------")
				ScriveLogGlobale(NomeFileLog, "Eliminazione idTipologia " & idTipologia & " idCategoria " & idCategoria & " idMultimedia " & idMultimedia)
				ScriveLogGlobale(NomeFileLog, "Nome File origine: " & filetto)
				ScriveLogGlobale(NomeFileLog, "Nome File destinazione: " & daCopiare)

				Dim R As String = ""
				Try
					R = gf.CopiaFileFisico(filetto, daCopiare, True)
				Catch ex As Exception
					If giaProvatoACancellare = False Then
						Threading.Thread.Sleep(1000)
						ScriveLogGlobale(NomeFileLog, "Problema con i permessi... Riprovo a eseguire la funzione")
						giaProvatoACancellare = True
						Dim Ritorno2 As String = EliminaMultimediaDaId(idTipologia, idCategoria, idMultimedia)
						Return Ritorno2
					Else
						giaProvatoACancellare = False
						Return "ERROR: " & ex.Message
					End If
				End Try

				If r.Contains("ERROR:") Then
					ScriveLogGlobale(NomeFileLog, r)
				Else
					ScriveLogGlobale(NomeFileLog, "File backuppato")

					gf.ImpostaAttributiFile(filetto, FileAttribute.Normal)
					ScriveLogGlobale(NomeFileLog, "Impostati attributi file")

					Ritorno = gf.EliminaFileFisico(filetto)
					If Ritorno = "" Then
						ScriveLogGlobale(NomeFileLog, "File eliminato")

						Sql = "Update dati Set Eliminata='S' Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria & " And Progressivo=" & idMultimedia
						Dim Rit As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

						If Rit = "OK" Then
							ScriveLogGlobale(NomeFileLog, "Eliminazione multimedia su tabella OK")

							If idTipologia = "1" Then
								Sql = "Update informazioniimmagini Set Eliminata='S' Where idCategoria=" & idCategoria & " And idMultimedia=" & idMultimedia
								Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
								If Rit = "OK" Then
									ScriveLogGlobale(NomeFileLog, "Eliminazione informazioni immagine su tabella OK")

									Sql = "Delete From preferiti Where idTipologia=1 And idCategoria=" & idCategoria & " And Progressivo=" & idMultimedia
									Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
									If Rit = "OK" Then
										ScriveLogGlobale(NomeFileLog, "Eliminazione preferito su tabella OK")

										Sql = "Delete From preferitiprot Where idTipologia=1 And idCategoria=" & idCategoria & " And Progressivo=" & idMultimedia
										Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
										If Rit = "OK" Then
											ScriveLogGlobale(NomeFileLog, "Eliminazione preferito protetto su tabella OK")

											Ritorno = "*"
										Else
											ScriveLogGlobale(NomeFileLog, Rit)
										End If
									Else
										ScriveLogGlobale(NomeFileLog, Rit)
									End If
								Else
									ScriveLogGlobale(NomeFileLog, Rit)
								End If
							End If

							If idTipologia = "2" Then
								Sql = "Update informazionivideo Set Eliminata='S' Where idTipologia=2 And idCategoria=" & idCategoria & " And idMultimedia=" & idMultimedia
								Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
								If Rit = "OK" Then
									ScriveLogGlobale(NomeFileLog, "Eliminazione informazioni video su tabella OK")

									Sql = "Delete From preferiti Where idTipologia=2 And idCategoria=" & idCategoria & " And Progressivo=" & idMultimedia
									Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
									If Rit = "OK" Then
										ScriveLogGlobale(NomeFileLog, "Eliminazione preferito su tabella OK")

										Sql = "Delete From preferitiprot Where idTipologia=2 And idCategoria=" & idCategoria & " And Progressivo=" & idMultimedia
										Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
										If Rit = "OK" Then
											ScriveLogGlobale(NomeFileLog, "Eliminazione preferito protetto su tabella OK")

											Ritorno = "*"
										Else
											ScriveLogGlobale(NomeFileLog, Rit)
										End If
									Else
										ScriveLogGlobale(NomeFileLog, Rit)
									End If
								Else
									ScriveLogGlobale(NomeFileLog, Rit)
								End If
							End If
						Else
							ScriveLogGlobale(NomeFileLog, Rit)
						End If
					Else
						ScriveLogGlobale(NomeFileLog, Ritorno)
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

	<WebMethod()>
	Public Function CalcolaHashTutteImmagini(idCategoria As String) As String
		idCategoriaGlobalePerConversione = idCategoria

		timerConvI = New Timers.Timer(100)
		AddHandler timerConvI.Elapsed, New ElapsedEventHandler(AddressOf ConverteTutteLeImmaginiThread)
		timerConvI.Start()

		Return "*"
	End Function

	Private Sub ConverteTutteLeImmaginiThread()
		timerConvI.Stop()

		Dim NomeFileLog As String = Server.MapPath(".") & "/Logs/RitornaInformazioniImmagine.txt"

		If StaEffettuandoConversioneAutomaticaI = True Then
			Exit Sub
		End If

		Dim Db As New clsGestioneDB(TipoDB)

		Try
			Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
			If ConnessioneSQL <> "" Then
				Dim Rec As Object
				Dim gf As New GestioneFilesDirectory
				gf.EliminaFileFisico(NomeFileLog)

				Dim Sql As String = "Select A.idCategoria, A.Progressivo From Dati A " &
					"Left Join informazioniimmagini B On A.idcategoria = B.idCategoria And A.progressivo = B.idMultimedia " &
					"Where A.idTipologia=1 And A.idCategoria = " & idCategoriaGlobalePerConversione & " And (A.Eliminata='N' Or A.Eliminata='n') And B.Hash is Null" '  And (B.Eliminata='N' Or B.Eliminata='n') "
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				If TypeOf (Rec) Is String Then
					ScriveLogGlobale(NomeFileLog, Rec)
					Exit Sub
				End If

				StaEffettuandoConversioneAutomaticaI = True
				StaEffettuandoConversioneAutomaticaFinaleI = True

				Dim Quale As Long = 1

				Do Until Rec.Eof
					Dim idCategoria As String = Rec("idCategoria").Value
					Dim idMultimedia As String = Rec("Progressivo").Value

					'ScriveLogGlobale(NomeFileLog, "-----------------------------------------")
					'ScriveLogGlobale(NomeFileLog, "Acquisizione informazioni immagine idCategoria " & idCategoria & " idMultimedia " & idMultimedia)

					Dim Rit As String = CalcolaHashImmagine(idCategoria, idMultimedia, "N", Quale)
					If Not Rit.Contains("ERROR:") Then
						Quale += 1
						'If Quale = 5 Then
						'	StaEffettuandoConversioneAutomaticaI = False
						'	StaEffettuandoConversioneAutomaticaFinaleI = False

						'	Exit Do
						'End If
					End If

					'ScriveLogGlobale(NomeFileLog, "Acquisizione informazioni immagine " & Rit)

					If StaEffettuandoConversioneAutomaticaI = False Then
						StaEffettuandoConversioneAutomaticaFinaleI = False
						ScriveLogGlobale(NomeFileLog, "Blocco elaborazione da web")

						Exit Do
					End If

					Rec.MoveNext
				Loop
				Rec.Close

				StaEffettuandoConversioneAutomaticaI = False
				StaEffettuandoConversioneAutomaticaFinaleI = False
			End If

		Catch ex As Exception

		End Try

		idCategoriaGlobalePerConversione = ""
	End Sub

	' http://looigi.ddns.net:1050/looVF.asmx/CalcolaPuntiniImmagine?idCategoria=4&Refresh=S
	<WebMethod()>
	Public Function CalcolaPuntiniImmagine(idCategoria As String, Refresh As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		Dim gi As New GestioneImmagini
		Dim gf As New GestioneFilesDirectory
		Dim NomeFileLog As String = Server.MapPath(".") & "/Logs/RitornaPuntiniImmagine.txt"
		Dim Ritorno As String = "*"

		gf.EliminaFileFisico(NomeFileLog)

		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim Sql As String = ""

			Dim NumeroImmagineConvertita As Long = 0

			If Refresh = "S" Then
				Sql = "Update informazioniimmagini Set PuntiDiagonale=null, PuntiCornice=null"
				Dim Rit As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

				If Rit = "OK" Then
					ScriveLogGlobale(NomeFileLog, "Eliminazione tabella effettuata")
				Else
					ScriveLogGlobale(NomeFileLog, Rit)
					Return Rit
				End If
			End If

			Dim q As Integer = 0
			' Acquisizione puntini per ricerca più precisa
			Sql = "Select * From informazioniimmagini A " &
				"Left Join categorie B On B.idTipologia = 1 And A.idCategoria = B.idCategoria " &
				"Left Join dati C On C.idtipologia = 1 And C.idcategoria = A.idCategoria And C.progressivo = A.idMultimedia " &
				"Where (PuntiDiagonale is Null Or PuntiDiagonale = '') And (A.Eliminata='N' Or A.Eliminata='n') And A.idCategoria=" & idCategoria
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL, False)
			Do Until Rec.Eof
				Dim Percorso As String = Rec("Percorso").Value
				Dim NomeFile As String = Rec("NomeFile").Value
				Dim idMultimedia As String = Rec("idMultimedia").Value

				ScriveLogGlobale(NomeFileLog, "-----------------------------------------")
				ScriveLogGlobale(NomeFileLog, "Acquisizione puntini " & Percorso & "/" & NomeFile)

				NumeroImmagineConvertita += 1
				Dim Puntini As String = gi.CalcolaPuntini(Server.MapPath("."), Percorso & "/" & NomeFile, NomeFileLog, NumeroImmagineConvertita)
				If Puntini.Contains("ERROR") Then
					ScriveLogGlobale(NomeFileLog, Puntini)
				Else
					Dim Pv() As String = Puntini.Split(";")

					Dim PuntiDiagonale As String = Pv(0)
					Dim PuntiCornice As String = Pv(1)

					ScriveLogGlobale(NomeFileLog, "Acquisizione puntini. Punti diagonale " & PuntiDiagonale)
					ScriveLogGlobale(NomeFileLog, "Acquisizione puntini. Punti cornice " & PuntiCornice)

					Sql = "Update informazioniimmagini " &
						"Set PuntiDiagonale=" & PuntiDiagonale & ", PuntiCornice=" & PuntiCornice & " " &
						"Where idCategoria=" & idCategoria & " And idMultimedia=" & idMultimedia
					Dim Rit As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

					If Rit = "OK" Then
						ScriveLogGlobale(NomeFileLog, "Scrittura in tabella effettuata")
					Else
						ScriveLogGlobale(NomeFileLog, Rit)
					End If
				End If

				'q += 1
				'If q > 0 Then
				'	Exit Do
				'End If

				Rec.MoveNext
			Loop
			Rec.Close

			Ritorno = "*"
			' Acquisizione puntini per ricerca più precisa
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function CalcolaHashImmagine(idCategoria As String, idMultimedia As String, Refresh As String, NumeroImmagineConvertita As String) As String
		Dim NomeFileLog As String = Server.MapPath(".") & "/Logs/RitornaInformazioniImmagine.txt"
		Dim Ritorno As String = ""
		Dim gi As New GestioneImmagini
		Dim gf As New GestioneFilesDirectory
		Dim Db As New clsGestioneDB(TipoDB)

		Try
			ScriveLogGlobale(NomeFileLog, "-----------------------------------------")
			ScriveLogGlobale(NomeFileLog, "Acquisizione informazioni immagine idCategoria " & idCategoria & " idMultimedia " & idMultimedia)

			Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
			If ConnessioneSQL <> "" Then
				Dim Rec As Object
				Dim Sql As String = ""
				Dim EsegueRicerca As Boolean

				If Refresh = "N" Then
					Sql = "Select * From InformazioniImmagini Where idCategoria=" & idCategoria & " And idMultimedia=" & idMultimedia & " And (Eliminata='N' Or Eliminata='n')"
					Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
					If TypeOf (Rec) Is String Then
						Return Rec
					End If

					If Rec.Eof = False Then
						ScriveLogGlobale(NomeFileLog, "Dati rilevati in tabella")

						EsegueRicerca = False
						Ritorno = Rec("Hash").Value & ";" & Rec("Punti").Value & ";" & Rec("Width").Value & ";" & Rec("Height").Value & ";" & Rec("DataOra").Value
					Else
						EsegueRicerca = True
					End If
				Else
					EsegueRicerca = True
				End If

				If EsegueRicerca Or Refresh = "S" Then
					ScriveLogGlobale(NomeFileLog, "Dati NON rilevati in tabella oppure Refresh = 'S'")

					Sql = "Select * From Dati A " &
						"Left Join Categorie B On A.idTipologia = B.idTipologia And A.idCategoria = B.idCategoria " &
						"Where A.idTipologia=1 And A.idCategoria=" & idCategoria & " And Progressivo=" & idMultimedia
					Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
					If TypeOf (Rec) Is String Then
						Return Rec
					End If

					If Rec.Eof = True Then
						Ritorno = "ERROR: File non trovato sulla tabella"
						ScriveLogGlobale(NomeFileLog, Ritorno)
					Else
						Dim Percorso As String = Rec("Percorso").Value
						Dim NomeFile As String = Rec("NomeFile").Value
						Rec.Close

						Dim Estensione As String = gf.TornaEstensioneFileDaPath(NomeFile).Replace(".", "").ToUpper

						If NomeFile.ToUpper.Contains(".NOMEDIA") Or NomeFile.ToUpper.Contains("THUMBS") Or (Estensione <> "JPG" And Estensione <> "JPEG" And Estensione <> "PNG" And Estensione <> "BMP") Then
							' If NomeFile.ToUpper.Contains(".NOMEDIA") Then
							ScriveLogGlobale(NomeFileLog, "Acquisizione informazioni immagine. Skippo per nome non valido: " & Percorso & "/" & NomeFile)
						Else
							ScriveLogGlobale(NomeFileLog, "Acquisizione informazioni immagine. Nome immagine: " & Percorso & "/" & NomeFile)

							Dim RitornoHash As StrutturaJPG = gi.CalcolaMD5(Server.MapPath("."), Percorso & "/" & NomeFile, NomeFileLog, NumeroImmagineConvertita)

							If RitornoHash.Hash.Contains("ERROR:") Then
								Ritorno = RitornoHash.Hash

								''ScriveLogGlobale(NomeFileLog, Ritorno)
							Else
								ScriveLogGlobale(NomeFileLog, "Acquisizione informazioni immagine effettuata")

								Sql = "Delete From informazioniimmagini Where idCategoria=" & idCategoria & " And idMultimedia=" & idMultimedia
								Dim Rit As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
								If Rit = "OK" Then
									Sql = "Insert Into informazioniimmagini Values (" & idCategoria & ", " & idMultimedia & ", '" & RitornoHash.Hash & "', " & RitornoHash.Punti & ", " & RitornoHash.Width & ", " & RitornoHash.Height & ", '" & RitornoHash.DataOra & "', " & RitornoHash.PuntiDiagonale & ", " & RitornoHash.PuntiCornice & ", 'N')"
									Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

									If Rit = "OK" Then
										ScriveLogGlobale(NomeFileLog, "Scrittura effettuata in tabella")

										Ritorno = RitornoHash.Hash & ";" & RitornoHash.Punti & ";" & RitornoHash.Width & ";" & RitornoHash.Height & ";" & RitornoHash.DataOra
									Else
										Ritorno = "ERROR: Errore sull'inserimento dei dati in tabella. " & Sql

										ScriveLogGlobale(NomeFileLog, Ritorno)
									End If
								Else
									Ritorno = "ERROR: Errore sull'eliminazione dei dati dalla tabella. " & Sql

									ScriveLogGlobale(NomeFileLog, Ritorno)
								End If
							End If
						End If
					End If
				End If
			Else
				Ritorno = "ERROR: Errore sulla connessione al db. " & ConnessioneSQL

				ScriveLogGlobale(NomeFileLog, Ritorno)
			End If

			If Ritorno <> "" Then
				ScriveLogGlobale(NomeFileLog, "Acquisizione informazioni immagine " & Ritorno)
			End If
		Catch ex As Exception
			ScriveLogGlobale(NomeFileLog, "ERROR: " & ex.Message)
			Ritorno = ex.Message
		End Try

		Return Ritorno
	End Function

	' Select 'Punti pieni', count(*) From informazioniimmagini A Where (PuntiDiagonale is Not Null Or PuntiDiagonale <> '') And A.idCategoria=4 union ALL Select 'Punti nulli', count(*) From informazioniimmagini A where (PuntiDiagonale is Null Or PuntiDiagonale = '') And A.idCategoria = 4 Union ALL Select 'Tutto', Count(*) From informazioniimmagini A

	<WebMethod()>
	Public Function TrovaImmaginiUguali(idCategoria As String, ricercaPerHash As String, ricercaPerData As String, ricercaPerDimensioni As String, ricercaPerPunti As String, QuanteImmagini As String, Inizio As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		Dim Ritorno As String = "*"

		If ConnessioneSQL <> "" Then
			Dim RitornoHash As String = ""
			Dim RitornoPunti As String = ""
			Dim RitornoDataOra As String = ""
			Dim RitornoDimensioni As String = ""

			If ricercaPerHash = "S" Then
				RitornoHash = RitornaUguaglianze(Server.MapPath("."), Db, ConnessioneSQL, "Hash", "HASH", idCategoria, QuanteImmagini, Inizio)
			End If
			RitornoHash &= "|"

			If ricercaPerData = "S" Then
				RitornoDataOra = RitornaUguaglianze(Server.MapPath("."), Db, ConnessioneSQL, "DataOra", "DATA ORA", idCategoria, QuanteImmagini, Inizio)
			End If
			RitornoDataOra &= "|"

			If ricercaPerDimensioni = "S" Then
				RitornoDimensioni = RitornaUguaglianze(Server.MapPath("."), Db, ConnessioneSQL, "Concat(Width, 'x', height)", "DIMENSIONI", idCategoria, QuanteImmagini, Inizio)
			End If
			RitornoDimensioni &= "|"

			If ricercaPerPunti = "S" Then
				RitornoPunti = RitornaUguaglianze(Server.MapPath("."), Db, ConnessioneSQL, "Concat(PuntiDiagonale, '-', PuntiCornice)", "PUNTI", idCategoria, QuanteImmagini, Inizio)
			End If
			RitornoPunti &= "|"

			Ritorno = RitornoHash & RitornoDataOra & RitornoDimensioni & RitornoPunti
		End If

		Return Ritorno
	End Function

End Class