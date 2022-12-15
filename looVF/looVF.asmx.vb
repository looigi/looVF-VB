Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.IO
Imports System.Threading
Imports System.Timers
Imports System.Web.Script.Serialization
Imports looVF.GestioneImmagini
Imports System.Drawing

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

	Private Function RitornaSuccessivoMultimediaPerPiccole(db As clsGestioneDB, ConnessioneSQL As String, idTipologia As String, Categoria As String, Filtro As String,
															 Random As String, NomeFileLog As String, idCategoria As String) As String
		Dim Ritorno As String = ""
		Dim QuanteRighePiccole As Long = 0
		Dim Sql As String = ""
		Dim NomeTabella As String = ""
		Dim Altro As String = ""
		Dim Rec As Object

		If Filtro <> "" Then
			Altro = " And Upper(B.NomeFile) Like '%" & Filtro.ToUpper & "%'"
		End If

		ScriveLogGlobale(NomeFileLog, "Ricerca Successivo per piccole: idTipologia " & idTipologia & " idCategoria " & idCategoria & " Categoria " & Categoria & " Filtro " & Filtro & " Random " & Random)

		Sql = "Select Coalesce(Count(*), 0) As Quante From dati A " &
			"Left Join Dati B On A.idTipologia=B.idTipologia And A.idCategoria=B.idCategoria And A.Progressivo=B.Progressivo " &
			"Where (B.Eliminata = 'N' Or B.Eliminata = 'n') And A.idTipologia=" & idTipologia & " " & Altro
		If idCategoria <> "" Then
			Sql &= " And A.idCategoria=" & idCategoria
		End If
		ScriveLogGlobale(NomeFileLog, Sql)
		Rec = db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
		QuanteRighePiccole = Rec("Quante").Value
		Rec.Close
		ScriveLogGlobale(NomeFileLog, "Quante righe piccole: " & QuanteRighePiccole)

		Dim Ultimo As Integer = -1

		Ultimo = UltimoMultimediaImm
		ScriveLogGlobale(NomeFileLog, "Ultimo MM impostato: " & Ultimo)

		Static x As Random = New Random()

		Dim y As Long = -1

		If Random = "S" Or Random = "" Then
			y = x.Next(QuanteRighePiccole)
			ScriveLogGlobale(NomeFileLog, "Valore Random per Random 'S': " & y & "/" & QuanteRighePiccole)
		Else
			If Ultimo <> -1 Then
				y = Ultimo + 1
				If y > QuanteRighePiccole Then
					y = 0
				End If
				ScriveLogGlobale(NomeFileLog, "Valore Sequenziale per Random 'N': " & y & "/" & QuanteRighePiccole)
			Else
				y = x.Next(QuanteRighePiccole)
				ScriveLogGlobale(NomeFileLog, "Valore Random per Random 'N': " & y & "/" & QuanteRighePiccole & " Ultimo = -1")
			End If
		End If

		If y = -1 Then
			Return "ERROR: Non riesco a impostare il valore. Quante righe: " & QuanteRighePiccole
		End If

		UltimoMultimediaImm = y

		Dim NumeroImmagine As Long
		Dim idCategoriaRilevato As Integer
		Dim CategoriaRilevata As String
		Dim Altro2 As String = ""

		If Filtro <> "" Then
			Altro2 = " And Upper(C.NomeFile) Like '%" & Filtro.ToUpper & "%'"
		End If

		Sql = "Select idTipologia, idCategoria, Progressivo, Categoria, NomeFile From ( " &
			"SELECT ROW_NUMBER() OVER(Order BY idTipologia, idCategoria, progressivo) As Numero, A.idTipologia, A.idCategoria, A.Progressivo, B.Categoria, A.NomeFile " &
			"FROM dati A " &
			"Left Join categorie B On A.idtipologia = B.idtipologia And A.idCategoria = B.idcategoria " &
			"Where A.idCategoria=" & idCategoria & " " &
			" " & Altro2 & " " &
			") As a Where Numero=" & y
		'"where idtipologia=" & idTipologia & " " & ' " and idcategoria=" & idCategoria & " " &
		ScriveLogGlobale(NomeFileLog, Sql)
		Rec = db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
		If Rec.Eof = False Then
			NumeroImmagine = Rec("Progressivo").Value
			idCategoriaRilevato = Rec("idCategoria").Value
			CategoriaRilevata = Rec("Categoria").Value

			Ritorno = NumeroImmagine.ToString & ";" & idCategoriaRilevato & ";" & QuanteRighePiccole & ";" & CategoriaRilevata & ";" & Ultimo & ";" & y
		Else
			Ritorno = "ERROR: Nessun multimedia rilevato"
		End If
		Rec.Close

		ScriveLogGlobale(NomeFileLog, "Ritorno: " & Ritorno)

		Return Ritorno
	End Function

	Private Function RitornaSuccessivoMultimediaPerTutto(db As clsGestioneDB, ConnessioneSQL As String, idTipologia As String, Filtro As String,
															 Random As String, NomeFileLog As String) As String
		Dim Ritorno As String = ""
		Dim QuanteRigheTutto As Long = 0
		Dim Sql As String = ""
		Dim NomeTabella As String = ""
		Dim Altro As String = ""
		Dim Rec As Object

		If Filtro <> "" Then
			Altro = " And Upper(B.NomeFile) Like '%" & Filtro.ToUpper & "%'"
		End If

		ScriveLogGlobale(NomeFileLog, "Ricerca Successivo per tutto: idTipologia " & idTipologia & " Filtro " & Filtro & " Random " & Random)

		Sql = "Select Coalesce(Count(*), 0) As Quante From dati A " &
			"Left Join Dati B On A.idTipologia=B.idTipologia And A.idCategoria=B.idCategoria And A.Progressivo=B.Progressivo " &
			"Where (B.Eliminata = 'N' Or B.Eliminata = 'n') And A.idTipologia=" & idTipologia & " " & Altro
		ScriveLogGlobale(NomeFileLog, Sql)
		Rec = db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
		QuanteRigheTutto = Rec("Quante").Value
		Rec.Close
		ScriveLogGlobale(NomeFileLog, "Quante righe tutto: " & QuanteRigheTutto)

		Dim Ultimo As Integer = -1

		Ultimo = UltimoMultimediaImm
		ScriveLogGlobale(NomeFileLog, "Ultimo MM impostato: " & Ultimo)

		Static x As Random = New Random()

		Dim y As Long = -1

		If Random = "S" Or Random = "" Then
			y = x.Next(QuanteRigheTutto)
			ScriveLogGlobale(NomeFileLog, "Valore Random per Random 'S': " & y & "/" & QuanteRigheTutto)
		Else
			If Ultimo <> -1 Then
				y = Ultimo + 1
				If y > QuanteRigheTutto Then
					y = 0
				End If
				ScriveLogGlobale(NomeFileLog, "Valore Sequenziale per Random 'N': " & y & "/" & QuanteRigheTutto)
			Else
				y = x.Next(QuanteRigheTutto)
				ScriveLogGlobale(NomeFileLog, "Valore Random per Random 'N': " & y & "/" & QuanteRigheTutto & " Ultimo = -1")
			End If
		End If

		If y = -1 Then
			Return "ERROR: Non riesco a impostare il valore. Quante righe: " & QuanteRigheTutto
		End If

		UltimoMultimediaImm = y

		Dim NumeroImmagine As Long
		Dim idCategoriaRilevato As Integer
		Dim CategoriaRilevata As String
		Dim Altro2 As String = ""

		If Filtro <> "" Then
			Altro2 = " Where Upper(C.NomeFile) Like '%" & Filtro.ToUpper & "%'"
		End If

		Sql = "Select idTipologia, idCategoria, Progressivo, Categoria, NomeFile From ( " &
			"SELECT ROW_NUMBER() OVER(Order BY idTipologia, idCategoria, progressivo) As Numero, A.idTipologia, A.idCategoria, A.Progressivo, B.Categoria, A.NomeFile " &
			"FROM dati A " &
			"Left Join categorie B On A.idtipologia = B.idtipologia And A.idCategoria = B.idcategoria " &
			" " & Altro2 & " " &
			") As a Where Numero=" & y
		'"where idtipologia=" & idTipologia & " " & ' " and idcategoria=" & idCategoria & " " &
		ScriveLogGlobale(NomeFileLog, Sql)
		Rec = db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
		If Rec.Eof = False Then
			NumeroImmagine = Rec("Progressivo").Value
			idCategoriaRilevato = Rec("idCategoria").Value
			CategoriaRilevata = Rec("Categoria").Value

			Ritorno = NumeroImmagine.ToString & ";" & idCategoriaRilevato & ";" & QuanteRigheTutto & ";" & CategoriaRilevata & ";" & Ultimo & ";" & y
		Else
			Ritorno = "ERROR: Nessun multimedia rilevato"
		End If
		Rec.Close

		ScriveLogGlobale(NomeFileLog, "Ritorno: " & Ritorno)

		Return Ritorno
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
		If idCategoria <> "" Then
			Altro2 = " And C.idCategoria=" & idCategoria
		End If

		'Sql = "Select idTipologia, idCategoria, Progressivo, Categoria, NomeFile From ( " &
		'	"SELECT ROW_NUMBER() OVER(Order BY idTipologia, idCategoria, progressivo) As Numero, A.idTipologia, A.idCategoria, A.Progressivo, B.Categoria, C.NomeFile " &
		'	"FROM " & NomeTabella & " A " &
		'	"Left Join categorie B On A.idtipologia = B.idtipologia And A.idCategoria = B.idcategoria " &
		'	"Left Join Dati C On A.idtipologia = C.idtipologia And A.idCategoria = C.idcategoria And A.Progressivo = C.Progressivo " &
		'	" " & Altro2 & " " &
		'	") As a Where Numero=" & y

		Sql = "Select idTipologia, idCategoria, Progressivo, Categoria, NomeFile From ( " &
				"SELECT ROW_NUMBER() OVER(Order BY idTipologia, idCategoria, progressivo) As Numero, A.idTipologia, A.idCategoria, A.Progressivo, B.Categoria, C.NomeFile " &
				"FROM " & NomeTabella & " A " &
				"Left Join categorie B On A.idtipologia = B.idtipologia And A.idCategoria = B.idcategoria " &
				"Left Join Dati C On A.idtipologia = C.idtipologia And A.idCategoria = C.idcategoria And A.Progressivo = C.Progressivo " &
				" " & Altro & " " &
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

			If Categoria = "Piccole" Then
				ScriveLogGlobale(NomeFileLog, "Categoria uguale a Piccole: " & Categoria & ". Vado alla funzione adatta")

				Ritorno = RitornaSuccessivoMultimediaPerPiccole(Db, ConnessioneSQL, idTipologia, Categoria, Filtro, Random, NomeFileLog, idCategoria)

				Return Ritorno
			End If

			If Categoria = "Tutto" Then
				ScriveLogGlobale(NomeFileLog, "Categoria uguale a Tutto: " & Categoria & ". Vado alla funzione adatta")

				Ritorno = RitornaSuccessivoMultimediaPerTutto(Db, ConnessioneSQL, idTipologia, Filtro, Random, NomeFileLog)

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
				If Filtro <> "" Then
					ScriveLogGlobale(NomeFileLog, "Valore Random per Random 'S': " & y & "/" & Quante & " -> Indice: " & Indici.Item(y))
				Else
					ScriveLogGlobale(NomeFileLog, "Valore Random per Random 'S': " & y & "/" & Quante)
				End If
			Else
				If Ultimo <> -1 Then
					If Filtro <> "" Then
						Dim c As Integer = 0

						ScriveLogGlobale(NomeFileLog, "Scanno indice per rilevare ultimo MM: " & Ultimo)
						For Each i As Integer In Indici
							If i = Ultimo Then
								Ultimo = c
							End If
							c += 1
						Next

						ScriveLogGlobale(NomeFileLog, "Imposto ultimo MM da indice: " & c)
					End If

					y = Ultimo + 1
					If y > Quante Then
						y = 0
					End If

					If Filtro <> "" Then
						ScriveLogGlobale(NomeFileLog, "Valore Sequenziale: " & y & "/" & Quante & " -> Indice: " & Indici.Item(y))
					Else
						ScriveLogGlobale(NomeFileLog, "Valore Sequenziale: " & y & "/" & Quante)
					End If
				Else
					y = x.Next(Quante)
					If Filtro <> "" Then
						ScriveLogGlobale(NomeFileLog, "Valore Random per modalità sequenziale " & y & "/" & Quante & " Ultimo = -1 -> Indice: " & Indici.Item(y))
					Else
						ScriveLogGlobale(NomeFileLog, "Valore Random per modalità sequenziale " & y & "/" & Quante & " Ultimo = -1")
					End If
				End If
			End If

			If Filtro <> "" Then
				y = Indici.Item(y)
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
			'Dim Inizio As Long = 0

			'If idCategoria <> "" Then
			If Filtro = "" Then '  And Categoria <> "Preferiti" And Categoria <> "Preferiti Prot"
				ScriveLogGlobale(NomeFileLog, "Filtro nullo")

				'Sql = "Select Coalesce(Min(Progressivo),0) From Dati Where (Eliminata = 'N' Or Eliminata = 'n') And idTipologia=" & idTipologia & " " & Altro & " "
				'If idCategoria <> "" Then
				'	Sql &= "And idCategoria=" & idCategoria
				'End If
				''gf.ScriveTestoSuFileAperto(NomeFile, Sql)
				'Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				'Inizio = (Rec(0).Value - 1)
				'Rec.Close

				'ScriveLogGlobale(NomeFileLog, "Inizio: " & Inizio)
			Else
				ScriveLogGlobale(NomeFileLog, "Filtro NON Nullo (" & Filtro & ")")

				'If Indici.Count > y Then
				'	Inizio = Indici(y)
				'Else
				'	Inizio = -1
				'End If
				'y = 0

				'If Categoria <> "Preferiti" And Categoria <> "Preferiti Prot" Then
				'End If

				'ScriveLogGlobale(NomeFileLog, "Inizio: " & Inizio & " Indici: " & Indici.Count - 1)
			End If

			'If Random = "N" Then ' And (Categoria = "Preferiti" Or Categoria = "Preferiti Prot") Then
			'	'Inizio = 0

			'	ScriveLogGlobale(NomeFileLog, "Random No e NON Preferiti, azzero l'inizio")
			'End If

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

			'If Filtro <> "" Then
			'	Altro = " And Upper(B.NomeFile) Like '%" & Filtro.ToUpper & "%'"
			'End If

			'Sql = "Select idTipologia, idCategoria, Progressivo, Categoria, NomeFile From ( " &
			'	"SELECT ROW_NUMBER() OVER(Order BY idTipologia, idCategoria, progressivo) As Numero, A.idTipologia, A.idCategoria, A.Progressivo, B.Categoria, A.NomeFile " &
			'	"FROM dati A " &
			'	"Left Join categorie B On A.idtipologia = B.idtipologia And A.idCategoria = B.idcategoria " &
			'	" " & Altro & " " &
			'	") As a Where Numero=" & y
			'ScriveLogGlobale(NomeFileLog, Sql)
			'Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			'If TypeOf (Rec) Is String Then
			'	ScriveLogGlobale(NomeFileLog, Rec)
			'	Return "ERROR: Errore sql -> " & Sql
			'End If
			'If Rec.Eof = False Then
			'	y = Rec("Progressivo").Value
			'	idCategoria = Rec("idCategoria").Value
			'Else
			'	Return "ERROR: Nessun valore ritornato. " & Sql
			'End If
			'Rec.Close

			'If idCategoria = "" Then
			'	' ScriveLogGlobale(NomeFileLog, "Categoria Nulla. La ricerco per idTipologia " & idTipologia & " e idMultimedia " & Inizio + y)
			'	ScriveLogGlobale(NomeFileLog, "Categoria Nulla. La ricerco per idTipologia " & idTipologia & " e idMultimedia " & y)

			'	' Sql = "Select idCategoria From Dati Where (Eliminata = 'N' Or Eliminata = 'n') And idTipologia=" & idTipologia & " And Progressivo=" & Inizio + y
			'	Sql = "Select idCategoria From Dati Where (Eliminata = 'N' Or Eliminata = 'n') And idTipologia=" & idTipologia & " And Progressivo=" & y
			'	Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			'	idCategoria = Rec("idCategoria").Value
			'	Rec.Close

			'	ScriveLogGlobale(NomeFileLog, "Categoria Rilevata: " & idCategoria)
			'End If

			'Dim Quanto As Integer

			'Quanto = y ' (Inizio + y)
			'Return Inizio & ";" & y
			' ScriveLogGlobale(NomeFileLog, "Id Multimedia ritornato: " & Quanto & " - Inizio: " & Inizio & " Y: " & y)
			ScriveLogGlobale(NomeFileLog, "Id Multimedia ritornato: " & y)
			ScriveLogGlobale(NomeFileLog, "Id Categoria Multimedia ritornata: " & idCategoria)

			' Ritorno = Quanto.ToString & ";" & idCategoria & ";" & Quante & ";" & Categoria & ";" & Ultimo & ";-1" ' & Inizio
			Ritorno = y.ToString & ";" & idCategoria & ";" & Quante & ";" & Categoria & ";" & Ultimo & ";-1" ' & Inizio
			ScriveLogGlobale(NomeFileLog, "Ritorno: " & Ritorno)
		End If
		'gf.ScriveTestoSuFileAperto(NomeFile, Ritorno)

		Return Ritorno
	End Function

	'<WebMethod()>
	'Public Function RitornaSuccessivoMultimedia(idTipologia As String, Categoria As String) As String
	'	Dim Db As New clsGestioneDB(TipoDB)
	'	Dim Ritorno As String = ""
	'	Dim Sql As String

	'	'Dim gf As New GestioneFilesDirectory
	'	'Dim NomeFile As String = Server.MapPath(".") & "\Log\LogRitorno.txt"
	'	Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
	'	If ConnessioneSQL <> "" Then
	'		Dim Rec As Object
	'		Dim idCategoria As String = ""

	'		'gf.ScriveTestoSuFileAperto(NomeFile, idTipologia & "-" & Categoria)

	'		If Categoria <> "" And Categoria <> "Tutto" Then
	'			Sql = "Select * From Categorie Where idTipologia=" & idTipologia & " And Categoria='" & Categoria & "'"
	'			'gf.ScriveTestoSuFileAperto(NomeFile, Sql)
	'			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
	'			If Rec.Eof = False Then
	'				idCategoria = Rec("idCategoria").Value
	'			Else
	'				Return "ERROR: Categoria non trovata"
	'			End If
	'			Rec.Close
	'		End If

	'		Sql = "Select Count(*) From Dati Where idTipologia=" & idTipologia
	'		If Categoria <> "" And Categoria <> "Tutto" Then
	'			Sql &= " And idCategoria=" & idCategoria
	'		End If
	'		'gf.ScriveTestoSuFileAperto(NomeFile, Sql)
	'		Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
	'		Dim Quante As Long = Rec(0).Value
	'		Rec.Close
	'		'gf.ScriveTestoSuFileAperto(NomeFile, Quante)

	'		Static x As Random = New Random()
	'		Dim y As Long = x.Next(Quante)
	'		Dim Inizio As Long = 0

	'		If Categoria <> "" And Categoria <> "Tutto" Then
	'			Sql = "Select Min(Progressivo) From Dati Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria
	'			'gf.ScriveTestoSuFileAperto(NomeFile, Sql)
	'			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
	'			If Rec(0).Value Is DBNull.Value Then
	'				Inizio = 0
	'			Else
	'				Inizio = Rec(0).Value - 1
	'			End If
	'			Rec.Close
	'		Else
	'			idCategoria = -1
	'		End If

	'		Ritorno = (Inizio + y).ToString & ";" & idCategoria
	'	End If
	'	'gf.ScriveTestoSuFileAperto(NomeFile, Ritorno)

	'	Return Ritorno
	'End Function

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
				Ritorno &= Rec("idCategoria").Value & ";" & Rec("Categoria").Value & ";" & Rec("QuanteAggiunte").Value & ";"
			End If
			Rec.Close()
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaCategorie(idTipologia As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Sql As String

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object

			Sql = "Select * From Categorie Where idTipologia=" & idTipologia & " Order By AggiuntaInutile, QuanteAggiunte Desc, Categoria"
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			Do Until Rec.Eof
				Dim Nome As String = ""

				If Rec("SottoCategoria").Value <> -1 Then
					Sql = "Select * From Categorie Where idTipologia=" & idTipologia & " And idCategoria=" & Rec("SottoCategoria").Value & " And idTipologia=" & Rec("idTipologia").Value
					Dim Rec2 As Object
					Rec2 = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
					Dim CategoriaMadre As String = Rec2("Categoria").Value.ToString.Replace(";", "***PV***")
					Rec2.Close

					Nome = CategoriaMadre & "/" & Rec("Categoria").Value.ToString.Replace(";", "***PV***")
				Else
					Nome = Rec("Categoria").Value.ToString.Replace(";", "***PV***")
				End If

				Ritorno &= Rec("idCategoria").Value & ";" & Rec("idTipologia").Value & ";" & Nome & ";" & Rec("Protetta").Value & ";" & Rec("QuanteAggiunte").Value & ";" & Rec("Percorso").Value & ";" & Rec("AggiuntaInutile").Value & ";" & Rec("SottoCategoria").Value & "§"

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
	Public Function RitornaQuantiPreferiti(idTipologia As String, idCategoria As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Sql As String

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object

			Sql = "Select Coalesce(Count(*), 0) From Preferiti Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			Ritorno = Rec(0).Value & ";"
			Rec.Close()

			Sql = "Select Coalesce(Count(*), 0) From PreferitiProt Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			Ritorno &= Rec(0).Value
			Rec.Close()
		End If

		Return Ritorno
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

			If Not Ritorno.Contains("ERROR:") Then
				Dim Rec As Object

				Sql = "Select Coalesce(Count(*), 0) From Preferiti Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				Ritorno = Rec(0).Value & ";"
				Rec.Close()

				Sql = "Select Coalesce(Count(*), 0) From PreferitiProt Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				Ritorno &= Rec(0).Value
				Rec.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaMultimediaDaId(idTipologia As String, idCategoria As String, idMultimedia As String, rilevaUguali As String, UgualiSoloCategoria As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Ritorno As String = ""
		Dim Sql As String
		Dim NomeFileLog As String = Server.MapPath(".") & "/Logs/RitornaMultimediaDaID.txt"

		ScriveLogGlobale(NomeFileLog, "-----------------------------------------")
		ScriveLogGlobale(NomeFileLog, "Inizio")

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim Rec2 As Object
			Dim gf As New GestioneFilesDirectory

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
				Dim PathImmagineG As String = Rec("Percorso").Value & "/" & Rec("NomeFile").Value
				Dim FilePresente As String = "N"
				ScriveLogGlobale(NomeFileLog, "Immagine da ricercare: " & PathImmagineG)

				If gf.EsisteFile(PathImmagineG) Then
					FilePresente = "S"
					ScriveLogGlobale(NomeFileLog, "File esistente")
				Else
					ScriveLogGlobale(NomeFileLog, "File NON esistente")
				End If

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
						DatiHash = Rec2("Sezione1").Value & ";" & Rec2("Punti").Value & ";" & Rec2("Width").Value & ";" & Rec2("Height").Value & ";" & Rec2("DataOra").Value
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

				Dim Uguali As String = ""

				If rilevaUguali = "S" Then
					Dim gi As New GestioneImmagini

						Sql = "SELECT * FROM informazioniimmagini A " &
						"Left Join dati B On A.idMultimedia = B.progressivo And A.idCategoria = B.idcategoria And B.idtipologia = 1 " &
						"Left Join categorie C On A.idCategoria = C.idcategoria And B.idtipologia = B.idTipologia " &
						"Where A.idcategoria = " & idCategoria & " And A.idMultimedia = " & idMultimedia & " And B.idTipologia = 1"
					ScriveLogGlobale(NomeFileLog, Sql)
					Rec2 = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL, False)
					If Rec2.Eof = False Then
						Dim Sezione1 As String = Rec2("Sezione1").Value
						ScriveLogGlobale(NomeFileLog, "Sezione1: " & Sezione1)
						Dim Sezione2 As String = Rec2("Sezione2").Value
						ScriveLogGlobale(NomeFileLog, "Sezione2: " & Sezione2)
						Dim Punti As String = Rec2("Punti").Value
						ScriveLogGlobale(NomeFileLog, "Punti: " & Punti)
						Dim PuntiDiagonale As String = Rec2("PuntiDiagonale").Value
						ScriveLogGlobale(NomeFileLog, "PuntiDiagonale: " & PuntiDiagonale)
						Dim PuntiCornice As String = Rec2("PuntiCornice").Value
						ScriveLogGlobale(NomeFileLog, "PuntiCornice: " & PuntiCornice)
						Dim Hash As String = Rec2("Hash").Value
						ScriveLogGlobale(NomeFileLog, "Hash: " & Hash)
						Dim Width As String = Rec2("Width").Value
						ScriveLogGlobale(NomeFileLog, "Width: " & Width)
						Dim Height As String = Rec2("Height").Value
						ScriveLogGlobale(NomeFileLog, "Height: " & Height)
						Dim Dimensioni As String = Rec2("Dimensioni").Value.ToString
						ScriveLogGlobale(NomeFileLog, "Dimensioni: " & Dimensioni)
						Dim DataOra As String = Rec2("DataOra").Value
						ScriveLogGlobale(NomeFileLog, "DataOra: " & DataOra)
						Dim SoloNome As String = Rec2("SoloNome").Value
						ScriveLogGlobale(NomeFileLog, "SoloNome: " & SoloNome)

						Dim NomeEsteso As String = Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00")
						Dim PathImmagine As String = Rec2("Percorso").Value & "/" & Rec2("NomeFile").Value
						Dim Imm1 As Bitmap
						Dim Trovata As Boolean = False

						If gf.EsisteFile(PathImmagine) Then
							gi.RidimensionaMantenendoProporzioni(PathImmagine, Server.MapPath(".") & "/Appoggio/1Resize_" & NomeEsteso & ".jpg", 50, False)
							gi.ConverteImmaginInBN(Server.MapPath(".") & "/Appoggio/1Resize_" & NomeEsteso & ".jpg", Server.MapPath(".") & "/Appoggio/1BN_" & NomeEsteso & ".jpg", False)
							Imm1 = gi.LoadBitmapSenzaLock(Server.MapPath(".") & "/Appoggio/1BN_" & NomeEsteso & ".jpg")
							gf.EliminaFileFisico(Server.MapPath(".") & "/Appoggio/1Resize_" & NomeEsteso & ".jpg")
							gf.EliminaFileFisico(Server.MapPath(".") & "/Appoggio/1BN_" & NomeEsteso & ".jpg")
							Trovata = True
						End If

						Rec2.Close

						If Trovata Then
							Dim Altro As String = ""

							If UgualiSoloCategoria = "S" Then
								Altro = "A.idcategoria = " & idCategoria & " And "
							End If

							Sql = "SELECT Percorso, NomeFile, A.idMultimedia, A.idCategoria FROM informazioniimmagini A " &
								"Left Join dati B On A.idMultimedia = B.progressivo And A.idCategoria = B.idcategoria And B.idtipologia = 1 " &
								"Left Join categorie C On A.idCategoria = C.idcategoria And B.idtipologia = 1 " &
								"Where " & Altro & " A.idMultimedia <> " & idMultimedia & " And " &
								"(" &
								"(Cast(Sezione1 As Integer) > " & Val(Sezione1) - 10 & " And Cast(Sezione1 As Integer) < " & Val(Sezione1) + 10 & ") " &
								"Or (Cast(Sezione2 As Integer) > " & Val(Sezione2) - 10 & " And Cast(Sezione2 As Integer) < " & Val(Sezione2) + 10 & ") " &
								"Or Punti = '" & Punti & "' Or PuntiDiagonale = '" & PuntiDiagonale & "' " &
								"Or PuntiCornice = '" & PuntiCornice & "' Or Hash = '" & Hash & "' Or DataOra = '" & DataOra & "' " &
								"Or ((Cast(Width As Integer) > " & Val(Width) - 10 & " And Cast(Width As Integer) < " & Val(Width) + 10 & ") And (Cast(Height As Integer) > " & Val(Height) - 10 & " And Cast(Height As Integer) < " & Val(Height) + 10 & ")) " &
								"Or (Cast(Dimensioni As Integer) > " & Val(Dimensioni) - 10 & " And Cast(Dimensioni As Integer) < " & Val(Dimensioni) + 10 & ") " &
								"Or Upper(Trim(SoloNome)) = '" & SoloNome.ToUpper.Trim & "'" &
								") " &
								"And Hash <> '' And Hash Is Not Null"
							ScriveLogGlobale(NomeFileLog, Sql)
							Rec2 = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL, False)
							Dim Conta As Integer = 0
							Do Until Rec2.Eof
								Conta += 1
								NomeEsteso = Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & "_" & Conta
								Dim PathImmagine2 As String = Rec2("Percorso").Value & "/" & Rec2("NomeFile").Value
								gi.RidimensionaMantenendoProporzioni(PathImmagine2, Server.MapPath(".") & "/Appoggio/2Resize_" & NomeEsteso & ".jpg", 50, False)
								gi.ConverteImmaginInBN(Server.MapPath(".") & "/Appoggio/2Resize_" & NomeEsteso & ".jpg", Server.MapPath(".") & "/Appoggio/2BN_" & NomeEsteso & ".jpg", False)
								Dim Imm2 As Bitmap
								If gf.EsisteFile(Server.MapPath(".") & "/Appoggio/2BN_" & NomeEsteso & ".jpg") Then
									Imm2 = gi.LoadBitmapSenzaLock(Server.MapPath(".") & "/Appoggio/2BN_" & NomeEsteso & ".jpg")
									gf.EliminaFileFisico(Server.MapPath(".") & "/Appoggio/2Resize_" & NomeEsteso & ".jpg")
									gf.EliminaFileFisico(Server.MapPath(".") & "/Appoggio/2BN_" & NomeEsteso & ".jpg")
									Dim Uguale As Boolean = False

									ScriveLogGlobale(NomeFileLog, "Controllo immagine: " & Rec2("NomeFile").Value)
									If Imm1 Is Imm2 Then
										ScriveLogGlobale(NomeFileLog, "Immagine uguale in tutto")
										Uguale = True
									Else
										Dim QuantiUguali As Integer = 0
										Dim QuantiDiversi As Integer = 0
										Dim PuntiTotali = (Imm1.Width) * (Imm1.Height)

										For X = 0 To Imm1.Width - 1
											For y = 0 To Imm1.Height - 1
												'			End If
												'If Imm1.GetPixel(X, y) = Imm2.GetPixel(X, y) Then
												Dim Primo As Color = Imm1.GetPixel(X, y)
												Dim secondo As Color = Imm2.GetPixel(X, y)

												If Primo.R > 128 And secondo.R > 128 Then
													QuantiUguali += 1
												Else
													QuantiDiversi += 1
												End If
											Next
										Next
										Dim Perc As Single = CInt((QuantiDiversi / PuntiTotali) * 100)

										'ScriveLogGlobale(NomeFileLog, "Punti rilevati uguali: " & QuantiUguali)
										'ScriveLogGlobale(NomeFileLog, "Punti rilevati diversi: " & QuantiDiversi)
										'ScriveLogGlobale(NomeFileLog, "Punti totali: " & PuntiTotali)
										ScriveLogGlobale(NomeFileLog, "Percentuale Diversi: " & Perc)

										If Perc < 20 Then
											Uguale = True
										End If
									End If

									If Uguale Then
										Uguali &= Rec2("Percorso").Value.ToString.Replace(";", "***PV***") & ";" & Rec2("NomeFile").Value.ToString.Replace(";", "***PV***") & ";" & Rec2("idMultimedia").Value & ";" & Rec2("idCategoria").Value & "§"
									End If
								End If

								Rec2.MoveNext
							Loop
						End If
					End If
					Rec2.Close
					'ScriveLogGlobale(NomeFileLog, "Sezione 1: " & Sezione1)
					'ScriveLogGlobale(NomeFileLog, "Sezione 2: " & Sezione2)
					'ScriveLogGlobale(NomeFileLog, "Punti: " & Punti)
					'ScriveLogGlobale(NomeFileLog, "PuntiDiagonale: " & PuntiDiagonale)
					'ScriveLogGlobale(NomeFileLog, "PuntiCornice: " & PuntiCornice)
					'ScriveLogGlobale(NomeFileLog, "Hash: " & Hash)
					'ScriveLogGlobale(NomeFileLog, "Percorso: " & PathImmagine)

				End If

				Ritorno = Thumb & "§" & Rec("NomeFile").Value.ToString.Replace(";", "***PV***") & ";" & Rec("Dimensioni").Value & ";" & Rec("Data").Value & ";" & Rec("idCategoria").Value & ";" & idMultimedia.ToString & ";" & Preferito & ";" & DatiHash & ";" & FilePresente & "|" & Uguali
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

			Dim Inizio As Long = 1
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

			'If Categoria <> "" And Categoria <> "Tutto" And Filtro = "" Then
			'	Sql = "Select Min(Progressivo) From Dati Where Dati.idTipologia=" & idTipologia & " And Dati.idCategoria=" & idCategoria ' & Altro
			'	Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			'	Inizio = Rec(0).Value - 1
			'	Rec.Close
			'End If

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
	Public Function RitornaFilesNuovo(idTipologia As String, Categoria As String, sEseguiSoloTest As String) As String
		If StaLeggendoImmagini Then
			Return "ERROR: Sto già caricando multimedia"
		End If

		If Categoria = "" Or Categoria.ToUpper = "TUTTE" Then
			RitornaFiles()
		End If

		Dim NomeFileLog As String = Server.MapPath(".") & "\Logs\RefreshFiles_" & dataAttuale() & "_Categoria_" & Categoria & ".txt"
		'gf.ApreFileDiTestoPerScrittura(NomeFileLog)

		Dim EseguiSoloTest As Boolean = IIf(sEseguiSoloTest = "S", True, False)

		StaLeggendoImmagini = True

		Dim gf As New GestioneFilesDirectory

		Try
			gf.CreaDirectoryDaPercorso(Server.MapPath(".") & "\Logs")
		Catch ex As Exception

		End Try

		ScriveLogGlobale(NomeFileLog, "-----------------------------------------------------------")
		Try
			Dim sPathsVideo As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PercorsiVideo.txt")
			Dim sPathsImm As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PercorsiImmagini.txt")
			ScriveLogGlobale(NomeFileLog, "Stringa Path Video: " & sPathsVideo)
			ScriveLogGlobale(NomeFileLog, "Stringa Path Immagini: " & sPathsImm)

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

				ScriveLogGlobale(NomeFileLog, "Lettura categoria: " & Categoria & " - idTipologia: " & idTipologia)

				Sql = "Select * From Categorie Where idTipologia=" & idTipologia & " And Categoria='" & Categoria & "'"
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				If Rec.Eof = False Then
					idCategoria = Rec("idCategoria").Value
				Else
					Return "ERROR: Categoria non trovata"
				End If
				Rec.Close

				ScriveLogGlobale(NomeFileLog, "Categoria rilevata: " & idCategoria)
				ScriveLogGlobale(NomeFileLog, " -----------------------------------------------------------")

				'Db.EsegueSql(Server.MapPath("."), "Delete From Dati Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria, ConnessioneSQL)

				If idTipologia = "2" Then
					Conta = 0
					ScriveLogGlobale(NomeFileLog, "Acquisizione files VIDEO")
					ScriveLogGlobale(NomeFileLog, " -----------------------------------------------------------")
					Try
						For Each p As String In PathVideo
							Dim pp() As String = p.Split(";")
							Dim Nome As String = pp(0)
							'Dim idCategoria As Integer = -1
							'Dim idVecchioCategoria As Integer = 0

							If Nome.ToUpper.Trim = Categoria.ToUpper.Trim Then
								'	If idCategoria = -1 Or idCategoria <> idVecchioCategoria Then
								'		idVecchioCategoria = idCategoria

								'		ScriveLogGlobale(NomeFileLog,"Lettura categoria: " & Categoria)
								'		Sql = "Select * From Categorie Where idTipologia=" & idTipologia & " And Categoria='" & Categoria & "'"
								'		Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
								'		If Rec.Eof = False Then
								'			idCategoria = Rec("idCategoria").Value
								'			ScriveLogGlobale(NomeFileLog,"Letto id categoria: " & idCategoria)
								'		Else
								'			ScriveLogGlobale(NomeFileLog,"Lettura categoria: " & "ERROR: Categoria non trovata")
								'			Return "ERROR: Categoria non trovata"
								'		End If
								'		Rec.Close
								'	End If

								ScriveLogGlobale(NomeFileLog, "Elaborazione video: Percorso " & p)

								'idCategoria += 1
								'Sql = "Insert Into categorie Values (" & idCategoria & ", 2, '" & pp(0).Replace("'", "''") & "', '" & pp(1).Replace("'", "''") & "', '" & pp(2) & "', '" & pp(3) & "')"
								'Dim sRitorno As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

								If Strings.Right(pp(1), 1) <> Barra Then
									pp(1) &= Barra
								End If
								ScriveLogGlobale(NomeFileLog, "Scansione cartella: " & pp(1))
								gf.ScansionaDirectorySingola(pp(1))
								Dim qFiles As Integer = gf.RitornaQuantiFilesRilevati
								Dim Files() As String = gf.RitornaFilesRilevati
								ScriveLogGlobale(NomeFileLog, "Numero files rilevati: " & qFiles)

								Sql = "Select Coalesce(Max(progressivo),0)+1 As Quante from dati where idTipologia=2 And idCategoria=" & idCategoria
								Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
								Conta = Rec("Quante").Value
								Rec.Close
								ScriveLogGlobale(NomeFileLog, "Max Progressivo per categoria: " & Conta)

								For i As Integer = 1 To qFiles
									If (i / 1000) = Int(i / 1000) Then
										ScriveLogGlobale(NomeFileLog, "Scrittura: " & i & "/" & qFiles)
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
										ScriveLogGlobale(NomeFileLog, "Aggiungo file " & Files(i))

										Dim Dime As String = gf.TornaDimensioneFile(Files(i))
										Dim Datella As String = gf.TornaDataDiCreazione(Files(i))

										'Conta += 1

										Dim NomeSingolo As String = gf.TornaNomeFileDaPath(SoloNome)

										Sql = "Insert Into dati Values (" &
											" " & Conta & ", " &
											"2, " &
											" " & idCategoria & ", " &
											"'" & SoloNome & "', " &
											" " & Dime & ", " &
											"'" & Datella & "', " &
											"'N', " &
											"'" & NomeSingolo.Replace("'", "''") & "' " &
											")"
										Dim sRitorno As String = ""
										If EseguiSoloTest = True Then
											sRitorno = "OK"
										Else
											sRitorno = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
										End If
										If sRitorno <> "OK" Then
											ScriveLogGlobale(NomeFileLog, "Errore: " & sRitorno)
											ScriveLogGlobale(NomeFileLog, "SQL: " & Sql)

											StaLeggendoImmagini = False

											gf.ChiudeFileDiTestoDopoScrittura()

											Return "ERROR: " & sRitorno & " -> " & Sql
										Else
											Conta += 1
											If EseguiSoloTest Then
												ScriveLogGlobale(NomeFileLog, "Scrittura FAKE in tabella eseguita correttamente")
											Else
												ScriveLogGlobale(NomeFileLog, "Scrittura in tabella eseguita correttamente")
											End If
										End If
									End If
								Next

								Dim Eliminati As Integer = 0

								ScriveLogGlobale(NomeFileLog, "Rimozione files non trovati")
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
										If EseguiSoloTest Then
											ScriveLogGlobale(NomeFileLog, "Elimino FAKE file " & Nometto & " Progressivo " & Rec("Progressivo").Value)
										Else
											ScriveLogGlobale(NomeFileLog, "Elimino file " & Nometto & " Progressivo " & Rec("Progressivo").Value)
										End If

										Sql = "Update dati Set Eliminata = 'S' Where idTipologia=2 And idCategoria=" & idCategoria & " And Progressivo=" & Rec("Progressivo").Value
										Dim sRitorno As String = ""
										If EseguiSoloTest = True Then
											sRitorno = "OK"
										Else
											sRitorno = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
										End If
										If sRitorno <> "OK" Then
											ScriveLogGlobale(NomeFileLog, "Errore su eliminazione: " & sRitorno)
											ScriveLogGlobale(NomeFileLog, "SQL: " & Sql)

											StaLeggendoImmagini = False

											gf.ChiudeFileDiTestoDopoScrittura()

											Return "ERROR: " & sRitorno & " -> " & Sql
										End If

										Eliminati += 1
									End If

									Rec.MoveNext
								Loop
								Rec.Close

								ScriveLogGlobale(NomeFileLog, "Fine Scrittura: " & qFiles & "/" & qFiles)
								ScriveLogGlobale(NomeFileLog, "Aggiunti: " & Conta)
								ScriveLogGlobale(NomeFileLog, "Eliminati: " & Eliminati)
							End If
						Next
					Catch ex As Exception
						ScriveLogGlobale(NomeFileLog, "ERRORE su elaborazione video: Tipologia: " & idTipologia & " Categoria:" & Categoria & " -> " & ex.Message)
					End Try

					ScriveLogGlobale(NomeFileLog, " -----------------------------------------------------------")

					gf.ScriveTestoSuFileAperto("")
					gf.ScriveTestoSuFileAperto("")
				End If

				If idTipologia = "1" Then
					ScriveLogGlobale(NomeFileLog, "IMMAGINI")
					ScriveLogGlobale(NomeFileLog, "-----------------------------------------------------------")
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

								'	ScriveLogGlobale(NomeFileLog,"Lettura categoria: " & Categoria)
								'	Sql = "Select * From Categorie Where idTipologia=" & idTipologia & " And Categoria='" & Categoria & "'"
								'	Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
								'	If Rec.Eof = False Then
								'		idCategoria = Rec("idCategoria").Value
								'		ScriveLogGlobale(NomeFileLog,"Letto id categoria: " & idCategoria)
								'	Else
								'		ScriveLogGlobale(NomeFileLog,"Lettura categoria: " & "ERROR: Categoria non trovata")
								'		Return "ERROR: Categoria non trovata"
								'	End If
								'	Rec.Close
								'End If

								ScriveLogGlobale(NomeFileLog, "Elaborazione immagini: Percorso " & p)

								'idCategoria += 1
								'Sql = "Insert Into categorie Values (" & idCategoria & ", 1, '" & pp(0).Replace("'", "''") & "', '" & pp(1).Replace("'", "''") & "', '" & pp(2) & "', '" & pp(3) & "')"
								'Dim sRitorno As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

								If Strings.Right(pp(1), 1) <> Barra Then
									pp(1) &= Barra
								End If
								ScriveLogGlobale(NomeFileLog, "Scansione cartella: " & pp(1))
								gf.ScansionaDirectorySingola(pp(1))
								Dim qFiles As Integer = gf.RitornaQuantiFilesRilevati
								Dim Files() As String = gf.RitornaFilesRilevati
								ScriveLogGlobale(NomeFileLog, "Numero files: " & qFiles)

								Sql = "Select Coalesce(Max(progressivo),0)+1 As Quante from dati where idTipologia=1 And idCategoria=" & idCategoria
								Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
								Conta = Rec("Quante").Value
								Rec.Close
								ScriveLogGlobale(NomeFileLog, "Max Progressivo per categoria: " & Conta)

								'Conta = 0
								For i As Integer = 1 To qFiles
									If (i / 1000 = Int(i / 1000)) Then
										ScriveLogGlobale(NomeFileLog, "Scrittura: " & i & "/" & qFiles)
									End If

									Dim SoloNome As String = Files(i).Replace(pp(1), "").Replace("'", "''")
									If TipoDB = "SQLSERVER" Then
										SoloNome = SoloNome.Replace("\", "/")
									End If

									Sql = "Select * from dati where idTipologia=1 And idCategoria=" & idCategoria & " And nomefile='" & SoloNome & "'"
									Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
									If Rec.Eof = True Then
										ScriveLogGlobale(NomeFileLog, "Aggiungo file " & Files(i))

										Dim Dime As String = gf.TornaDimensioneFile(Files(i))
										Dim Datella As String = gf.TornaDataDiCreazione(Files(i))

										'Conta += 1

										Dim NomeSingolo As String = gf.TornaNomeFileDaPath(SoloNome)

										Sql = "Insert Into dati Values (" &
											" " & Conta & ", " &
											"1, " &
											" " & idCategoria & ", " &
											"'" & SoloNome & "', " &
											" " & Dime & ", " &
											"'" & Datella & "', " &
											"'N', " &
											"'" & NomeSingolo.Replace("'", "''") & "' " &
											")"
										Dim sRitorno As String = ""
										If EseguiSoloTest = True Then
											sRitorno = "OK"
										Else
											sRitorno = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
										End If
										If sRitorno <> "OK" Then
											ScriveLogGlobale(NomeFileLog, "Errore: " & sRitorno)
											ScriveLogGlobale(NomeFileLog, "SQL: " & Sql)
											StaLeggendoImmagini = False

											gf.ChiudeFileDiTestoDopoScrittura()

											Return "ERROR: " & sRitorno
										Else
											Conta += 1
											If EseguiSoloTest Then
												ScriveLogGlobale(NomeFileLog, "Scrittura FAKE in tabella eseguita correttamente")
											Else
												ScriveLogGlobale(NomeFileLog, "Scrittura in tabella eseguita correttamente")
											End If
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
										If EseguiSoloTest Then
											ScriveLogGlobale(NomeFileLog, "Elimino FAKE file " & Nometto & " Progressivo " & Rec("Progressivo").Value)
										Else
											ScriveLogGlobale(NomeFileLog, "Elimino file " & Nometto & " Progressivo " & Rec("Progressivo").Value)
										End If

										Sql = "Update dati Set Eliminata = 'N' Where idTipologia=1 And idCategoria=" & idCategoria & " And Progressivo=" & Rec("Progressivo").Value
										Dim sRitorno As String = ""
										If EseguiSoloTest = True Then
											sRitorno = "OK"
										Else
											sRitorno = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
										End If
										If sRitorno <> "OK" Then
											ScriveLogGlobale(NomeFileLog, "Errore su eliminazione: " & sRitorno)
											ScriveLogGlobale(NomeFileLog, "SQL: " & Sql)

											StaLeggendoImmagini = False

											gf.ChiudeFileDiTestoDopoScrittura()

											Return "ERROR: " & sRitorno & " -> " & Sql
										End If

										Eliminati += 1
									End If

									Rec.MoveNext
								Loop
								Rec.Close

								ScriveLogGlobale(NomeFileLog, "Fine Scrittura: " & qFiles & "/" & qFiles)
								ScriveLogGlobale(NomeFileLog, "Aggiunti: " & Conta)
								ScriveLogGlobale(NomeFileLog, "Eliminati: " & Eliminati)
							End If
						Next
					Catch ex As Exception
						ScriveLogGlobale(NomeFileLog, "ERRORE su elaborazione immagini: Tipologia: " & idTipologia & " Categoria:" & Categoria & " -> " & ex.Message)
					End Try
					ScriveLogGlobale(NomeFileLog, " -----------------------------------------------------------")
				End If

				'gf.ScriveTestoSuFileAperto("")
				ScriveLogGlobale(NomeFileLog, "RIEPILOGO")
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

						ScriveLogGlobale(NomeFileLog, "" & Tipologia & ": " & Categoria2 & " -> Files " & Quanti)

						Rec.MoveNext
					Loop
					Rec.Close
				Catch ex As Exception
					ScriveLogGlobale(NomeFileLog, " ERRORE Nel riepilogo: " & ex.Message)
				End Try
				ScriveLogGlobale(NomeFileLog, " -----------------------------------------------------------")
			End If
		Catch ex As Exception
			ScriveLogGlobale(NomeFileLog, ex.Message)
		End Try

		VecchiaRicerca = ""
		StaLeggendoImmagini = False

		'gf.ChiudeFileDiTestoDopoScrittura()

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

							Dim NomeSingolo As String = gf.TornaNomeFileDaPath(SoloNome)

							Sql = "Insert Into dati Values (" &
								" " & Conta & ", " &
								"2, " &
								" " & idCategoria & ", " &
								"'" & SoloNome & "', " &
								" " & Dime & ", " &
								"'" & Datella & "', " &
								"'N', " &
								"'" & NomeSingolo.Replace("'", "''") & "' " &
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

							Dim NomeSingolo As String = gf.TornaNomeFileDaPath(SoloNome)

							Sql = "Insert Into dati Values (" &
								" " & Conta & ", " &
								"1, " &
								" " & idCategoria & ", " &
								"'" & SoloNome & "', " &
								" " & Dime & ", " &
								"'" & Datella & "', " &
								"'N', " &
								"'" & NomeSingolo.Replace("'", "''") & "' " &
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
	Public Function RinumeraCategoria(idTipologia As String, idCategoria As String) As String
		Dim Ritorno As String = ""
		Dim Sql As String
		Dim Barra As String = "\"
		Dim Db As New clsGestioneDB(TipoDB)

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim Rec2 As Object
			Dim Rit As String = ""

			Dim Progressivo As Integer = 1
			Sql = "Select * From dati Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria ' & " And progressivo >= 50638"
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			Do Until Rec.Eof
				Dim idMultimedia As String = Rec("Progressivo").Value

				Sql = "Select * From dati Where Progressivo=" & Progressivo
				Rec2 = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				If Rec2.Eof = True Then
					Sql = "Update dati Set progressivo=" & Progressivo & " Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria & " And Progressivo=" & idMultimedia
					Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

					If Rit = "OK" Then
						Sql = "Update preferiti Set Progressivo=" & Progressivo & " Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria & " And Progressivo=" & idMultimedia
						Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
						If Rit = "OK" Then
							Sql = "Update preferitiprot Set Progressivo=" & Progressivo & " Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria & " And Progressivo=" & idMultimedia
							Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
							If Rit = "OK" Then
								If idTipologia = "1" Or idTipologia = 1 Then
									Sql = "Update informazioniimmagini Set idMultimedia=" & Progressivo & " Where idCategoria=" & idCategoria & " And idMultimedia=" & idMultimedia
									Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
									If Rit = "OK" Then
										Sql = "Update exifimmagini Set idMultimedia=" & Progressivo & " Where idCategoria=" & idCategoria & " And idMultimedia=" & idMultimedia
										Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
										If Rit = "OK" Then
											Ritorno = "*"
										Else
											Ritorno = Sql & vbCrLf & vbCrLf & Rit
										End If
									Else
										Ritorno = Sql & vbCrLf & vbCrLf & Rit
										'Exit Do
									End If
								Else
									Sql = "Update InformazioniVideo Set idMultimedia=" & Progressivo & " Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria & " And idMultimedia=" & idMultimedia
									Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
									If Rit = "OK" Then
										Ritorno = "*"
									Else
										Ritorno = Sql & vbCrLf & vbCrLf & Rit
										'Exit Do
									End If
								End If
							Else
								Ritorno = Sql & vbCrLf & vbCrLf & Rit
								'Exit Do
							End If
						Else
							Ritorno = Sql & vbCrLf & vbCrLf & Rit
							'Exit Do
						End If
					Else
						Ritorno = Sql & vbCrLf & vbCrLf & Rit
						Exit Do
					End If
				End If

				Progressivo += 1

				Rec.MoveNext
			Loop
			Rec.Close
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaIDCategoria(idTipologia As String, Categoria As String) As String
		Dim Ritorno As String = ""
		Dim Sql As String
		Dim Barra As String = "\"
		Dim Db As New clsGestioneDB(TipoDB)

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object

			Sql = "Select * From categorie Where idTipologia=" & idTipologia & " And categoria='" & Categoria.Replace("'", "''") & "'"
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			If Rec.Eof = True Then
				Ritorno = "ERROR: Nessun id rilevato per la categoria"
			Else
				Ritorno = Rec("idCategoria").Value
			End If
			Rec.Close
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function SpostaTuttiMultimediaACategoria(idTipologia As String, idMultimedia As String, idNuovaCategoria As String) As String
		Dim Ritorno As String = ""

		If idMultimedia = "" Or Not idMultimedia.Contains("§") Then
			Ritorno = "ERROR: Nessun id passato"
		Else
			Dim id() As String = idMultimedia.Split("§")

			For Each i As String In id
				If i <> "" Then
					Dim c() As String = i.Split(";")

					Ritorno = SpostaMultimediaACategoria(idTipologia, c(0), c(1), idNuovaCategoria)
				End If
			Next

			Ritorno = "*"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function SpostaMultimediaACategoria(idTipologia As String, idVecchiaCategoria As String, idMultimedia As String, idNuovaCategoria As String) As String
		Dim Ritorno As String = ""
		Dim Sql As String
		Dim Barra As String = "\"
		Dim Db As New clsGestioneDB(TipoDB)
		Dim NomeFileLog As String = Server.MapPath(".") & "/Logs/SpostamentoFileACategoria.txt"

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object

			ScriveLogGlobale(NomeFileLog, "------------------------------------")
			ScriveLogGlobale(NomeFileLog, "Acquisizione Path Categoria per idTipologia " & idTipologia & " idCategoria " & idVecchiaCategoria)

			Sql = "Select * From Categorie Where idTipologia=" & idTipologia & " And idCategoria=" & idVecchiaCategoria
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			Dim pathVecchiaCategoria As String = Rec("Percorso").Value
			Rec.Close
			ScriveLogGlobale(NomeFileLog, "Path vecchia categoria " & pathVecchiaCategoria)

			ScriveLogGlobale(NomeFileLog, "Acquisizione Path Categoria per idTipologia " & idTipologia & " idCategoria " & idNuovaCategoria)
			Sql = "Select * From Categorie Where idTipologia=" & idTipologia & " And idCategoria=" & idNuovaCategoria
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			Dim pathNuovaCategoria As String = Rec("Percorso").Value
			Rec.Close
			ScriveLogGlobale(NomeFileLog, "Path nuova categoria " & pathNuovaCategoria)

			ScriveLogGlobale(NomeFileLog, "Acquisizione Nome File per idTipologia " & idTipologia & " idCategoria " & idVecchiaCategoria & " idMultimedia " & idMultimedia)
			Sql = "Select * From Dati Where idTipologia=" & idTipologia & " And idCategoria=" & idVecchiaCategoria & " And Progressivo=" & idMultimedia
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			Dim NomeFile As String = Rec("NomeFile").Value
			Rec.Close
			ScriveLogGlobale(NomeFileLog, "Nome File " & NomeFile)

			Dim gf As New GestioneFilesDirectory
			ScriveLogGlobale(NomeFileLog, "Creazione Cartella destinazione: " & pathNuovaCategoria & "/" & NomeFile)
			gf.CreaDirectoryDaPercorso(pathNuovaCategoria & "/" & NomeFile)
			'ScriveLogGlobale(NomeFileLog, "Imposto attributi file origine")
			'gf.ImpostaAttributiFile(pathVecchiaCategoria & "/" & NomeFile, FileAttribute.Normal)
			ScriveLogGlobale(NomeFileLog, "Copia file.")
			ScriveLogGlobale(NomeFileLog, "Origine: " & pathVecchiaCategoria & "/" & NomeFile)
			ScriveLogGlobale(NomeFileLog, "Copia file Destinazione: " & pathNuovaCategoria & "/" & NomeFile)
			Dim Rit As String = "" ' gf.CopiaFileFisico(pathVecchiaCategoria & "/" & NomeFile, pathNuovaCategoria & "/" & NomeFile, True)
			Dim EffettuaCopia As Boolean = True

			If gf.EsisteFile(pathVecchiaCategoria & "/" & NomeFile) Then
				If gf.EsisteFile(pathNuovaCategoria & "/" & NomeFile) Then
					ScriveLogGlobale(NomeFileLog, "Eliminazione destinazione")
					gf.EliminaFileFisico(pathNuovaCategoria & "/" & NomeFile)
					ScriveLogGlobale(NomeFileLog, "File di destinazione eliminato")
				End If

				Try
					File.Copy(pathVecchiaCategoria & "/" & NomeFile, pathNuovaCategoria & "/" & NomeFile, True)
				Catch ex As Exception
					ScriveLogGlobale(NomeFileLog, "Copia file: " & ex.Message)
					EffettuaCopia = False
				End Try
			Else
				ScriveLogGlobale(NomeFileLog, "File non esistente... Skippo la copia")
				EffettuaCopia = False
				Rit = "ERROR: Skippata copia per problemi di mancanza file di origine"
				ScriveLogGlobale(NomeFileLog, Rit)
			End If

			'If Rit.Contains("ERROR") Then
			'	ScriveLogGlobale(NomeFileLog, Rit)
			'	Ritorno = Rit
			'Else
			If gf.EsisteFile(pathNuovaCategoria & "/" & NomeFile) Or EffettuaCopia = False Then
				If EffettuaCopia Then
					ScriveLogGlobale(NomeFileLog, "File copiato")

					gf.EliminaFileFisico(pathVecchiaCategoria & "/" & NomeFile)
					ScriveLogGlobale(NomeFileLog, "File di origine eliminato")
				End If

				ScriveLogGlobale(NomeFileLog, "Acquisizione nuovo id per categoria")
				Sql = "Select Coalesce(Max(Progressivo)+1,1) From Dati  Where idTipologia=" & idTipologia & " And idCategoria=" & idNuovaCategoria
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				Dim Progressivo As String = Rec(0).Value
				Rec.Close
				ScriveLogGlobale(NomeFileLog, "ID acquisito per categoria: " & Progressivo)

				ScriveLogGlobale(NomeFileLog, "Aggiornamento tabella dati")
				Sql = "Update dati Set progressivo=" & Progressivo & ", idCategoria=" & idNuovaCategoria & " Where idTipologia=" & idTipologia & " And idCategoria=" & idVecchiaCategoria & " And Progressivo=" & idMultimedia
				Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

				If Rit = "OK" Then
					ScriveLogGlobale(NomeFileLog, "Aggiornamento tabella effettuato")

					ScriveLogGlobale(NomeFileLog, "Aggiornamento tabella preferiti")
					Sql = "Update preferiti Set Progressivo=" & Progressivo & ", idCategoria=" & idNuovaCategoria & " Where idTipologia=" & idTipologia & " And idCategoria=" & idVecchiaCategoria & " And Progressivo=" & idMultimedia
					Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
					If Rit = "OK" Then
						ScriveLogGlobale(NomeFileLog, "Aggiornamento tabella preferiti effettuato")

						ScriveLogGlobale(NomeFileLog, "Aggiornamento tabella preferiti protetti")
						Sql = "Update preferitiprot Set Progressivo=" & Progressivo & ", idCategoria=" & idNuovaCategoria & " Where idTipologia=" & idTipologia & " And idCategoria=" & idVecchiaCategoria & " And Progressivo=" & idMultimedia
						Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
						If Rit = "OK" Then
							ScriveLogGlobale(NomeFileLog, "Aggiornamento tabella preferiti protetti effettuato")

							If idTipologia = "1" Or idTipologia = 1 Then
								ScriveLogGlobale(NomeFileLog, "Aggiornamento tabella informazioni immagini")
								Sql = "Update informazioniimmagini Set idMultimedia=" & Progressivo & ", idCategoria=" & idNuovaCategoria & " Where idCategoria=" & idVecchiaCategoria & " And idMultimedia=" & idMultimedia
								Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
								If Rit = "OK" Then
									ScriveLogGlobale(NomeFileLog, "Aggiornamento tabella informazioni immagini effettuato")

									Sql = "Update exifimmagini Set idMultimedia=" & Progressivo & ", idCategoria=" & idNuovaCategoria & " Where idCategoria=" & idVecchiaCategoria & " And idMultimedia=" & idMultimedia
									Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
									If Rit = "OK" Then
										ScriveLogGlobale(NomeFileLog, "Aggiornamento tabella quante aggiunte su categorie")

										Sql = "Update categorie Set QuanteAggiunte = QuanteAggiunte + 1 Where idCategoria=" & idNuovaCategoria & " And idTipologia=" & idTipologia
										Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
										If Rit = "OK" Then
											ScriveLogGlobale(NomeFileLog, "Aggiornamento tabella quante aggiunte su categorie")

											Ritorno = "*"
										Else
											ScriveLogGlobale(NomeFileLog, Sql)
											ScriveLogGlobale(NomeFileLog, Rit)
											Ritorno = Rit
										End If
									Else
										ScriveLogGlobale(NomeFileLog, Sql)
										ScriveLogGlobale(NomeFileLog, Rit)
										Ritorno = Rit
									End If
								Else
									ScriveLogGlobale(NomeFileLog, Sql)
									ScriveLogGlobale(NomeFileLog, Rit)
									Ritorno = Rit
								End If
							Else
								ScriveLogGlobale(NomeFileLog, "Aggiornamento tabella informazioni video")
								Sql = "Update InformazioniVideo Set idMultimedia=" & Progressivo & ", idCategoria=" & idNuovaCategoria & " Where idTipologia=" & idTipologia & " And idCategoria=" & idVecchiaCategoria & " And idMultimedia=" & idMultimedia
								Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
								If Rit = "OK" Then
									ScriveLogGlobale(NomeFileLog, "Aggiornamento tabella informazioni video effettuato")

									Ritorno = "*"
								Else
									ScriveLogGlobale(NomeFileLog, Sql)
									ScriveLogGlobale(NomeFileLog, Rit)
									Ritorno = Rit
								End If
							End If
						Else
							ScriveLogGlobale(NomeFileLog, Sql)
							ScriveLogGlobale(NomeFileLog, Rit)
							Ritorno = Rit
						End If
					Else
						ScriveLogGlobale(NomeFileLog, Sql)
						ScriveLogGlobale(NomeFileLog, Rit)
						Ritorno = Rit
					End If
				Else
					ScriveLogGlobale(NomeFileLog, Sql)
					ScriveLogGlobale(NomeFileLog, Rit)
					Ritorno = Rit
				End If
			Else
				ScriveLogGlobale(NomeFileLog, "File NON copiato")

				Ritorno = "ERROR: Nessun file copiato"
			End If
			'End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function CreaCategoria(idTipologia As String, NomeCategoria As String, Protetta As String, AggiuntaInutile As String, SottoCategoria As String) As String
		Dim Ritorno As String = ""
		Dim Sql As String
		Dim Barra As String = "\"
		Dim Db As New clsGestioneDB(TipoDB)
		Dim NomeFileLog As String = Server.MapPath(".") & "/Logs/CreazioneCategoria.txt"

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object

			ScriveLogGlobale(NomeFileLog, "------------------------------------")
			ScriveLogGlobale(NomeFileLog, "Ricerca nome categoria esistente: " & NomeCategoria)

			Sql = "Select * From Categorie Where Categoria='" & NomeCategoria & "'"
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			If Rec.Eof = False Then
				Ritorno = "ERROR: Categoria già esistente"
				ScriveLogGlobale(NomeFileLog, Ritorno)
			End If
			Rec.Close

			If Not Ritorno.Contains("ERROR:") Then
				ScriveLogGlobale(NomeFileLog, "Acquisizione idCategoria per idTipologia " & idTipologia)

				Sql = "Select Coalesce(Max(idCategoria)+1, 1) From Categorie Where idTipologia=" & idTipologia
				Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
				Dim idCategoria As String = Rec(0).Value
				Rec.Close
				ScriveLogGlobale(NomeFileLog, "Nuovo idCategoria " & idCategoria)

				Dim gf As New GestioneFilesDirectory
				Dim Path As String = ""
				Dim idSottoCategoria As Integer = -1

				If SottoCategoria <> "" Then
					Sql = "Select * From Categorie Where idTipologia=" & idTipologia & " And Categoria='" & SottoCategoria & "'"
					Dim Rec2 As Object
					Rec2 = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
					idSottoCategoria = Rec2("idCategoria").Value
					Rec2.Close

					ScriveLogGlobale(NomeFileLog, "ID sotto categoria: " & idSottoCategoria)

					Path &= "/var/www/html/CartelleCondivise/Fotacce/Categorie/" & SottoCategoria & "/" & (NomeCategoria.Replace(" ", "_").Replace("/", "_") & "")
				Else
					Path = "/var/www/html/CartelleCondivise/Fotacce/Categorie/" & (NomeCategoria.Replace(" ", "_").Replace("/", "_") & "")
				End If

				ScriveLogGlobale(NomeFileLog, "Path categoria: " & Path)

				gf.CreaDirectoryDaPercorso(Path & "/")

				ScriveLogGlobale(NomeFileLog, "Test scrittura nel path")
				Dim Rit As String = gf.CreaAggiornaFile(Path & "/Buttami.txt", "PROVA")
				If Rit.Contains(StringaErrore) Then
					ScriveLogGlobale(NomeFileLog, Rit)
					Ritorno = Rit
				Else
					ScriveLogGlobale(NomeFileLog, "Test di scrittura ok")
					gf.EliminaFileFisico(Path & "/Buttami.txt")

					'Dim NomeFile As String = ""

					'If idTipologia = "1" Then
					'	NomeFile = Server.MapPath(".") & "/PercorsiImmagini.txt"
					'Else
					'	NomeFile = Server.MapPath(".") & "/PercorsiVideo.txt"
					'End If
					'ScriveLogGlobale(NomeFileLog, "Aggiornamento file " & NomeFile)

					'Dim Filetto As String = gf.LeggeFileIntero(NomeFile)
					'Filetto &= NomeCategoria & ";" & Path & ";" & Protetta & ";" & AggiuntaInutile & "§"
					'ScriveLogGlobale(NomeFileLog, "Copia di backup file: " & NomeFile & ".bck")
					'gf.CopiaFileFisico(NomeFile, NomeFile & ".bck", True)
					'If gf.EsisteFile(NomeFile & ".bck") Then
					'	ScriveLogGlobale(NomeFileLog, "File di backup creato")
					'	gf.EliminaFileFisico(NomeFile)

					'	ScriveLogGlobale(NomeFileLog, "Scrittura nuova categoria su file")
					'	gf.CreaAggiornaFile(NomeFile, Filetto)

					ScriveLogGlobale(NomeFileLog, "Inserimento nuova categoria in tabella")
					Sql = "Insert Into categorie Values (" &
							" " & idCategoria & ", " &
							" " & idTipologia & ", " &
							"'" & NomeCategoria.Replace("'", "''") & "', " &
							"'" & Path & "', " &
							"'" & Protetta.ToUpper & "', " &
							"'--', " &
							"0, " &
							"'" & AggiuntaInutile & "', " &
							" " & idSottoCategoria & " " &
							")"
					Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

					If Rit = "OK" Then
						ScriveLogGlobale(NomeFileLog, "Inserimento nuova categoria in tabella effettuata")
						Return "*"
					Else
						ScriveLogGlobale(NomeFileLog, Sql)
						ScriveLogGlobale(NomeFileLog, Rit)
						Return Rit
					End If
					'Else
					'	ScriveLogGlobale(NomeFileLog, "File di backup NON creato")
					'	Ritorno = "ERROR: File di backup NON creato"
					'End If

				End If
			End If

		End If

		Return Ritorno
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

				'Dim CartellaIntermedia As String = PathOriginale
				'If CartellaIntermedia.Contains("CartelleCondivise") Then
				'	CartellaIntermedia = Mid(CartellaIntermedia, CartellaIntermedia.IndexOf("CartelleCondivise") + 18, CartellaIntermedia.Length)
				'End If
				'Dim daCopiare As String = Server.MapPath(".") & "BackupCancellazioni" & CartellaIntermedia & Rec("NomeFile").value

				Dim CartellaIntermedia As String = PathOriginale
				CartellaIntermedia = CartellaIntermedia.Replace("/Fotacce/", "/Fotacce/Categorie/BackupCancellazioni/")
				' Dim daCopiare As String = Server.MapPath(".") & "Piccole" & CartellaIntermedia & Rec("NomeFile").value
				Dim daCopiare As String = CartellaIntermedia & Rec("NomeFile").value
				'Dim PercorsoDestinazione As String = Server.MapPath(".") & "Piccole" ' & CartellaIntermedia
				Dim PercorsoDestinazione As String = CartellaIntermedia
				PercorsoDestinazione = PercorsoDestinazione.Substring(0, PercorsoDestinazione.IndexOf("BackupCancellazioni/") + 19)

				gf.CreaDirectoryDaPercorso(daCopiare)
				'gf.ImpostaAttributiFile(Server.MapPath(".") & "BackupCancellazioni" & CartellaIntermedia, FileAttribute.Normal)
				'ScriveLogGlobale(NomeFileLog, "Impostati attributi cartella di backup: " & Server.MapPath(".") & "BackupCancellazioni" & CartellaIntermedia)

				ScriveLogGlobale(NomeFileLog, "-----------------------------------------")
				ScriveLogGlobale(NomeFileLog, "Eliminazione idTipologia " & idTipologia & " idCategoria " & idCategoria & " idMultimedia " & idMultimedia)
				ScriveLogGlobale(NomeFileLog, "Nome File origine: " & filetto)
				ScriveLogGlobale(NomeFileLog, "Nome File destinazione: " & daCopiare)

				Dim Copio As Boolean = True
				Dim R As String = ""

				If gf.EsisteFile(filetto) Then
					Try
						R = gf.CopiaFileFisico(filetto, daCopiare, True)
					Catch ex As Exception
						'If giaProvatoACancellare = False Then
						'	Threading.Thread.Sleep(1000)
						'ScriveLogGlobale(NomeFileLog, "Problema con i permessi... Riprovo a eseguire la funzione. " & ex.Message)
						'	giaProvatoACancellare = True
						'	Dim Ritorno2 As String = EliminaMultimediaDaId(idTipologia, idCategoria, idMultimedia)
						Return "ERROR: " & ex.Message
						'Else
						'	giaProvatoACancellare = False
						'	Return "ERROR: " & ex.Message
						'End If
					End Try
				Else
					ScriveLogGlobale(NomeFileLog, "File origine non esistente")
					Copio = False
				End If

				If Copio Then
					If gf.EsisteFile(daCopiare) Then
						ScriveLogGlobale(NomeFileLog, "File copiato")
					Else
						Return "ERROR: Copia file non riuscita"
					End If
				End If

				If R.Contains("ERROR:") Then
					ScriveLogGlobale(NomeFileLog, R)
				Else
					If Copio Then
						ScriveLogGlobale(NomeFileLog, "File backuppato")
					End If

					'gf.ImpostaAttributiFile(filetto, FileAttribute.Normal)
					'ScriveLogGlobale(NomeFileLog, "Impostati attributi file")
					Ritorno = ""
					If gf.EsisteFile(filetto) Then
						Ritorno = gf.EliminaFileFisico(filetto)

						If Ritorno = "" Then
							ScriveLogGlobale(NomeFileLog, "File eliminato")
						End If
					End If

					If Ritorno = "" Then
						Sql = "Update dati Set Eliminata='S' Where idTipologia=" & idTipologia & " And idCategoria=" & idCategoria & " And Progressivo=" & idMultimedia
						Dim Rit As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

						If Rit = "OK" Then
							ScriveLogGlobale(NomeFileLog, "Eliminazione multimedia su tabella OK")

							If idTipologia = "1" Then
								Sql = "Update informazioniimmagini Set Eliminata='S' Where idCategoria=" & idCategoria & " And idMultimedia=" & idMultimedia
								Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
								If Rit = "OK" Then
									ScriveLogGlobale(NomeFileLog, "Eliminazione informazioni immagine su tabella OK")

									Sql = "Update exifimmagini Set Eliminata='S' Where idCategoria=" & idCategoria & " And idMultimedia=" & idMultimedia
									Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
									If Rit = "OK" Then
										ScriveLogGlobale(NomeFileLog, "Eliminazione exif immagini su tabella OK")

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
					"Where A.idTipologia=1 And A.idCategoria = " & idCategoriaGlobalePerConversione & " And (A.Eliminata='N' Or A.Eliminata='n') And B.Sezione1 is Null" '  And (B.Eliminata='N' Or B.Eliminata='n') "
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
	Public Function CalcolaPuntiniImmaginePerTest() As String
		Dim Ritorno As String = ""
		Dim NomeFileLog As String = Server.MapPath(".") & "/Logs/RitornaPuntiniImmagine.txt"
		Dim gi As New GestioneImmagini
		Dim gf As New GestioneFilesDirectory
		gf.ScansionaDirectorySingola("D:\MP3daPassare\ImmaginiUguali")
		Dim qFiles As Integer = gf.RitornaQuantiFilesRilevati
		Dim Files() As String = gf.RitornaFilesRilevati

		For i As Integer = 1 To qFiles
			Dim NomeFile As String = Files(i)
			Dim Puntini1 As StrutturaJPG = gi.CalcolaPuntini(Server.MapPath("."), NomeFile, NomeFileLog, i, True, False)

			If Puntini1.Sezione1.Contains("ERROR") Then
				ScriveLogGlobale(NomeFileLog, Puntini1.Sezione1)
			Else
				Ritorno &= vbCrLf & Puntini1.PuntiDiagonale & ";" & Puntini1.PuntiCornice & ";" & Puntini1.Punti & ";" & Puntini1.Sezione1 & ";" & Puntini1.Width & ";" & Puntini1.Height & ";" &
					Puntini1.DataOra & ";" & Puntini1.Sezione2 & ";" & Puntini1.Descrizione & ";" & Puntini1.Commento & ";" & Puntini1.DateTimeTag & ";" & Puntini1.Software & ";" &
					Puntini1.XRes & ";" & Puntini1.YRes & ";" & Puntini1.Hash
			End If

			If i > 5 Then
				Exit For
			End If
		Next

		gf.EliminaFileFisico(Server.MapPath(".") & "/Appoggio/Ritorno.txt")
		gf.CreaAggiornaFile(Server.MapPath(".") & "/Appoggio/Ritorno.txt", Ritorno)

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function CalcolaPuntiniImmagine2(idCategoria As String, Refresh As String) As String
		Dim Ritorno As String = "*"

		If staCaricandoPuntini = True Then
			Ritorno = "ERROR: Caricamento punti già in corso"
		Else
			idCategoriaGlobalePerPuntini = idCategoria
			RefreshPerPuntini = Refresh

			timerConvP = New Timers.Timer(100)
			AddHandler timerConvP.Elapsed, New ElapsedEventHandler(AddressOf CalcolaPuntiniImmagine)
			timerConvP.Start()
		End If

		Return Ritorno
	End Function

	Public Sub CalcolaPuntiniImmagine()
		Dim Debug As Boolean = False

		timerConvP.Stop()

		staCaricandoPuntini = True
		Dim idCategoria As String = idCategoriaGlobalePerPuntini
		Dim Refresh As String = RefreshPerPuntini

		Dim Db As New clsGestioneDB(TipoDB)
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		Dim gi As New GestioneImmagini
		Dim gf As New GestioneFilesDirectory
		Dim NomeFileLog As String = Server.MapPath(".") & "/Logs/RitornaPuntiniImmagine.txt"
		Dim Ritorno As String = "*"

		gf.EliminaFileFisico(NomeFileLog)

		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim Rec2 As Object
			Dim Sql As String = ""

			Dim NumeroImmagineConvertita As Long = 0
			Dim Altro As String = ""

			If idCategoria <> -1 Then
				Altro = " Where idCategoria=" & idCategoria
			End If

			If Refresh = "S" Then
				Sql = "Update informazioniimmagini Set PuntiDiagonale=null, PuntiCornice=null, Punti=null, Sezione1=null, Sezione2=null " & Altro
				Dim Rit As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

				If Rit = "OK" Then
					ScriveLogGlobale(NomeFileLog, "Eliminazione tabella effettuata")
				Else
					ScriveLogGlobale(NomeFileLog, Rit)
					Exit Sub
				End If
			End If

			Dim q As Integer = 0
			' Acquisizione puntini per ricerca più precisa
			Altro = ""
			If idCategoria <> -1 Then
				Altro = " And A.idCategoria=" & idCategoria
			End If

			Sql = "Select * From dati A " &
				"Left Join categorie B On B.idTipologia = 1 And A.idCategoria = B.idCategoria " &
				"Left Join informazioniimmagini C On C.idcategoria = A.idCategoria And A.progressivo = C.idMultimedia " &
				"Where (PuntiDiagonale is Null Or PuntiDiagonale = '') And (A.Eliminata='N' Or A.Eliminata='n') And A.idTipologia=1" & Altro
			'ScriveLogGlobale(NomeFileLog, Sql)

			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL, False)
			Do Until Rec.Eof
				Dim Percorso As String = Rec("Percorso").Value
				Dim NomeFile As String = Rec("NomeFile").Value
				Dim idMultimedia As String = Rec("Progressivo").Value
				Dim idCategoriaLetta As String = Rec("idCategoria").Value

				If Not Percorso.ToUpper.Contains("/PICCOLE") And Not NomeFile.ToUpper.Contains("/PICCOLE") Then
					ScriveLogGlobale(NomeFileLog, "-----------------------------------------")
					ScriveLogGlobale(NomeFileLog, "Percorso: " & Percorso)
					ScriveLogGlobale(NomeFileLog, "Nome File: " & NomeFile)
					' ScriveLogGlobale(NomeFileLog, "Acquisizione puntini " & Percorso & "/" & NomeFile)
					If idCategoria <> -1 Then
						ScriveLogGlobale(NomeFileLog, "per categoria " & idCategoria)
					Else
						ScriveLogGlobale(NomeFileLog, "per tutte le categorie")
					End If

					NumeroImmagineConvertita += 1
					Dim Puntini As StrutturaJPG = gi.CalcolaPuntini(Server.MapPath("."), Percorso & "/" & NomeFile, NomeFileLog, NumeroImmagineConvertita)
					If Puntini.Sezione1.Contains("ERROR") Then
						ScriveLogGlobale(NomeFileLog, Puntini.Sezione1)
					Else
						'Dim Pv() As String = Puntini.Split(";")

						'Dim PuntiDiagonale As String = Pv(0)
						'Dim PuntiCornice As String = Pv(1)
						'Dim PuntiCorpo As String = Pv(2)
						'Dim Hash As String = Pv(3)
						'Dim Width As String = Pv(4)
						'Dim Height As String = Pv(5)
						'Dim DataOra As String = Pv(6)
						'Dim HashColore As String = Pv(7)

						ScriveLogGlobale(NomeFileLog, "Acquisizione puntini. Punti diagonale " & Puntini.PuntiDiagonale)
						ScriveLogGlobale(NomeFileLog, "Acquisizione puntini. Punti cornice " & Puntini.PuntiCornice)
						ScriveLogGlobale(NomeFileLog, "Acquisizione puntini. Punti corpo " & Puntini.Punti)
						ScriveLogGlobale(NomeFileLog, "Acquisizione puntini. Punti negativi " & Puntini.Sezione1)
						ScriveLogGlobale(NomeFileLog, "Acquisizione puntini. Punti essenziali " & Puntini.Sezione2)
						If idCategoria = -1 Then
							ScriveLogGlobale(NomeFileLog, "Acquisizione puntini. ID Categoria Integer tabella: " & idCategoriaLetta)
						End If

						ScriveLogGlobale(NomeFileLog, "Controllo esistenza riga su informazioni immagine")
						Sql = "Select * From informazioniimmagini Where idCategoria=" & idCategoriaLetta & " And idMultimedia=" & idMultimedia
						Rec2 = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL, False)
						If Rec2.Eof = True Then
							ScriveLogGlobale(NomeFileLog, "Riga NON esistente. La aggiungo")
							Sql = "Insert Into informazioniimmagini Values (" &
								" " & idCategoriaLetta & ", " &
								" " & idMultimedia & ", " &
								"'" & Puntini.Sezione1 & "', " &
								"'" & Puntini.Punti & "', " &
								" " & Puntini.Width & ", " &
								" " & Puntini.Height & ", " &
								"'" & Puntini.DataOra & "', " &
								"'" & Puntini.PuntiDiagonale & "', " &
								"'" & Puntini.PuntiCornice & "', " &
								"'N', " &
								"'" & Puntini.Sezione2 & "', " &
								"'" & Puntini.Hash & "' " &
								")"
						Else
							ScriveLogGlobale(NomeFileLog, "Riga esistente. La modifico")
							Sql = "Update informazioniimmagini " &
								"Set PuntiDiagonale='" & Puntini.PuntiDiagonale & "', " &
								"PuntiCornice='" & Puntini.PuntiCornice & "', " &
								"Punti='" & Puntini.Punti & "', " &
								"Sezione1='" & Puntini.Sezione1 & "', Width=" & Puntini.Width & ", Height=" & Puntini.Height & ", DataOra='" & Puntini.DataOra & "', Sezione2='" & Puntini.Sezione2 & "', Hash='" & Puntini.Hash & "' " &
								"Where idCategoria=" & idCategoriaLetta & " And idMultimedia=" & idMultimedia
						End If
						Rec2.Close

						Dim Rit As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

						If Rit = "OK" Then
							ScriveLogGlobale(NomeFileLog, "Scrittura 1 in tabella effettuata")

							'Sql = "Delete From exifimmagini where idCategoria=" & idCategoriaLetta & " And idMultimedia=" & idMultimedia
							'Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
							'If Rit = "OK" Then
							'Try
							'	Sql = "Insert Into exifimmagini Values (" &
							'		" " & idCategoriaLetta & ", " &
							'		" " & idMultimedia & ", " &
							'		"'" & Puntini.Descrizione & "', " &
							'		"'" & Puntini.Commento & "', " &
							'		"'" & Puntini.XRes & "', " &
							'		"'" & Puntini.YRes & "', " &
							'		"'" & Puntini.DataOra & "', " &
							'		"'" & Puntini.Software & "' " &
							'		")"
							'	ScriveLogGlobale(NomeFileLog, Sql)
							'	Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
							'	'ScriveLogGlobale(NomeFileLog, Rit)
							'	If Rit = "OK" Then
							'		ScriveLogGlobale(NomeFileLog, "Scrittura 2 in tabella effettuata")
							'	Else
							'		ScriveLogGlobale(NomeFileLog, Sql)
							'		ScriveLogGlobale(NomeFileLog, Rit)
							'		Exit Do
							'	End If
							'Catch ex As Exception
							'	ScriveLogGlobale(NomeFileLog, "ERROR: " & ex.Message)
							'	Exit Do
							'End Try
							'Else
							'		ScriveLogGlobale(NomeFileLog, Sql)
							'		ScriveLogGlobale(NomeFileLog, Rit)
							'		Exit Do
							'	End If
						Else
							ScriveLogGlobale(NomeFileLog, Sql)
							ScriveLogGlobale(NomeFileLog, Rit)
							Exit Do
						End If
					End If

					If Debug = True Then
						q += 1
						If q > 14 Then
							Exit Do
						End If
					End If
				End If

				Rec.MoveNext
			Loop
			Rec.Close

			Ritorno = "*"
			' Acquisizione puntini per ricerca più precisa
		End If
		ScriveLogGlobale(NomeFileLog, "Fine elaborazione")
		staCaricandoPuntini = False

		'Return Ritorno
	End Sub

	<WebMethod()>
	Public Function PrendeExifImmagini(idCategoria As String, Refresh As String) As String
		Dim Ritorno As String = "*"

		If staCaricandoExif = True Then
			Ritorno = "ERROR: Caricamento punti già in corso"
		Else
			idCategoriaGlobalePerExif = idCategoria
			RefreshPerExif = Refresh

			timerConvE = New Timers.Timer(100)
			AddHandler timerConvE.Elapsed, New ElapsedEventHandler(AddressOf PrendeExifImmagini2)
			timerConvE.Start()
		End If

		Return Ritorno
	End Function

	Private Sub PrendeExifImmagini2()
		timerConvE.Stop()

		Dim idCategoria As String = idCategoriaGlobalePerExif
		Dim Refresh As String = RefreshPerExif

		Dim NomeFileLog As String = Server.MapPath(".") & "/Logs/PrendeExifImmagini.txt"
		Dim Ritorno As String = ""
		Dim gi As New GestioneImmagini
		Dim gf As New GestioneFilesDirectory
		Dim Db As New clsGestioneDB(TipoDB)

		staCaricandoExif = True

		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim Altro As String = ""
			If idCategoria <> "-1" Then
				Altro = " And A.idCategoria=" & idCategoria
			End If
			Dim Sql As String = ""

			ScriveLogGlobale(NomeFileLog, "---------------------------------------------------")
			If Refresh = "S" Then
				ScriveLogGlobale(NomeFileLog, "Eliminazione dati in tabella")

				Sql = "Delete From exifimmagini"
				If idCategoria <> "-1" Then
					Sql &= " Where idCategoria=" & idCategoria
				End If
				Dim Rit As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
				If Rit = "OK" Then
					ScriveLogGlobale(NomeFileLog, "Dati eliminati da tabella")
				Else
					ScriveLogGlobale(NomeFileLog, Rit)
				End If
			End If

			ScriveLogGlobale(NomeFileLog, "Lettura dati da inserire")
			Sql = "Select * From dati A " &
				"Left Join Categorie B On A.idTipologia=B.idTipologia And A.idCategoria=B.idCategoria " &
				"Left Join exifimmagini C On A.idCategoria=C.idCategoria And A.progressivo=C.idMultimedia " &
				"Where A.idTipologia=1" & Altro & " And C.idCategoria Is Null"
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			If TypeOf (Rec) Is String Then
				ScriveLogGlobale(NomeFileLog, "ERROR: " & Rec)
				staCaricandoExif = False
				Exit Sub
			Else
				Do Until Rec.Eof
					Dim Percorso As String = Rec("Percorso").Value
					If Strings.Right(Percorso, 1) = "/" Then
						Percorso = Mid(Percorso, 1, Percorso.Length - 1)
					End If
					Dim Nome As String = Rec("NomeFile").Value
					Dim NomeFile As String = Percorso & "/" & Nome

					Dim RitornoHash As StrutturaJPG = gi.AcquisisceExif(NomeFileLog, NomeFile)
					ScriveLogGlobale(NomeFileLog, "Lettura exif completata: " & RitornoHash.Descrizione)
					'If RitornoHash.Descrizione <> "" Then
					Sql = "Insert Into exifimmagini Values (" &
							" " & Rec("idCategoria").Value & ", " &
							" " & Rec("progressivo").Value & ", " &
							"'" & RitornoHash.Descrizione & "', " &
							"'" & RitornoHash.Commento & "', " &
							"'" & RitornoHash.XRes & "', " &
							"'" & RitornoHash.YRes & "', " &
							"'" & RitornoHash.DataOra & "', " &
							"'" & RitornoHash.Software & "' " &
							")"
					Dim Rit As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
					If Rit = "OK" Then
						ScriveLogGlobale(NomeFileLog, "Scrittura effettuata in tabella")
					Else
						ScriveLogGlobale(NomeFileLog, Rit)
					End If
					'End If

					Rec.MoveNext
				Loop
				Rec.Close
			End If
		End If

		staCaricandoExif = False
	End Sub

	<WebMethod()>
	Public Function CalcolaHashImmagine(idCategoria As String, idMultimedia As String, Refresh As String, NumeroImmagineConvertita As String) As String
		Dim NomeFileLog As String = Server.MapPath(".") & "/Logs/RitornaInformazioniImmagine.txt"
		Dim Ritorno As String = ""
		Dim gi As New GestioneImmagini
		Dim gf As New GestioneFilesDirectory
		Dim Db As New clsGestioneDB(TipoDB)
		Dim Debug As Boolean = True

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
						Ritorno = Rec("Sezione1").Value & ";" & Rec("Punti").Value & ";" & Rec("Width").Value & ";" & Rec("Height").Value & ";" & Rec("DataOra").Value
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

							Dim RitornoHash As StrutturaJPG = gi.CalcolaPuntini(Server.MapPath("."), Percorso & "/" & NomeFile, NomeFileLog, NumeroImmagineConvertita)

							If RitornoHash.Sezione1.Contains("ERROR:") Then
								Ritorno = RitornoHash.Sezione1

								''ScriveLogGlobale(NomeFileLog, Ritorno)
							Else
								ScriveLogGlobale(NomeFileLog, "Acquisizione informazioni immagine effettuata")

								Sql = "Delete From informazioniimmagini Where idCategoria=" & idCategoria & " And idMultimedia=" & idMultimedia
								Dim Rit As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
								If Rit = "OK" Then
									Sql = "Insert Into informazioniimmagini Values (" & idCategoria & ", " & idMultimedia & ", '" & RitornoHash.Sezione1 & "', " & RitornoHash.Punti & ", " & RitornoHash.Width & ", " & RitornoHash.Height & ", '" & RitornoHash.DataOra & "', " & RitornoHash.PuntiDiagonale & ", " & RitornoHash.PuntiCornice & ", '" & RitornoHash.Sezione2 & "', '" & RitornoHash.Hash & "', 'N')"
									Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

									If Rit = "OK" Then
										ScriveLogGlobale(NomeFileLog, "Scrittura 1 effettuata in tabella")

										'Sql = "Insert Into exifimmagini Values (" &
										'	" " & idCategoria & ", " &
										'	" " & idMultimedia & ", " &
										'	"'" & RitornoHash.Descrizione.Replace("'", "''") & "', " &
										'	"'" & RitornoHash.Commento.Replace("'", "''") & "', " &
										'	"'" & RitornoHash.XRes.Replace("'", "''") & "', " &
										'	"'" & RitornoHash.YRes.Replace("'", "''") & "', " &
										'	"'" & RitornoHash.DataOra.Replace("'", "''") & "', " &
										'	"'" & RitornoHash.Software.Replace("'", "''") & "' " &
										'	")"
										'Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

										'If Rit = "OK" Then
										'	ScriveLogGlobale(NomeFileLog, "Scrittura 2 effettuata in tabella")

										Ritorno = RitornoHash.Sezione1 & ";" & RitornoHash.Punti & ";" & RitornoHash.Width & ";" & RitornoHash.Height & ";" & RitornoHash.DataOra & ";" & RitornoHash.Sezione2 & ";" & RitornoHash.Hash
										'Else
										'	Ritorno = "ERROR: Errore sull'inserimento dei dati in tabella 2. " & Sql

										'	ScriveLogGlobale(NomeFileLog, Ritorno)
										'End If
									Else
										Ritorno = "ERROR: Errore sull'inserimento dei dati in tabella 1. " & Sql

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

	<WebMethod()>
	Public Function TrovaImmagini1280(idCategoria As String, Inizio As String, QuanteImmagini As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		Dim Ritorno As String = ""
		Dim NomeFileLog As String = Server.MapPath(".") & "/Logs/Immagini1280.txt"

		If ConnessioneSQL <> "" Then
			ScriveLogGlobale(NomeFileLog, "Cerca stringa 1280")
			Ritorno &= CercaStringa(Server.MapPath("."), Db, ConnessioneSQL, NomeFileLog, idCategoria, " 1280", Inizio, QuanteImmagini, True)

			ScriveLogGlobale(NomeFileLog, "Cerca stringa 500")
			Ritorno &= CercaStringa(Server.MapPath("."), Db, ConnessioneSQL, NomeFileLog, idCategoria, " 500", Inizio, QuanteImmagini, True)

			ScriveLogGlobale(NomeFileLog, "Cerca stringa 400")
			Ritorno &= CercaStringa(Server.MapPath("."), Db, ConnessioneSQL, NomeFileLog, idCategoria, " 400", Inizio, QuanteImmagini, True)

			ScriveLogGlobale(NomeFileLog, "Cerca stringa 0001")
			Ritorno &= CercaStringa(Server.MapPath("."), Db, ConnessioneSQL, NomeFileLog, idCategoria, " 0001", Inizio, QuanteImmagini, False)
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function TrovaImmaginiUguali(idCategoria As String, ricercaPerMetodo1 As String, ricercaPerMetodo2 As String, ricercaPerData As String, ricercaPerDimensioni As String, ricercaPerPunti As String,
										ricercaPerPuntiDiagonale As String, ricercaPerPuntiCornice As String, ricercaPerPuntiCorpo As String, ricercaPerNomeUguale As String,
										ricercaPerPeso As String, ricercaPerStringa As String, stringaRicerca As String, QuanteImmagini As String, Inizio As String, AndOr As String,
										TutteLeCategorie As String, ricercaPerNegativo As String, ricercaPerEssenziale As String, ricercaPer1280 As String, Caratteri As String,
										Ordinamento As String, ricercaPerHash As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		Dim Ritorno As String = "*"

		If ConnessioneSQL <> "" Then
			Dim RitornoHash As String = ""
			Dim RitornoHashColore As String = ""
			Dim RitornoPunti As String = ""
			Dim RitornoDataOra As String = ""
			Dim RitornoDimensioni As String = ""
			Dim RitornoNomeUguale As String = ""
			Dim RitornoPeso As String = ""
			Dim RitornoStringa As String = ""
			Dim Ritorno1280 As String = ""
			Dim RitornoEssenziale As String = ""

			If ricercaPer1280 = "S" Then
				Ritorno1280 = TrovaImmagini1280(IIf(TutteLeCategorie = "S", "-1", idCategoria), Inizio, QuanteImmagini)
			End If
			Ritorno1280 &= "|"

			If ricercaPerMetodo1 = "S" Then
				RitornoHash = RitornaUguaglianze(Server.MapPath("."), Db, ConnessioneSQL, "Exif", "EXIF DESCR.", idCategoria, QuanteImmagini, Inizio, "", AndOr, TutteLeCategorie, Caratteri, Ordinamento)
			End If
			RitornoHash &= "|"

			If ricercaPerMetodo2 = "S" Then
				RitornoHashColore = RitornaUguaglianze(Server.MapPath("."), Db, ConnessioneSQL, "Exif", "EXIF COMMENTO", idCategoria, QuanteImmagini, Inizio, "", AndOr, TutteLeCategorie, Caratteri, Ordinamento)
			End If
			RitornoHashColore &= "|"

			If ricercaPerData = "S" Then
				RitornoDataOra = RitornaUguaglianze(Server.MapPath("."), Db, ConnessioneSQL, "DataOra", "DATA ORA", idCategoria, QuanteImmagini, Inizio, "", AndOr, TutteLeCategorie, Caratteri, Ordinamento)
			End If
			RitornoDataOra &= "|"

			If ricercaPerDimensioni = "S" Then
				RitornoDimensioni = RitornaUguaglianze(Server.MapPath("."), Db, ConnessioneSQL, "Concat(Width, 'x', height)", "DIMENSIONI", idCategoria, QuanteImmagini, Inizio, "", AndOr, TutteLeCategorie, Caratteri, Ordinamento)
			End If
			RitornoDimensioni &= "|"

			If ricercaPerEssenziale = "S" Then
				RitornoEssenziale = RitornaEssenziali(Server.MapPath("."), Db, ConnessioneSQL, idCategoria, Inizio, QuanteImmagini, TutteLeCategorie, "DIMENSIONI", Ordinamento)
			End If
			RitornoEssenziale &= "|"

			If ricercaPerPunti = "S" Then
				Dim aTipo = New List(Of String)
				Dim Tipo As String = ""
				Dim q As Integer = 0
				Dim Fai As Boolean = True

				If ricercaPerPuntiDiagonale = "S" Then
					aTipo.Add("PuntiDiagonale")
				End If
				If ricercaPerPuntiCornice = "S" Then
					aTipo.Add("PuntiCornice")
				End If
				If ricercaPerPuntiCorpo = "S" Then
					aTipo.Add("Punti")
				End If
				If ricercaPerNegativo = "S" Then
					aTipo.Add("Sezione1")
				End If
				If ricercaPerHash = "S" Then
					aTipo.Add("Hash")
				End If

				For Each t As String In aTipo
					Tipo &= t & ", '-',"
					q += 1
				Next
				If Tipo.Length > 0 Then
					Tipo = Mid(Tipo, 1, Tipo.Length - 6)
				End If

				If q > 1 Then
					Tipo = "Concat(" & Tipo & ")"
				Else
					If q = 0 Then
						Fai = False
					End If
				End If

				If Fai Then
					RitornoPunti = RitornaUguaglianze(Server.MapPath("."), Db, ConnessioneSQL, Tipo, "PUNTI", idCategoria, QuanteImmagini, Inizio, "", AndOr, TutteLeCategorie, Caratteri, Ordinamento)
				End If
			End If
			RitornoPunti &= "|"

			If ricercaPerNomeUguale = "S" Then
				RitornoNomeUguale = RitornaUguaglianze(Server.MapPath("."), Db, ConnessioneSQL, "B.solonome", "NOME UGUALE", idCategoria, QuanteImmagini, Inizio, "", AndOr, TutteLeCategorie, Caratteri, Ordinamento)
			End If
			RitornoNomeUguale &= "|"

			If ricercaPerPeso = "S" Then
				RitornoPeso = RitornaUguaglianze(Server.MapPath("."), Db, ConnessioneSQL, "B.Dimensioni", "PESO", idCategoria, QuanteImmagini, Inizio, "", AndOr, TutteLeCategorie, Caratteri, Ordinamento)
			End If
			RitornoPeso &= "|"

			If ricercaPerStringa = "S" Then
				RitornoStringa = RitornaUguaglianze(Server.MapPath("."), Db, ConnessioneSQL, "STRINGA", "STRINGA", idCategoria, QuanteImmagini, Inizio, stringaRicerca, AndOr, TutteLeCategorie, Caratteri, Ordinamento)
			End If
			RitornoStringa &= "|"

			Ritorno = RitornoHash & RitornoDataOra & RitornoDimensioni & RitornoPunti & RitornoNomeUguale & RitornoPeso & RitornoStringa & RitornoHashColore & Ritorno1280 & RitornoEssenziale
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function TrovaImmaginiPiccole(idCategoria As String, QuanteImmagini As String, Inizio As String, Dimensioni As String, Width As String, Height As String, TutteLeCategorie As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		Dim Ritorno As String = ""

		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim Sql As String = ""
			Dim RicercaCategoria As String = ""

			If TutteLeCategorie = "N" Or TutteLeCategorie = "" Then
				RicercaCategoria = "And A.idcategoria = " & idCategoria & " "
			End If

			Sql = "Select Coalesce(Count(*),0) From dati As A " &
				"Left Join informazioniimmagini B On A.idcategoria = B.idCategoria And A.progressivo = B.idMultimedia " &
				"Left Join categorie C On A.idtipologia = C.idtipologia And A.idcategoria = C.idcategoria " &
				"left join preferiti d On d.idTipologia = 1 And A.idCategoria = d.idCategoria And A.progressivo=d.progressivo " &
				"left join preferitiprot e On e.idTipologia = 1 And A.idCategoria = e.idCategoria And A.progressivo=e.progressivo " &
				"Where A.idtipologia = 1 " & RicercaCategoria & " And (A.dimensioni < " & Dimensioni & " Or (B.Width < " & Width & " And B.Height < " & Height & ") Or Instr(A.NomeFile, 'PICCOLE') > 0) " &
				"And (A.Eliminata = 'N' Or A.Eliminata = 'n') And (B.Eliminata = 'N' Or B.Eliminata = 'n')"
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			If TypeOf (Rec) Is String Then
				Return Rec
			End If

			Dim QuanteRigheTotali As Integer = -1

			If Rec.Eof = True Then
				Ritorno = "ERROR: Nessun file rilevato"
			Else
				QuanteRigheTotali = Rec(0).Value
				Rec.Close
			End If

			Sql = "Select * From (" &
				"Select ROW_NUMBER() OVER(Order BY A.progressivo) As NumeroRiga, A.progressivo, B.DataOra, C.percorso, C.Protetta, A.nomefile, A.dimensioni, " &
				"Coalesce(B.Width, 0) As Width, Coalesce(B.Height, 0) As Height, Coalesce(d.progressivo, '') As preferito, Coalesce(e.progressivo, '') As preferitoprot, A.solonome, A.idcategoria From dati As A " &
				"Left Join informazioniimmagini B On A.idcategoria = B.idCategoria And A.progressivo = B.idMultimedia " &
				"Left Join categorie C On A.idtipologia = C.idtipologia And A.idcategoria = C.idcategoria " &
				"left join preferiti d On d.idTipologia = 1 And A.idCategoria = d.idCategoria And A.progressivo=d.progressivo " &
				"left join preferitiprot e On e.idTipologia = 1 And A.idCategoria = e.idCategoria And A.progressivo=e.progressivo " &
				"Where A.idtipologia = 1 " & RicercaCategoria & " And (A.dimensioni < " & Dimensioni & " Or (B.Width < " & Width & " And B.Height < " & Height & ") Or Instr(percorso, 'PICCOLE') > 0) " & ' And B.idCategoria Is Not Null " &
				"And (A.Eliminata = 'N' Or A.Eliminata = 'n') And (B.Eliminata = 'N' Or B.Eliminata = 'n') " &
				"Order By A.dimensioni, B.Width, B.Height, A.solonome " &
				") As A Where NumeroRiga > " & (Val(Inizio) - 1) & " And NumeroRiga < " & (Inizio + Val(QuanteImmagini))
			' Return Sql
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
			If TypeOf (Rec) Is String Then
				Return Rec
			End If

			If Rec.Eof = True Then
				Ritorno = "ERROR: Nessun file rilevato"
			Else
				Do Until Rec.Eof
					Dim P As String = Rec("Percorso").Value
					Dim N As String = Rec("NomeFile").Value

					If Not P.ToUpper.Contains("/VIDEOS/") And Not N.ToUpper.Contains("/VIDEOS/") Then
						Ritorno &= Rec("Progressivo").Value & ";"
						Ritorno &= P.Replace(";", "---PV---") & ";"
						Ritorno &= N.Replace(";", "---PV---") & ";"
						Ritorno &= Rec("Preferito").Value & ";"
						Ritorno &= Rec("PreferitoProt").Value & ";"
						Ritorno &= Rec("dimensioni").Value & ";"
						Ritorno &= Rec("Width").Value & ";"
						Ritorno &= Rec("Height").Value & ";"
						Ritorno &= Rec("DataOra").Value & ";"
						Ritorno &= Rec("Protetta").Value & ";"
						Ritorno &= Rec("idCategoria").Value & ";"
						Ritorno &= "§"
					End If

					Rec.MoveNext
				Loop
				Rec.Close

				Ritorno &= "*" & QuanteRigheTotali
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function SpostaImmaginePiccole(idCategoria As String, idMultiMedia As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		Dim Ritorno As String = "*"
		Dim NomeFileLog As String = Server.MapPath(".") & "/Logs/SpostamentoImmagine.txt"
		Dim Barra As String = ""

		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim Sql As String = ""

			If TipoDB = "SQLSERVER" Then
				Barra = "\"
			Else
				Barra = "/"
			End If

			Sql = "Select Dati.NomeFile, Dati.Dimensioni, Dati.Data, Dati.idCategoria, Categorie.Categoria, Categorie.Percorso, Categorie.LetteraDisco From Dati " &
				"Left Join Categorie On Dati.idCategoria=Categorie.idCategoria And Dati.idTipologia=Categorie.idTipologia " &
				"Where Dati.idTipologia=1 And Dati.idCategoria=" & idCategoria & " And Progressivo=" & idMultiMedia
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
				CartellaIntermedia = CartellaIntermedia.Replace("/Fotacce/", "/Fotacce/Categorie/Piccole/")
				' Dim daCopiare As String = Server.MapPath(".") & "Piccole" & CartellaIntermedia & Rec("NomeFile").value
				Dim daCopiare As String = CartellaIntermedia & Rec("NomeFile").value
				'Dim PercorsoDestinazione As String = Server.MapPath(".") & "Piccole" ' & CartellaIntermedia
				Dim PercorsoDestinazione As String = CartellaIntermedia
				PercorsoDestinazione = PercorsoDestinazione.Substring(0, PercorsoDestinazione.IndexOf("Piccole/") + 7)

				gf.CreaDirectoryDaPercorso(daCopiare)
				gf.ImpostaAttributiFile(Server.MapPath(".") & "Piccole" & CartellaIntermedia, FileAttribute.Normal)
				ScriveLogGlobale(NomeFileLog, "Impostati attributi cartella di backup: " & Server.MapPath(".") & "BackupCancellazioni" & CartellaIntermedia)

				ScriveLogGlobale(NomeFileLog, "-----------------------------------------")
				ScriveLogGlobale(NomeFileLog, "Spostamento idCategoria " & idCategoria & " idMultimedia " & idMultiMedia)
				ScriveLogGlobale(NomeFileLog, "Nome File origine: " & filetto)
				ScriveLogGlobale(NomeFileLog, "Nome File destinazione: " & daCopiare)
				ScriveLogGlobale(NomeFileLog, "Percorso categoria destinazione: " & PercorsoDestinazione)

				Dim R As String = ""
				Try
					R = gf.CopiaFileFisico(filetto, daCopiare, True)
				Catch ex As Exception
					'If giaProvatoACancellare = False Then
					'	Threading.Thread.Sleep(1000)
					ScriveLogGlobale(NomeFileLog, "Problema con i permessi... Riprovo a eseguire la funzione di spostamento: " & ex.Message)
					'	giaProvatoACancellare = True
					'	Dim Ritorno2 As String = SpostaImmaginePiccole(idCategoria, idMultiMedia)
					'	Return Ritorno2
					'Else
					'	giaProvatoACancellare = False
					Return "ERROR: " & ex.Message
					'End If
				End Try
				If gf.EsisteFile(daCopiare) Then
					ScriveLogGlobale(NomeFileLog, "File copiato")
				Else
					Return "ERROR: Copia file non riuscita"
				End If

				gf.ImpostaAttributiFile(filetto, FileAttribute.Normal)
				Dim Rit As String = gf.EliminaFileFisico(filetto)
				ScriveLogGlobale(NomeFileLog, "File eliminato " & Rit)

				If Rit = "" Then
					Dim idCategoriaPiccole As String = ""

					Sql = "Select * From Categorie Where idTipologia=1 And Categoria='Piccole'"
					Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
					If TypeOf (Rec) Is String Then
						Return Rec
					End If

					If Rec.Eof = True Then
						Rec.Close
						ScriveLogGlobale(NomeFileLog, "Categoria Non Trovata. La inserisco")

						Sql = "Select Coalesce(Max(idCategoria) + 1, 1) From Categorie Where idTipologia = 1"
						Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
						If TypeOf (Rec) Is String Then
							Return Rec
						End If
						idCategoriaPiccole = Rec(0).Value
						Rec.Close
						ScriveLogGlobale(NomeFileLog, "Categoria Non Trovata. idCategoria: " & idCategoriaPiccole)

						Sql = "Insert Into categorie Values (" & idCategoriaPiccole & ", 1, 'Piccole', '" & PercorsoDestinazione & "', 'S', '--')"
						Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

						If Rit <> "OK" Then
							Return Rit
						End If
					Else
						idCategoriaPiccole = Rec("idCategoria").Value
						ScriveLogGlobale(NomeFileLog, "Categoria Trovata. idCategoria: " & idCategoriaPiccole)
						Rec.Close
					End If

					ScriveLogGlobale(NomeFileLog, "Aggiorno categoria immagine")
					Sql = "Update dati Set idCategoria=" & idCategoriaPiccole & " Where idTipologia=1 And idCategoria=" & idCategoria & " And Progressivo=" & idMultiMedia
					Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

					If Rit <> "OK" Then
						ScriveLogGlobale(NomeFileLog, Rit)
						Return Rit
					Else
						ScriveLogGlobale(NomeFileLog, "Aggiorno categoria preferito")
						Sql = "Update preferiti Set idCategoria=" & idCategoriaPiccole & " Where idTipologia=1 And idCategoria=" & idCategoria & " And Progressivo=" & idMultiMedia
						Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

						If Rit <> "OK" Then
							ScriveLogGlobale(NomeFileLog, Rit)
							Return Rit
						Else
							ScriveLogGlobale(NomeFileLog, "Aggiorno categoria preferito protetto")
							Sql = "Update preferitiprot Set idCategoria=" & idCategoriaPiccole & " Where idTipologia=1 And idCategoria=" & idCategoria & " And Progressivo=" & idMultiMedia
							Rit = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

							If Rit <> "OK" Then
								ScriveLogGlobale(NomeFileLog, Rit)
								Return Rit
							Else
								Ritorno = "*"
							End If
						End If
					End If
				Else
					Ritorno = Rit
				End If
			End If
		End If

		Return Ritorno
	End Function

	'<WebMethod()>
	'Public Function PrendeExif(idCategoria As String) As String
	'	Dim gi As New GestioneImmagini
	'	Dim Db As New clsGestioneDB(TipoDB)
	'	Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
	'	Dim Ritorno As String = ""

	'	If ConnessioneSQL <> "" Then
	'		Dim Rec As Object
	'		Dim Sql As String = ""

	'		Sql = "Select B.Percorso, A.NomeFile From dati A " &
	'			"Left Join categorie B On A.idTipologia = B.idTipologia And A.idCategoria = B.idCategoria " &
	'			"Where A.idTipologia=1 And A.idCategoria=" & idCategoria
	'		Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
	'		If TypeOf (Rec) Is String Then
	'			Return Rec
	'		End If

	'		If Rec.Eof = True Then
	'			Ritorno = "ERROR: Nessun file rilevato"
	'		Else
	'			Dim q As Integer = 0

	'			Do Until Rec.Eof
	'				Dim Nome As String = Rec("Percorso").Value & "/" & Rec("NomeFile").Value

	'				Ritorno &= Nome & ";" & gi.RitornaExif(Nome) & "§"

	'				q += 1
	'				If q = 50 Then
	'					Exit Do
	'				End If

	'				Rec.MoveNext
	'			Loop
	'			Rec.Close
	'		End If
	'	End If

	'	Return Ritorno
	'End Function

	'<WebMethod()>
	'Public Function AggiornaNomiFile() As String
	'	Dim Db As New clsGestioneDB(TipoDB)
	'	Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
	'	Dim Ritorno As String = "*"
	'	Dim gf As New GestioneFilesDirectory

	'	If ConnessioneSQL <> "" Then
	'		Dim Rec As Object
	'		Dim Sql As String = ""

	'		Sql = "Select * From dati where solonome is null or solonome = ''"
	'		Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL)
	'		If TypeOf (Rec) Is String Then
	'			Return Rec
	'		End If

	'		If Rec.Eof = True Then
	'			Ritorno = "ERROR: Nessun file rilevato"
	'		Else
	'			Do Until Rec.Eof
	'				Dim Nome As String = Rec("nomefile").Value
	'				If Nome.Contains("ERROR:") Then
	'					Sql = "Delete From dati " &
	'						"Where idTipologia=" & Rec("idTipologia").Value & " And idCategoria=" & Rec("idCategoria").Value & " And Progressivo=" & Rec("Progressivo").Value

	'					Dim Rit As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

	'					If Rit <> "OK" Then
	'						Ritorno = Rit & "  --- " & Sql
	'						Exit Do
	'					End If
	'				Else
	'					Dim SoloNome As String = gf.TornaNomeFileDaPath(Nome)
	'					If SoloNome = "" Then SoloNome = Nome

	'					Sql = "Update dati Set solonome='" & SoloNome.Replace("'", "''") & "' " &
	'						"Where idTipologia=" & Rec("idTipologia").Value & " And idCategoria=" & Rec("idCategoria").Value & " And Progressivo=" & Rec("Progressivo").Value

	'					Dim Rit As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)

	'					If Rit <> "OK" Then
	'						' Ritorno = Rit & "  --- " & Sql
	'						' Exit Do
	'					End If
	'				End If

	'				Rec.MoveNext
	'			Loop
	'			Rec.Close
	'		End If
	'	End If

	'	Return Ritorno
	'End Function

	<WebMethod()>
	Public Function RitornaInformazioni(idTipologia As String, idCategoria As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		Dim Ritorno As String = ""
		Dim gf As New GestioneFilesDirectory

		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim Sql As String = ""
			Dim Tipol As String = ""
			If idTipologia <> "-1" Then
				Tipol = "B.idTipologia=" & idTipologia & " And "
			End If

			If idTipologia = "1" Or idTipologia = "-1" Then
				Sql = "Select 'Punti Valorizzati', count(*) From informazioniimmagini A " &
					"Left Join dati B On " & Tipol & " A.idCategoria=B.idCategoria And B.Progressivo=A.idMultimedia " &
					"Where (PuntiDiagonale is Not Null Or PuntiDiagonale <> '') And A.idCategoria=" & idCategoria & " And (B.eliminata = 'N' Or B.eliminata = 'n') " &
					"Union All " &
					"Select 'Punti Valorizzati Exif', count(*) From exifimmagini A " &
					"Left Join dati B On " & Tipol & " A.idCategoria=B.idCategoria And B.Progressivo=A.idMultimedia " &
					"Where Descrizione is Not Null  And A.idCategoria=" & idCategoria & " And (B.eliminata = 'N' Or B.eliminata = 'n') " &
					"union ALL " &
					"Select 'Punti Nulli', count(*) From informazioniimmagini A " &
					"Left Join dati B On " & Tipol & " A.idCategoria=B.idCategoria And B.Progressivo=A.idMultimedia " &
					"Where (PuntiDiagonale Is Null Or PuntiDiagonale = '') And A.idCategoria = " & idCategoria & " And (B.eliminata = 'N' Or B.eliminata = 'n') " &
					"Union ALL " &
					"Select 'Tutte le immagini', Count(*) From dati B " &
					"Left Join informazioniimmagini A On " & Tipol & " A.idCategoria=B.idCategoria And B.Progressivo=A.idMultimedia " &
					"Where (B.eliminata = 'N' Or B.eliminata = 'n') And B.idCategoria=" & idCategoria & " "
			End If

			If idTipologia = "-1" Then
				Sql &= "Union All "
			End If

			If idTipologia = "2" Or idTipologia = "-1" Then
				Sql &= "Select 'Video Convertiti', Count(*) From InformazioniVideo A " &
					"Left Join dati B On B.idTipologia=A.idTipologia And A.idCategoria=B.idCategoria And B.Progressivo=A.idMultimedia " &
					"Where instr(JSone, 'h264') > 0 And (B.eliminata = 'N' Or B.eliminata = 'n') And B.idCategoria=" & idCategoria & " " &
					"Union All " &
					"Select 'Video Totali', Count(*) From InformazioniVideo A " &
					"Left Join dati B On B.idTipologia=A.idTipologia And A.idCategoria=B.idCategoria And B.Progressivo=A.idMultimedia " &
					"Where (B.eliminata = 'N' Or B.eliminata = 'n') And B.idCategoria=" & idCategoria & " "
			End If

			Sql &= "Union All " &
				"Select 'Preferiti', (Select Count(*) From preferitiprot where idTipologia=1 And idCategoria = " & idCategoria & ") + (Select Count(*) From preferiti where idTipologia=1 And idCategoria = " & idCategoria & ")" &
				"Union All " &
				"Select 'Percentuale Preferiti', Convert((((Select Count(*) From preferitiprot where idTipologia=1 And idCategoria = " & idCategoria & ") + (Select Count(*) From preferiti where idTipologia=1 And idCategoria = " & idCategoria & ")) / (Select Count(*) From dati where idTipologia=1 And idCategoria = " & idCategoria & ")) * 100, Integer)" &
				"Union All " &
				"Select 'Categorie', Count(*) From categorie"
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL, False)
			If TypeOf (Rec) Is String Then
				Return Rec
			End If

			If Rec.Eof = True Then
				Ritorno = "ERROR: Nessuna informazione rilevata"
			Else
				Do Until Rec.Eof
					Ritorno &= Rec(0).Value & ";" & Rec(1).Value & "§"

					Rec.MoveNext
				Loop
				Rec.Close
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaInformazioniCategorie(idTipologia As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		Dim Ritorno As String = ""
		Dim gf As New GestioneFilesDirectory

		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim Sql As String = ""
			Sql = "Select categoria, count(*), " &
				"(Select count(*) from preferitiprot Aa Left Join dati Bb On Aa.idTipologia=Bb.idTipologia And Aa.idCategoria=Bb.idcategoria And Aa.Progressivo=Bb.progressivo Where Aa.idTipologia=B.idTipologia And Aa.idCategoria=B.idCategoria And (Bb.Eliminata='N' Or Bb.Eliminata='n'))+" &
				"(Select count(*) from preferiti Aa Left Join dati Bb On Aa.idTipologia=Bb.idTipologia And Aa.idCategoria=Bb.idcategoria And Aa.Progressivo=Bb.progressivo Where Aa.idTipologia=B.idTipologia And Aa.idCategoria=B.idCategoria And (Bb.Eliminata='N' Or Bb.Eliminata='n')) As Preferiti, "
			If idTipologia = 1 Or idTipologia = -1 Then
				Sql &= "(select count(*) from informazioniimmagini C Left Join dati Dd On C.idCategoria=Dd.idCategoria And C.idMultimedia=Dd.progressivo Where C.idCategoria=B.idcategoria And PuntiDiagonale Is Not null And (Dd.Eliminata='N' Or Dd.Eliminata='n')) As InformazioniImmagine, "
			Else
				Sql &= "'' As InformazioniImmagine, "
			End If
			If idTipologia = 2 Or idTipologia = -1 Then
				Sql &= "(select count(*) from InformazioniVideo D Left Join dati Ee On D.idTipologia=Ee.idTipologia And D.idCategoria=Ee.idCategoria And D.idMultimedia=Ee.progressivo Where D.idTipologia=B.idTipologia And D.idCategoria=B.idcategoria And (Ee.Eliminata='N' Or Ee.Eliminata='n')) As InformazioniVideo "
			Else
				Sql &= "'' As InformazioniVideo "
			End If
			Sql &= "From categorie A " &
				"Left Join dati B On A.idcategoria=B.idcategoria And A.idtipologia = B.idtipologia " &
				"Where A.idTipologia = " & idTipologia & " " &
				"Group By categoria " &
				"Order By categoria"
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL, False)
			If TypeOf (Rec) Is String Then
				Return Rec
			End If

			If Rec.Eof = True Then
				Ritorno = "ERROR: Nessuna informazione rilevata"
			Else
				Do Until Rec.Eof
					Ritorno &= Rec("Categoria").Value & ";" & Rec(1).Value & ";" & Rec(2).Value & ";" & Rec(3).Value & ";" & Rec(4).Value & "§"

					Rec.MoveNext
				Loop
				Rec.Close
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function SalvaRicerca(Stringa As String, AndOr As String, Pagina As String, DimeThumbs As String, NumeroRighe As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		Dim Ritorno As String = ""

		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim Sql As String = ""

			Sql = "Select * From ricerche Where Testo='" & Stringa.Replace("'", "''") & "'"
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL, False)
			If TypeOf (Rec) Is String Then
				Return Rec
			End If

			If Rec.Eof = True Then
				Sql = "Insert Into ricerche Values ('" & Stringa.Replace("'", "''") & "', " & AndOr & ", " & Pagina & ", " & DimeThumbs & ", " & NumeroRighe & ")"
			Else
				Sql = "Update ricerche set Testo='" & Stringa.Replace("'", "''") & "', AndOr=" & AndOr & ", Pagina=" & Pagina & ", DimensioneThumb=" & DimeThumbs & ", NumeroRighe=" & NumeroRighe & " Where Testo='" & Stringa.Replace("'", "''") & "'"
			End If
			Rec.Close

			Ritorno = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
			If Ritorno = "OK" Then
				Ritorno = "*"
			End If

		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function TornaRicerche() As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		Dim Ritorno As String = ""

		If ConnessioneSQL <> "" Then
			Dim Rec As Object
			Dim Sql As String = ""

			Sql = "Select * From ricerche"
			Rec = Db.LeggeQuery(Server.MapPath("."), Sql, ConnessioneSQL, False)
			If TypeOf (Rec) Is String Then
				Return Rec
			End If
			Do Until Rec.Eof
				Ritorno &= Rec("Testo").Value & ";"
				Ritorno &= Rec("AndOr").Value & ";"
				Ritorno &= Rec("Pagina").Value & ";"
				Ritorno &= Rec("DimensioneThumb").Value & ";"
				Ritorno &= Rec("NumeroRighe").Value & ";"
				Ritorno &= "§"

				Rec.MoveNext
			Loop
			Rec.Close
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaRicerche(Stringa As String) As String
		Dim Db As New clsGestioneDB(TipoDB)
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBase()
		Dim Ritorno As String = ""

		If ConnessioneSQL <> "" Then
			Dim Sql As String = ""

			Sql = "Delete From Where Testo='" & Stringa.Replace("'", "''") & "'"
			Ritorno = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
			If Ritorno = "OK" Then
				Ritorno = "*"
			End If
		End If

		Return Ritorno
	End Function

End Class