Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Security.Cryptography

Module mdlLooVF
	Public FaiLog As Boolean = True
	Public effettuaLog As Boolean = True

	Public ListaImmagini As List(Of String) = New List(Of String)
	Public QuanteImmaginiSfondi As Integer = 0
	Public ContatoreRiletturaImmagini As Integer = 0
	Public StaLeggendoImmagini As Boolean = False
	Public TipoDB As String = "MARIADB"
	Public StringaErrore As String = "ERROR: "
	Public nomeFileLogGenerale As String = ""
	Public listaLog As New List(Of String)
	Public timerLog As New Timers.Timer
	Public timerConv As New Timers.Timer
	Public timerConvI As New Timers.Timer
	Public timerConvP As New Timers.Timer
	Public timerConvE As New Timers.Timer
	Public VecchiaRicerca As String = ""
	Public VecchioQuante As Long
	Public UltimoMultimediaImm As Long
	Public UltimoMultimediaVid As Long
	Public StaEffettuandoConversioneAutomatica As Boolean = False
	Public StaEffettuandoConversioneAutomaticaFinale As Boolean = False
	Public StaEffettuandoConversioneAutomaticaI As Boolean = False
	Public StaEffettuandoConversioneAutomaticaFinaleI As Boolean = False
	Public NumeroFrames As String = ""
	Public NomeFileDaConvertire As String = ""
	Public DaWebGlobale As Boolean = False
	Public idTipologiaGlobale As String = ""
	Public idCategoriaGlobale As String = ""
	Public idMultimediaGlobale As String = ""
	Public processoFFMpeg As Process = New Process()
	Public PathVideoInput As String = ""
	Public PathVideoOutput As String = ""
	Public nomeNuovoGlobale As String = ""
	Public bytesVecchiGlobale As String = ""
	Public bytesNuoviGlobale As String = ""
	Public idCategoriaGlobalePerConversione As String = ""
	Public idCategoriaGlobalePerPuntini As String = ""
	Public idCategoriaGlobalePerExif As String = ""
	Public RefreshPerExif As String = ""
	Public RefreshPerPuntini As String = ""
	Public staCaricandoPuntini As Boolean = False
	Public staCaricandoExif As Boolean = False
	Private Listona As New List(Of String)

	Public Function dataAttuale() As String
		Return Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00")
	End Function

	Public Sub ScriveLogGlobale(Path As String, Cosa As String)
		If FaiLog = True Then
			Try
				Dim gf As New GestioneFilesDirectory

				gf.ApreFileDiTestoPerScrittura(Path)
				gf.ScriveTestoSuFileAperto(dataAttuale() & ": " & Cosa)
				gf.ChiudeFileDiTestoDopoScrittura()
			Catch ex As Exception

			End Try
		End If
	End Sub

	Public Function RitornaEssenziali(Mp As String, db As clsGestioneDB, ConnessioneSql As String, idCategoria As String, Inizio As String,
									  QuanteImmagini As String, TutteLeCategorie As String, ScrittaRitorno As String, Ordinamento As String) As String
		Dim Ritorno As String = ""
		Dim Sql As String = ""
		Dim RicercaCate As String = ""
		Dim Rec As Object
		Dim Rec2 As Object
		Dim Ok As Boolean = True
		Dim NomeFileLog As String = Mp & "/Logs/RicercaEssenziali.txt"

		If TutteLeCategorie = "N" Or TutteLeCategorie = "" Then
			RicercaCate = " And A.idCategoria=" & idCategoria & " "
		End If

		Sql = "Select ***, dimensioni, Width, Height, count(*) From dati As A " &
			"Left Join informazioniimmagini B On A.progressivo = B.idMultimedia And A.idcategoria = B.idCategoria " &
			"Where A.idtipologia = 1 " & RicercaCate & "And Width Is Not Null And height Is Not Null And (A.eliminata = 'N' Or A.eliminata = 'n') " &
			"Group By dimensioni, Width, Height " &
			"Having Count(*) > 1"

		ScriveLogGlobale(NomeFileLog, "-----------------------------------------------")
		'Conteggio
		Dim Sql2 As String = "Select Coalesce(Count(*),0) From (" & Sql.Replace("***, ", "") & ") As Aa"
		ScriveLogGlobale(NomeFileLog, "Query di conteggio: " & Sql2)

		Rec = db.LeggeQuery(Mp, Sql2, ConnessioneSql)
		If TypeOf (Rec) Is String Then
			Ok = False
			Ritorno = Rec & " -> " & Sql2
			ScriveLogGlobale(NomeFileLog, Ritorno)
		End If
		If Ok Then
			Dim QuanteRigheTotali As Integer = Rec(0).Value
			Rec.Close

			ScriveLogGlobale(NomeFileLog, "Righe rilevate: " & QuanteRigheTotali)

			'Lettura righe
			Sql2 = "Select * From (" & Sql.Replace("***, ", "ROW_NUMBER() OVER(Order BY dimensioni, Width, Height) As NumeroRiga, ") & ") As Aa " &
				"Where NumeroRiga > " & (Val(Inizio) - 1) & " And NumeroRiga < " & (Inizio + Val(QuanteImmagini))
			ScriveLogGlobale(NomeFileLog, "Query di dettaglio: " & Sql2)
			Rec = db.LeggeQuery(Mp, Sql2, ConnessioneSql)
			If TypeOf (Rec) Is String Then
				Ok = False
				Ritorno = Rec & " -> " & Sql2
				ScriveLogGlobale(NomeFileLog, Ritorno)
				Return Ritorno
			End If
			If Ok Then
				Do Until Rec.Eof
					ScriveLogGlobale(NomeFileLog, "Ricerca per Dimensione: " & Rec("Dimensioni").Value & " Width: " & Rec("Width").Value & " Height: " & Rec("Height").Value)

					' Ricerca ultima
					Sql = "Select *, Coalesce(d.progressivo, '') As preferito, Coalesce(e.progressivo, '') As preferitoprot From dati A " &
						"Left Join informazioniimmagini b On a.progressivo = b.idMultimedia and a.idcategoria = b.idCategoria " &
						"left join categorie C On C.idtipologia = 1 And A.idCategoria = C.idcategoria " &
						"left join preferiti d On d.idTipologia = 1 And A.idCategoria = d.idCategoria And A.progressivo=d.progressivo " &
						"left join preferitiprot e On e.idTipologia = 1 And A.idCategoria = e.idCategoria And A.progressivo=e.progressivo " &
						"Where dimensioni = " & Rec("Dimensioni").Value & " And width = " & Rec("Width").Value & " And height = " & Rec("Height").Value & " " &
						"Order By " & Ordinamento & ", A.dimensioni, b.Width, b.Height"
					Rec2 = db.LeggeQuery(Mp, Sql, ConnessioneSql)
					If TypeOf (Rec2) Is String Then
						Ok = False
						Ritorno = Rec2 & " -> " & Sql
						ScriveLogGlobale(NomeFileLog, Ritorno)
						Exit Do
					End If
					If Ok Then
						Dim q As Integer = 0

						Do Until Rec2.Eof
							Dim P As String = Rec2("Percorso").Value
							Dim N As String = Rec2("NomeFile").Value

							If Not P.ToUpper.Contains("/VIDEOS/") And Not N.ToUpper.Contains("/VIDEOS/") Then
								Ritorno &= ScrittaRitorno & ";"
								Ritorno &= Rec2("idCategoria").Value & ";"
								Ritorno &= Rec2("idMultimedia").Value & ";"
								Ritorno &= Rec2("Sezione1").Value & ";"
								Ritorno &= Rec2("Punti").Value & ";"
								Ritorno &= Rec2("Width").Value & ";"
								Ritorno &= Rec2("Height").Value & ";"
								Ritorno &= Rec2("DataOra").Value & ";"
								Ritorno &= Rec2("PuntiDiagonale").Value & ";"
								Ritorno &= Rec2("PuntiCornice").Value & ";"
								Ritorno &= N.Replace(";", "---PV---") & ";"
								Ritorno &= P.Replace(";", "---PV---") & ";"
								Ritorno &= Rec2("Preferito").Value & ";"
								Ritorno &= Rec2("PreferitoProt").Value & ";"
								Ritorno &= Rec2("Protetta").Value & ";"
								Ritorno &= Rec2("SoloNome").Value & ";"
								Ritorno &= Rec2("Dimensioni").Value & ";"
								Ritorno &= Rec2("Sezione2").Value & ";"
								Ritorno &= ";"
								Ritorno &= ";"
								Ritorno &= "§"

								q += 1
							End If

							Rec2.MoveNext
						Loop
						Rec2.Close

						ScriveLogGlobale(NomeFileLog, "Immagini rilevate: " & q)
					End If

					Rec.MoveNext
				Loop
				Rec.Close
			End If
		End If
		ScriveLogGlobale(NomeFileLog, "Uscita: " & Ritorno)

		Return Ritorno
	End Function

	Public Function RitornaUguaglianze(Mp As String, db As clsGestioneDB, ConnessioneSql As String, TipoRicerca As String, ScrittaRitorno As String, idCategoria As String, QuanteImmagini As String,
									   Inizio As String, StringaRicerca As String, AndOr As String, TutteLeCategorie As String, Caratteri As String, Ordinamento As String) As String
		Dim Ritorno As String = ""
		Dim Ok As Boolean = True
		Dim Rec As Object
		Dim Rec2 As Object
		Dim Sql As String = ""
		Dim NomeFileLog As String = Mp & "/Logs/RicercaUguali.txt"

		ScriveLogGlobale(NomeFileLog, "-----------------------------------------")

		Dim RicercaCategoria As String = ""

		If TutteLeCategorie = "N" Or TutteLeCategorie = "" Then
			RicercaCategoria = " And A.idCategoria=" & idCategoria & " "
			ScriveLogGlobale(NomeFileLog, "Ricerca per categoria " & idCategoria)
		Else
			ScriveLogGlobale(NomeFileLog, "Ricerca per tutte le categorie")
		End If

		If TipoRicerca = "STRINGA" Then
			ScriveLogGlobale(NomeFileLog, "Ricerca di tipo stringa: " & TipoRicerca)

			Dim StringonaRicerca As String = ""
			Dim Filtro As String = ""

			If AndOr = 1 Then
				Filtro = "And"
			Else
				Filtro = "Or "
			End If

			Dim NomeCampo As String = ""

			Select Case Ordinamento
				Case "SoloNome"
					NomeCampo = "solonome"
				Case "Percorso"
					NomeCampo = "nomefile"
				Case "NomeFile"
					NomeCampo = "solonome"
			End Select

			If StringaRicerca.IndexOf(";") > 0 Then
				Dim StringheRicerca() As String = StringaRicerca.Split(";")

				For Each s As String In StringheRicerca
					If s <> "" Then
						StringonaRicerca &= "B." & NomeCampo & " Like '%" & s.Replace("'", "''") & "%' " & Filtro & " "
						'StringonaRicerca &= "B.Percorso Like '%" & s.Replace("'", "''") & "%' " & Filtro & " "
					End If
				Next
				If StringonaRicerca <> "" Then
					StringonaRicerca = " And (" & StringonaRicerca.Substring(0, StringonaRicerca.Length - 5) & ") "
				End If
			Else
				If StringaRicerca = "" Then
					StringonaRicerca = ""
				Else
					StringonaRicerca = "And (B." & NomeCampo & " Like '%" & StringaRicerca & "%' )" ' Or C.Percorso Like '%" & StringaRicerca & "%'
				End If
			End If

			Sql = "Select Coalesce(Count(*),0) FROM informazioniimmagini A " &
				"Left Join dati B On B.idtipologia = 1 And A.idCategoria = B.idcategoria And A.idMultimedia = B.progressivo " &
				"Left Join categorie C On A.idCategoria = C.idcategoria And C.idtipologia = B.idTipologia " &
				"Where (B.eliminata='N' Or B.eliminata='n') " & RicercaCategoria & " And Instr(C.Percorso, '/Videos/') = 0 " &
				" " & StringonaRicerca & " "
			ScriveLogGlobale(NomeFileLog, "Query di conteggio: " & Sql)

			'Return Sql
			Rec = db.LeggeQuery(Mp, Sql, ConnessioneSql)
			If TypeOf (Rec) Is String Then
				Ok = False
				Ritorno = Rec & " -> " & Sql
				ScriveLogGlobale(NomeFileLog, Ritorno)
			End If
			If Ok Then
				Dim QuanteRigheTotali As Integer = Rec(0).Value
				Rec.Close

				ScriveLogGlobale(NomeFileLog, "Righe rilevate: " & QuanteRigheTotali)

				'If QuanteRigheTotali > 1000 Then
				'	Return "ERROR: Troppe righe ritornate (" & QuanteRigheTotali & ")"
				'End If

				Sql = "Select * From (" &
					"Select ROW_NUMBER() OVER(Order BY B.dimensioni, A.Width, A.Height, B.solonome) As NumeroRiga, A.*, B.nomefile, C.percorso, C.protetta, Coalesce(d.progressivo, '') As preferito, Coalesce(e.progressivo, '') As preferitoprot, B.solonome, B.Dimensioni " &
					"From informazioniimmagini A " &
					"left join dati B On B.idtipologia = 1 And A.idCategoria = B.idCategoria And A.idMultimedia=B.progressivo " &
					"left join categorie C On C.idtipologia = 1 And A.idCategoria = C.idcategoria " &
					"left join preferiti d On d.idTipologia = 1 And A.idCategoria = d.idCategoria And A.idMultimedia=d.progressivo " &
					"left join preferitiprot e On e.idTipologia = 1 And A.idCategoria = e.idCategoria And A.idMultimedia=e.progressivo " &
					"Where (A.Eliminata='N' Or A.Eliminata='n') And (B.Eliminata='N' Or B.Eliminata='n') " & RicercaCategoria & " " &
					" " & StringonaRicerca & " " &
					") As A Where NumeroRiga > " & (Val(Inizio) - 1) & " And NumeroRiga < " & (Inizio + Val(QuanteImmagini)) & " " &
					"Order By " & Ordinamento & ", A.dimensioni, A.Width, A.Height"
				'Return Sql
				ScriveLogGlobale(NomeFileLog, "Query di ritorno righe: " & Sql)

				Rec = db.LeggeQuery(Mp, Sql, ConnessioneSql)
				If TypeOf (Rec) Is String Then
					Ok = False
					Ritorno = Rec & " -> " & Sql
					ScriveLogGlobale(NomeFileLog, Ritorno)
				End If
				If Ok Then
					Dim q As Integer = 0

					Do Until Rec.Eof
						Dim P As String = Rec("Percorso").Value
						Dim N As String = Rec("NomeFile").Value

						If Not P.ToUpper.Contains("/VIDEOS/") And Not N.ToUpper.Contains("/VIDEOS/") Then
							Dim Ok2 As Boolean = True
							If ScrittaRitorno = "NOME UGUALE" Then
								If Len(Rec("SoloNome").Value) < Caratteri Then
									Ok2 = False
								End If
							End If
							If Ok2 Then
								Ritorno &= ScrittaRitorno & ";"
								Ritorno &= Rec("idCategoria").Value & ";"
								Ritorno &= Rec("idMultimedia").Value & ";"
								Ritorno &= Rec("Sezione1").Value & ";"
								Ritorno &= Rec("Punti").Value & ";"
								Ritorno &= Rec("Width").Value & ";"
								Ritorno &= Rec("Height").Value & ";"
								Ritorno &= Rec("DataOra").Value & ";"
								Ritorno &= Rec("PuntiDiagonale").Value & ";"
								Ritorno &= Rec("PuntiCornice").Value & ";"
								Ritorno &= N.Replace(";", "---PV---") & ";"
								Ritorno &= P.Replace(";", "---PV---") & ";"
								Ritorno &= Rec("Preferito").Value & ";"
								Ritorno &= Rec("PreferitoProt").Value & ";"
								Ritorno &= Rec("Protetta").Value & ";"
								Ritorno &= Rec("SoloNome").Value & ";"
								Ritorno &= Rec("Dimensioni").Value & ";"
								Ritorno &= Rec("Sezione2").Value & ";"
								Ritorno &= ";"
								Ritorno &= ";"
								Ritorno &= "§"

								q += 1
							End If
						End If

						Rec.MoveNext
					Loop
					Rec.Close

					Ritorno &= "*" & QuanteRigheTotali

					ScriveLogGlobale(NomeFileLog, "Righe ritornate: " & q)
				End If
			End If

			Return Ritorno
		End If

		If TipoRicerca = "Exif" Then
			ScriveLogGlobale(NomeFileLog, "Ricerca di tipo EXIF: " & TipoRicerca)

			Dim RicercaCate As String = ""
			Dim Campo As String = ""

			If ScrittaRitorno = "EXIF DESCR." Then
				Campo = "Descrizione"
			Else
				Campo = "Commento"
			End If

			If TutteLeCategorie = "N" Or TutteLeCategorie = "" Then
				RicercaCate = "idCategoria=" & idCategoria & " And "
			End If

			Sql = "Select Count(*) From (SELECT " & Campo & " FROM exifimmagini " &
				"Where " & RicercaCate & " " & Campo & " <> '' " &
				"group by " & Campo & " having count(*) > 1) As A"
			ScriveLogGlobale(NomeFileLog, "Query di conteggio: " & Sql)

			'Return Sql
			Rec = db.LeggeQuery(Mp, Sql, ConnessioneSql)
			If TypeOf (Rec) Is String Then
				Ok = False
				Ritorno = Rec & " -> " & Sql
				ScriveLogGlobale(NomeFileLog, Ritorno)
			End If
			If Ok Then
				Dim QuanteRigheTotali As Integer = Rec(0).Value
				Rec.Close

				ScriveLogGlobale(NomeFileLog, "Righe rilevate: " & QuanteRigheTotali)

				'If QuanteRigheTotali > 1000 Then
				'	Return "ERROR: Troppe righe ritornate (" & QuanteRigheTotali & ")"
				'End If
				Sql = "Select * From (" &
					"SELECT ROW_NUMBER() OVER(Order BY " & Campo & ") As NumeroRiga, " & Campo & " FROM exifimmagini " &
					"Where " & RicercaCate & " " & Campo & " <> '' " &
					"group by " & Campo & " " &
					"having count(*) > 1 " &
					") As A Where NumeroRiga > " & (Val(Inizio) - 1) & " And NumeroRiga < " & (Inizio + Val(QuanteImmagini))
				'Return Sql
				ScriveLogGlobale(NomeFileLog, "Query di ritorno righe: " & Sql)

				Rec = db.LeggeQuery(Mp, Sql, ConnessioneSql)
				If TypeOf (Rec) Is String Then
					Ok = False
					Ritorno = Rec & " -> " & Sql
					ScriveLogGlobale(NomeFileLog, Ritorno)
				End If
				If Ok Then
					Try
						Dim q As Integer = 0

						If TutteLeCategorie = "N" Or TutteLeCategorie = "" Then
							RicercaCate = "And A.idCategoria=" & idCategoria & " "
						Else
							RicercaCate = ""
						End If

						Do Until Rec.Eof
							ScriveLogGlobale(NomeFileLog, "Query di ritorno dettaglio su descrizione: " & Rec(Campo).Value)

							Sql = "SELECT *, Coalesce(d.progressivo, '') As preferito, Coalesce(e.progressivo, '') As preferitoprot FROM exifimmagini A " &
								"left join dati B On B.idtipologia = 1 And A.idCategoria = B.idCategoria And A.idMultimedia=B.progressivo " &
								"left join categorie c On c.idtipologia = 1 And A.idCategoria = c.idcategoria " &
								"left join preferiti d On d.idTipologia = 1 And A.idCategoria = d.idCategoria And A.idMultimedia=d.progressivo " &
								"left join preferitiprot e On e.idTipologia = 1 And A.idCategoria = e.idCategoria And A.idMultimedia=e.progressivo " &
								"left join informazioniimmagini f On A.idCategoria = f.idCategoria And A.idMultimedia=f.idMultimedia " &
								"where (" & Campo & " = '" & Rec(Campo).Value.replace("'", "''") & "') And (B.Eliminata='N' Or B.Eliminata='n') " & RicercaCate & " " &
								"Order By " & Ordinamento & ", B.dimensioni, f.Width, f.Height"
							Rec2 = db.LeggeQuery(Mp, Sql, ConnessioneSql)
							If TypeOf (Rec2) Is String Then
								Ritorno = Rec2 & " -> " & Sql
								ScriveLogGlobale(NomeFileLog, Ritorno)
							Else
								Do Until Rec2.Eof
									Dim P As String = Rec2("Percorso").Value
									Dim N As String = Rec2("NomeFile").Value

									If Not P.ToUpper.Contains("/VIDEOS/") And Not N.ToUpper.Contains("/VIDEOS/") Then
										If Not Ritorno.Contains(P) Or Not Ritorno.Contains(N) Then
											ScriveLogGlobale(NomeFileLog, "Ritorno dettaglio per " & Campo & ": " & P & " / " & N)

											Ritorno &= ScrittaRitorno & ";"
											Ritorno &= Rec2("idCategoria").Value & ";"
											Ritorno &= Rec2("idMultimedia").Value & ";"
											Ritorno &= Rec2("Sezione1").Value & ";"
											Ritorno &= Rec2("Punti").Value & ";"
											Ritorno &= Rec2("Width").Value & ";"
											Ritorno &= Rec2("Height").Value & ";"
											Ritorno &= Rec2("DataOra").Value & ";"
											Ritorno &= Rec2("PuntiDiagonale").Value & ";"
											Ritorno &= Rec2("PuntiCornice").Value & ";"
											Ritorno &= N.Replace(";", "---PV---") & ";"
											Ritorno &= P.Replace(";", "---PV---") & ";"
											Ritorno &= Rec2("Preferito").Value & ";"
											Ritorno &= Rec2("PreferitoProt").Value & ";"
											Ritorno &= Rec2("Protetta").Value & ";"
											Ritorno &= Rec2("SoloNome").Value & ";"
											Ritorno &= Rec2("Dimensioni").Value & ";"
											Ritorno &= Rec2("Sezione2").Value & ";"
											Ritorno &= Rec2("Descrizione").Value & ";"
											Ritorno &= Rec2("Commento").Value & ";"
											Ritorno &= "§"

											q += 1
										End If
									End If

									Rec2.MoveNext
								Loop
								Rec2.Close
							End If

							Rec.MoveNext
						Loop
						Rec.Close

						Ritorno &= "*" & QuanteRigheTotali

						ScriveLogGlobale(NomeFileLog, "Righe ritornate: " & q)
					Catch ex As Exception
						ScriveLogGlobale(NomeFileLog, "ERROR: " & ex.Message)
					End Try
				End If
			End If

			Return Ritorno
		End If

		ScriveLogGlobale(NomeFileLog, "Ricerca di altro tipo")

		Sql = "Select Coalesce(Count(*),0) From ( " &
			"Select " & TipoRicerca & " As TipoRicerca FROM informazioniimmagini A " &
			"Left Join dati B On B.idtipologia = 1 And A.idCategoria = B.idcategoria And A.idMultimedia = B.progressivo " &
			"Where (B.eliminata='N' Or B.eliminata='n') " & RicercaCategoria & " And " & TipoRicerca & " Is Not Null And " & TipoRicerca & " <> '' " &
			"Group By " & TipoRicerca & " " &
			"Having Count(*) > 1 " &
			") As A"
		'"Left Join categorie C On A.idCategoria = C.idcategoria And C.idtipologia = B.idTipologia " &
		'"Where (B.eliminata='N' Or B.eliminata='n') And Instr(C.Percorso, '/Videos/') = 0 " &
		'Return Sql
		ScriveLogGlobale(NomeFileLog, "Query di conteggio righe: " & Sql)
		Rec = db.LeggeQuery(Mp, Sql, ConnessioneSql)
		If TypeOf (Rec) Is String Then
			Ok = False
			ScriveLogGlobale(NomeFileLog, "ERROR: " & Rec)
		End If
		If Ok Then
			Dim QuanteRigheTotali As Integer = Rec(0).Value
			Rec.Close

			ScriveLogGlobale(NomeFileLog, "Righe ritornate: " & QuanteRigheTotali)
			If QuanteRigheTotali > 0 Then
				'Sql = "Select TipoRicerca From ( " &
				'	"Select ROW_NUMBER() OVER(Order BY " & TipoRicerca & ") As NumeroRiga, " & TipoRicerca & " As TipoRicerca, count(*) FROM informazioniimmagini " &
				'	"where eliminata='N' or eliminata='n' group by " & TipoRicerca & " having count(*) > 1 " &
				'	") As A Where NumeroRiga > " & (Val(Inizio) - 1) & " And NumeroRiga < " & (Inizio + Val(QuanteImmagini))
				Sql = "Select Coalesce(TipoRicerca, '***') As TipoRicerca From ( " &
					"Select ROW_NUMBER() OVER(Order BY B.dimensioni, A.Width, A.Height, B.solonome) As NumeroRiga, " & TipoRicerca & " As TipoRicerca, count(*) " &
					"FROM informazioniimmagini A " &
					"Left Join dati B On B.idtipologia = 1 And A.idCategoria = B.idcategoria And A.idMultimedia = B.progressivo " &
					"Where (B.eliminata='N' Or B.eliminata='n') " & RicercaCategoria & " And " & TipoRicerca & " Is Not Null And " & TipoRicerca & " <> '' " &
					"Group By " & TipoRicerca & " " &
					"Having Count(*) > 1 " &
					") As A Where NumeroRiga > " & (Val(Inizio) - 1) & " And NumeroRiga < " & (Inizio + Val(QuanteImmagini))

				ScriveLogGlobale(NomeFileLog, "Query 2: " & Sql)
				'"Left Join categorie C On A.idCategoria = C.idcategoria And C.idtipologia = B.idTipologia " &
				'"Where B.eliminata='N' Or B.eliminata='n' And Instr(C.Percorso, '/Videos/') = 0 " &
				Rec = db.LeggeQuery(Mp, Sql, ConnessioneSql)
				If TypeOf (Rec) Is String Then
					Ok = False
					ScriveLogGlobale(NomeFileLog, "ERROR: " & Rec)
				End If
				If Ok Then
					Dim Lista As New List(Of String)

					Do Until Rec.Eof
						If (Rec("TipoRicerca").Value <> "***") Then
							Lista.Add(Rec("TipoRicerca").Value)
						End If

						Rec.MoveNext
					Loop
					Rec.Close

					ScriveLogGlobale(NomeFileLog, "Liste rilevate: " & Lista.Count)

					For Each l As String In Lista
						If l <> "***" Then
							ScriveLogGlobale(NomeFileLog, " ")

							Ok = True

							Sql = "Select A.*, B.nomefile, c.percorso, c.protetta, Coalesce(d.progressivo, '') As preferito, Coalesce(e.progressivo, '') As preferitoprot, B.solonome, B.Dimensioni " &
								"From informazioniimmagini A " &
								"left join dati b On B.idtipologia = 1 And A.idCategoria = B.idCategoria And A.idMultimedia=B.progressivo " &
								"left join categorie c On c.idtipologia = 1 And A.idCategoria = c.idcategoria " &
								"left join preferiti d On d.idTipologia = 1 And A.idCategoria = d.idCategoria And A.idMultimedia=d.progressivo " &
								"left join preferitiprot e On e.idTipologia = 1 And A.idCategoria = e.idCategoria And A.idMultimedia=e.progressivo " &
								"Where " & TipoRicerca & " = '" & l & "' And (A.Eliminata='N' Or A.Eliminata='n') And (B.Eliminata='N' Or B.Eliminata='n') " &
								"And length(substr(b.solonome,1,instr(b.solonome,'.')-1)) > " & (Val(Caratteri) - 1) & " " & RicercaCategoria & " " &
								"Order By " & Ordinamento & ", B.dimensioni, A.Width, A.Height"
							ScriveLogGlobale(NomeFileLog, "Ricerco per lista: " & Sql)
							Rec = db.LeggeQuery(Mp, Sql, ConnessioneSql)
							If TypeOf (Rec) Is String Then
								Ok = False
								ScriveLogGlobale(NomeFileLog, "ERROR: " & Rec)
							End If
							If Ok Then
								Dim q As Integer = 0
								Dim Ritorno2 As String = ""

								Do Until Rec.Eof
									Dim P As String = Rec("Percorso").Value
									Dim N As String = Rec("NomeFile").Value

									If Not P.ToUpper.Contains("/VIDEOS/") And Not N.ToUpper.Contains("/VIDEOS/") And Not N.ToUpper.Contains("/PICCOLE/") Then
										'Dim Ok2 As Boolean = True
										'If ScrittaRitorno = "NOME UGUALE" Then
										'	If Len(Rec("SoloNome").Value) < Caratteri Then
										'		Ok2 = False
										'	End If
										'End If
										'If Ok2 Then
										Ritorno2 &= ScrittaRitorno & ";"
										Ritorno2 &= Rec("idCategoria").Value & ";"
										Ritorno2 &= Rec("idMultimedia").Value & ";"
										Ritorno2 &= Rec("Sezione1").Value & ";"
										Ritorno2 &= Rec("Punti").Value & ";"
										Ritorno2 &= Rec("Width").Value & ";"
										Ritorno2 &= Rec("Height").Value & ";"
										Ritorno2 &= Rec("DataOra").Value & ";"
										Ritorno2 &= Rec("PuntiDiagonale").Value & ";"
										Ritorno2 &= Rec("PuntiCornice").Value & ";"
										Ritorno2 &= N.Replace(";", "---PV---") & ";"
										Ritorno2 &= P.Replace(";", "---PV---") & ";"
										Ritorno2 &= Rec("Preferito").Value & ";"
										Ritorno2 &= Rec("PreferitoProt").Value & ";"
										Ritorno2 &= Rec("Protetta").Value & ";"
										Ritorno2 &= Rec("SoloNome").Value & ";"
										Ritorno2 &= Rec("Dimensioni").Value & ";"
										Ritorno2 &= ";"
										Ritorno2 &= ";"
										Ritorno2 &= ";"
										Ritorno2 &= "§"
										'End If

										q += 1
									End If

									Rec.MoveNext
								Loop
								Rec.Close

								ScriveLogGlobale(NomeFileLog, "Righe rilevate per lista: " & l & " -> " & q)

								If q > 1 Then
									Ritorno &= Ritorno2
								End If
							End If
						End If
					Next

					Ritorno = Ritorno & "*" & QuanteRigheTotali
				End If
			Else
				Ritorno = "ERROR: Nessuna uguaglianza rilevata"
			End If
		End If

		Return Ritorno
	End Function

	Public Function CercaStringa(Mp As String, Db As clsGestioneDB, ConnessioneSql As String, NomeFileLog As String, idCategoria As String, Stringa As String,
								 Inizio As String, QuanteImmagini As String, MettiPunto As Boolean) As String
		Dim Sql As String = ""
		Dim Rec As Object
		Dim Rec2 As Object
		Dim Ritorno As String = ""
		Dim Ok As Boolean = True
		Dim gf As New GestioneFilesDirectory
		Dim Altro As String = ""
		Dim Rit As String = ""
		Dim Punto As String = ""

		If MettiPunto Then
			Punto = "."
		End If

		If idCategoria <> "-1" Then
			Altro = " and idCategoria=" & idCategoria
		End If

		Listona = New List(Of String)

		ScriveLogGlobale(NomeFileLog, "")
		ScriveLogGlobale(NomeFileLog, "-----------------------------------------------")
		ScriveLogGlobale(NomeFileLog, "Ricerca per " & Stringa & " con categoria " & idCategoria)
		Sql = "Select * From (" &
			"Select ROW_NUMBER() OVER(Order BY solonome) As NumeroRiga, solonome, count(*) From dati where solonome Like '%" & Stringa & Punto & "%'" & Altro & " group by solonome having count(*)>1) As A " &
			"Where NumeroRiga > " & (Val(Inizio) - 1) & " And NumeroRiga < " & (Inizio + Val(QuanteImmagini))

		ScriveLogGlobale(NomeFileLog, Sql)
		Rec = Db.LeggeQuery(Mp, Sql, ConnessioneSql)
		If TypeOf (Rec) Is String Then
			Ok = False
			Ritorno = Rec & " -> " & Sql
			ScriveLogGlobale(NomeFileLog, Ritorno)
		End If
		If Ok Then
			ScriveLogGlobale(NomeFileLog, "Inizio ciclo: " & Rec.Eof)
			Dim qq As Integer = 0
			Dim StringaRicerca As String = ""
			Do Until Rec.Eof
				ScriveLogGlobale(NomeFileLog, "Nome file: " & Rec("solonome").Value)
				Dim NF As String = Rec("solonome").Value
				Dim Este As String = gf.TornaEstensioneFileDaPath(NF)
				ScriveLogGlobale(NomeFileLog, "Estensione: " & Este)
				Dim NomeFile As String = NF.Replace(Este, "")
				NomeFile = NomeFile.Replace(Stringa, "").ToUpper
				' ScriveLogGlobale(NomeFileLog, "Ricerco '" & NomeFile & "'")

				If Not StringaRicerca.Contains(NomeFile) And NomeFile.Length > 3 Then
					StringaRicerca &= "upper(a.NomeFile) Like '%" & NomeFile & "%' Or "
				End If

				'qq += 1
				'If qq > 50 Then
				'	qq = 0
				Rit = EffettuaRicercaStringa(Mp, NomeFileLog, idCategoria, StringaRicerca, Db, ConnessioneSql, Stringa)
				'End If

				Rec.MoveNext
			Loop
			Rec.Close

			'Rit = EffettuaRicercaStringa(Mp, NomeFileLog, idCategoria, StringaRicerca, Db, ConnessioneSql, Stringa)
		End If

		'For i As Integer = 0 To Listona.Count - 1
		'	For k As Integer = i + 1 To Listona.Count - 1
		'		If Listona.Item(i) > Listona.Item(k) Then
		'			Dim l As String = Listona.Item(i)
		'			Listona.Item(i) = Listona.Item(k)
		'			Listona.Item(k) = l
		'		End If
		'	Next
		'Next

		Ritorno = ""
		For Each l As String In Listona
			If l <> "" Then
				Dim ll As String = Mid(l, l.IndexOf("#") + 2, l.Length)

				Ritorno &= ll
			End If
		Next

		Return Ritorno
	End Function

	Private Function EffettuaRicercaStringa(Mp As String, NomeFileLog As String, idCategoria As String, StringaRicerca As String, Db As clsGestioneDB, ConnessioneSql As String, Stringa As String) As String
		Dim Altro As String = ""
		Dim Rec2 As Object
		Dim Rec3 As Object
		Dim Ok As Boolean = True
		Dim Ritorno As String = ""
		Dim gf As New GestioneFilesDirectory

		If StringaRicerca <> "" Then
			StringaRicerca = "(" & Mid(StringaRicerca, 1, StringaRicerca.Length - 4) & ")"
			ScriveLogGlobale(NomeFileLog, "Stringa ricerca: " & StringaRicerca)

			If idCategoria <> "-1" Then
				Altro = " and a.idCategoria=" & idCategoria
			End If

			Dim Fatti As New List(Of String)

			Dim Sql As String = "" &
				"Select " & 'ROW_NUMBER() OVER(Order BY a.dimensioni, b.Width, b.Height, a.solonome) As NumeroRiga, 
				"a.idCategoria, a.nomefile, a.solonome, a.dimensioni, a.progressivo, " &
				"b.idMultimedia, b.width, b.height, b.punti, b.Sezione1, b.Sezione2, b.PuntiDiagonale, b.PuntiCornice, b.dataora, " &
				"c.percorso, c.protetta, " &
				"coalesce(d.idTipologia, '') as Preferito, coalesce(e.idTipologia, '') as PreferitoProt From dati a " &
				"left join informazioniimmagini b On a.idcategoria = b.idCategoria and a.progressivo = b.idMultimedia " &
				"left join categorie c on a.idtipologia = c.idtipologia and a.idcategoria = c.idcategoria " &
				"left join preferiti d on a.idtipologia = d.idtipologia and a.idcategoria = d.idcategoria and a.progressivo = d.progressivo " &
				"left join preferitiprot e on a.idtipologia = e.idtipologia and a.idcategoria = e.idcategoria and a.progressivo = e.progressivo " &
				"where " & StringaRicerca & " " & Altro & " and (a.eliminata = 'N' or a.eliminata = 'n') and nomefile not like '%PICCOLE/%' Order by solonome" ' order by a.solonome, a.percorso, b.dataora, b.width, b.height " ' &
			' ") As a Where NumeroRiga > " & (Val(Inizio) - 1) & " And NumeroRiga < " & (Inizio + Val(QuanteImmagini))
			StringaRicerca = ""
			ScriveLogGlobale(NomeFileLog, Sql)
			Rec2 = Db.LeggeQuery(Mp, Sql, ConnessioneSql)
			Ok = True
			If TypeOf (Rec2) Is String Then
				Ok = False
				Ritorno = Rec2 & " -> " & Sql
				ScriveLogGlobale(NomeFileLog, Ritorno)
			End If
			If Ok Then
				'Dim q As Integer = 0
				Do Until Rec2.eof
					Dim NF As String = Rec2("solonome").Value
					Dim Este As String = gf.TornaEstensioneFileDaPath(NF)
					Dim NomeFile As String = NF.Replace(Este, "")
					NomeFile = NomeFile.Replace(Stringa, "").ToUpper
					Dim Ok2 As Boolean = True

					For Each f As String In Fatti
						If f.ToUpper = NomeFile.ToUpper Or f.ToUpper.Contains(NomeFile.ToUpper) Then
							Ok2 = False
						End If
					Next

					If Ok2 Then
						Fatti.Add(NomeFile)
						'Dim Quanti As Integer = 0
						'Sql = "Select Count(*) From dati where idCategoria=" & idCategoria & " And (eliminata='N' or eliminata='n') and solonome like '%" & NomeFile & "%'"
						'ScriveLogGlobale(NomeFileLog, "Conto '" & NomeFile & "': " & Sql)
						'Rec3 = Db.LeggeQuery(Mp, Sql, ConnessioneSql)
						'Ok = True
						'If TypeOf (Rec3) Is String Then
						'	Ok = False
						'	Ritorno = Rec3 & " -> " & Sql
						'	ScriveLogGlobale(NomeFileLog, Ritorno)
						'	Exit Do
						'End If
						'If Ok Then
						'	Quanti = Rec3(0).Value
						'End If
						'Rec3.Close
						'ScriveLogGlobale(NomeFileLog, "Rilevati: " & Quanti)

						'If Quanti > 1 Then
						Sql = "Select *, " &
							"coalesce(d.idTipologia, '') as Preferito, coalesce(e.idTipologia, '') as PreferitoProt " &
							"From dati a " &
							"left join informazioniimmagini b On a.idcategoria = b.idCategoria And a.progressivo = b.idMultimedia " &
							"left join categorie c on a.idtipologia = c.idtipologia And a.idcategoria = c.idcategoria " &
							"left join preferiti d on a.idtipologia = d.idtipologia And a.idcategoria = d.idcategoria And a.progressivo = d.progressivo " &
							"left join preferitiprot e on a.idtipologia = e.idtipologia And a.idcategoria = e.idcategoria And a.progressivo = e.progressivo " &
							"where a.idCategoria=" & idCategoria & " And (a.eliminata='N' or a.eliminata='n') and upper(a.solonome) like '" & NomeFile & "%'"
						ScriveLogGlobale(NomeFileLog, "Rilevo '" & NomeFile & "': " & Sql)
						Rec3 = Db.LeggeQuery(Mp, Sql, ConnessioneSql)
						Ok = True
						If TypeOf (Rec3) Is String Then
							Ok = False
							Ritorno = Rec3 & " -> " & Sql
							ScriveLogGlobale(NomeFileLog, Ritorno)
							Exit Do
						End If
						If Ok Then
							Dim Ritorno2 As String = ""
							Dim q As Integer = 0

							Do Until Rec3.eof
								Dim P As String = Rec3("Percorso").Value
								Dim N As String = Rec3("NomeFile").Value

								Dim Ok3 As Boolean = True

								For Each l As String In Listona
									If l.ToUpper.Contains(N.Replace(";", "---PV---").ToUpper) Then
										Ok3 = False
										Exit For
									End If
								Next

								If Ok3 Then
									Dim NF2 As String = Rec3("SoloNome").Value
									Dim Este2 As String = gf.TornaEstensioneFileDaPath(NF2)
									Dim NomeFile2 As String = NF2.Replace(Este2, "")
									NomeFile2 = NomeFile2.Replace(Stringa, "").ToUpper

									Ritorno2 &= NomeFile2.Replace(";", "---PV---") & "#"
									Ritorno2 &= Stringa & ";"
									Ritorno2 &= Rec3("idCategoria").Value & ";"
									Ritorno2 &= Rec3("Progressivo").Value & ";"
									Ritorno2 &= Rec3("Sezione1").Value & ";"
									Ritorno2 &= Rec3("Punti").Value & ";"
									Ritorno2 &= Rec3("Width").Value & ";"
									Ritorno2 &= Rec3("Height").Value & ";"
									Ritorno2 &= Rec3("DataOra").Value & ";"
									Ritorno2 &= Rec3("PuntiDiagonale").Value & ";"
									Ritorno2 &= Rec3("PuntiCornice").Value & ";"
									Ritorno2 &= N.Replace(";", "---PV---") & ";"
									Ritorno2 &= P.Replace(";", "---PV---") & ";"
									Ritorno2 &= Rec3("Preferito").Value & ";"
									Ritorno2 &= Rec3("PreferitoProt").Value & ";"
									Ritorno2 &= Rec3("Protetta").Value & ";"
									Ritorno2 &= Rec3("SoloNome").Value & ";"
									Ritorno2 &= Rec3("Dimensioni").Value & ";"
									Ritorno2 &= Rec3("Sezione2").Value & ";"
									Ritorno2 &= NomeFile '  NomeFile2.Replace(";", "---PV---") & ";"
									Ritorno2 &= "§"

									q += 1
									ScriveLogGlobale(NomeFileLog, "Aggiungo: " & NomeFile & " -> " & q)

								End If

								Rec3.MoveNext
							Loop
							Rec3.Close

							If q > 1 Then
								ScriveLogGlobale(NomeFileLog, "Aggiungo alla lista")
								Listona.Add(Ritorno2)
							End If
						End If

						'Ok = True
						'For Each l As String In Listona
						'	If l.ToUpper.Contains(N.ToUpper) Then
						'		Ok = False
						'		Exit For
						'	End If
						'Next
						'If Ok Then
						'End If
					End If

					Rec2.MoveNext
					'q += 1
				Loop

				Rec2.Close
				'If q > 1 Then
				'End If
			End If

			'qq += 1
			'If qq > 50 Then
			'	Exit Do
			'End If
		End If

		Return Ritorno
	End Function
End Module
