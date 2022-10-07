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
	Public RefreshPerPuntini As String = ""
	Public staCaricandoPuntini As Boolean = False

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

	Public Function RitornaUguaglianze(Mp As String, db As clsGestioneDB, ConnessioneSql As String, TipoRicerca As String, ScrittaRitorno As String, idCategoria As String, QuanteImmagini As String,
									   Inizio As String, StringaRicerca As String, AndOr As String, TutteLeCategorie As String) As String
		Dim Ritorno As String = ""
		Dim Ok As Boolean = True
		Dim Rec As Object
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

			If StringaRicerca.IndexOf(";") > 0 Then
				Dim StringheRicerca() As String = StringaRicerca.Split(";")

				For Each s As String In StringheRicerca
					If s <> "" Then
						StringonaRicerca &= "(B.NomeFile Like '%" & s.Replace("'", "''") & "%' Or C.Percorso Like '%" & s.Replace("'", "''") & "%') " & Filtro & " "
					End If
				Next
				If StringonaRicerca <> "" Then
					StringonaRicerca = " And (" & StringonaRicerca.Substring(0, StringonaRicerca.Length - 5) & ") "
				End If
			Else
				If StringaRicerca = "" Then
					StringonaRicerca = ""
				Else
					StringonaRicerca = "And (B.NomeFile Like '%" & StringaRicerca & "%' Or C.Percorso Like '%" & StringaRicerca & "%')"
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
					"Order By percorso, nomefile, A.dimensioni, A.Width, A.Height"
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
							Ritorno &= ScrittaRitorno & ";"
							Ritorno &= Rec("idCategoria").Value & ";"
							Ritorno &= Rec("idMultimedia").Value & ";"
							Ritorno &= Rec("Hash").Value & ";"
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
							Ritorno &= Rec("HashColore").Value & ";"
							Ritorno &= "§"

							q += 1
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

		ScriveLogGlobale(NomeFileLog, "Ricerca di altro tipo")

		Sql = "Select Coalesce(Count(*),0) From ( " &
			"Select " & TipoRicerca & " As TipoRicerca FROM informazioniimmagini A " &
			"Left Join dati B On B.idtipologia = 1 And A.idCategoria = B.idcategoria And A.idMultimedia = B.progressivo " &
			"Where (B.eliminata='N' Or B.eliminata='n') " & RicercaCategoria & " " &
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
			'Sql = "Select TipoRicerca From ( " &
			'	"Select ROW_NUMBER() OVER(Order BY " & TipoRicerca & ") As NumeroRiga, " & TipoRicerca & " As TipoRicerca, count(*) FROM informazioniimmagini " &
			'	"where eliminata='N' or eliminata='n' group by " & TipoRicerca & " having count(*) > 1 " &
			'	") As A Where NumeroRiga > " & (Val(Inizio) - 1) & " And NumeroRiga < " & (Inizio + Val(QuanteImmagini))
			Sql = "Select Coalesce(TipoRicerca, '***') As TipoRicerca From ( " &
				"Select ROW_NUMBER() OVER(Order BY B.dimensioni, A.Width, A.Height, B.solonome) As NumeroRiga, " & TipoRicerca & " As TipoRicerca, count(*) " &
				"FROM informazioniimmagini A " &
				"Left Join dati B On B.idtipologia = 1 And A.idCategoria = B.idcategoria And A.idMultimedia = B.progressivo " &
				"Where (B.eliminata='N' Or B.eliminata='n') " & RicercaCategoria & " " &
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
							"Where " & TipoRicerca & " = '" & l & "' And (A.Eliminata='N' Or A.Eliminata='n') And (B.Eliminata='N' Or B.Eliminata='n') " & RicercaCategoria & " " &
							"Order By " & TipoRicerca & ", B.dimensioni, A.Width, A.Height, B.solonome"
						ScriveLogGlobale(NomeFileLog, "Ricerco per lista: " & Sql)
						Rec = db.LeggeQuery(Mp, Sql, ConnessioneSql)
						If TypeOf (Rec) Is String Then
							Ok = False
							ScriveLogGlobale(NomeFileLog, "ERROR: " & Rec)
						End If
						If Ok Then
							Dim q As Integer = 0

							Do Until Rec.Eof
								Dim P As String = Rec("Percorso").Value
								Dim N As String = Rec("NomeFile").Value

								If Not P.ToUpper.Contains("/VIDEOS/") And Not N.ToUpper.Contains("/VIDEOS/") Then
									Ritorno &= ScrittaRitorno & ";"
									Ritorno &= Rec("idCategoria").Value & ";"
									Ritorno &= Rec("idMultimedia").Value & ";"
									Ritorno &= Rec("Hash").Value & ";"
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
									Ritorno &= "§"

									q += 1
								End If

								Rec.MoveNext
							Loop
							Rec.Close

							ScriveLogGlobale(NomeFileLog, "Righe rilevate per lista: " & l & " -> " & q)

						End If
					End If
				Next

				Ritorno = Ritorno & "*" & QuanteRigheTotali
			End If
		End If

		Return Ritorno
	End Function
End Module
