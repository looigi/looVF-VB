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

	Public Function RitornaUguaglianze(Mp As String, db As clsGestioneDB, ConnessioneSql As String, TipoRicerca As String, ScrittaRitorno As String, idCategoria As String, QuanteImmagini As String, Inizio As String, StringaRicerca As String) As String
		Dim Ritorno As String = ""
		Dim Ok As Boolean = True
		Dim Rec As Object
		Dim Sql As String = ""

		If TipoRicerca = "STRINGA" Then
			Sql = "Select Coalesce(Count(*),0) FROM informazioniimmagini A " &
				"Left Join dati B On B.idtipologia = 1 And A.idCategoria = B.idcategoria And A.idMultimedia = B.progressivo " &
				"Left Join categorie C On A.idCategoria = C.idcategoria And C.idtipologia = B.idTipologia " &
				"Where (B.eliminata='N' Or B.eliminata='n') And A.idCategoria=" & idCategoria & " And Instr(C.Percorso, '/Videos/') = 0 " &
				"And (InStr(C.Percorso, '" & StringaRicerca.Replace("'", "''") & "') > 0 Or Instr(B.NomeFile, '" & StringaRicerca.Replace("'", "''") & "') > 0) "
			' Return Sql
			Rec = db.LeggeQuery(Mp, Sql, ConnessioneSql)
			If TypeOf (Rec) Is String Then
				Ok = False
			End If
			If Ok Then
				Dim QuanteRigheTotali As Integer = Rec(0).Value
				Rec.Close

				'If QuanteRigheTotali > 1000 Then
				'	Return "ERROR: Troppe righe ritornate (" & QuanteRigheTotali & ")"
				'End If

				Sql = "Select * From (" &
					"Select ROW_NUMBER() OVER(Order BY B.dimensioni, A.Width, A.Height, B.solonome) As NumeroRiga, A.*, b.nomefile, c.percorso, c.protetta, Coalesce(d.progressivo, '') As preferito, Coalesce(e.progressivo, '') As preferitoprot, b.solonome, b.Dimensioni " &
					"From informazioniimmagini A " &
					"left join dati b On b.idtipologia = 1 And A.idCategoria = b.idCategoria And A.idMultimedia=b.progressivo " &
					"left join categorie c On c.idtipologia = 1 And A.idCategoria = c.idcategoria " &
					"left join preferiti d On d.idTipologia = 1 And A.idCategoria = d.idCategoria And A.idMultimedia=d.progressivo " &
					"left join preferitiprot e On e.idTipologia = 1 And A.idCategoria = e.idCategoria And A.idMultimedia=e.progressivo " &
					"Where (A.Eliminata='N' Or A.Eliminata='n') And (b.Eliminata='N' Or b.Eliminata='n') And A.idCategoria=" & idCategoria & " And Instr(c.Percorso, '/Videos/') = 0 " &
					"And (Instr(c.Percorso, '" & StringaRicerca.Replace("'", "''") & "') > 0 Or Instr(b.NomeFile, '" & StringaRicerca.Replace("'", "''") & "') > 0) " &
					") As A Where NumeroRiga > " & (Val(Inizio) - 1) & " And NumeroRiga < " & (Inizio + Val(QuanteImmagini)) & " " &
					"Order By percorso, nomefile, A.dimensioni, A.Width, A.Height"
				Rec = db.LeggeQuery(Mp, Sql, ConnessioneSql)
				If TypeOf (Rec) Is String Then
					Ok = False
				End If
				If Ok Then
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
						End If

						Rec.MoveNext
					Loop
					Rec.Close

					Ritorno &= "*" & QuanteRigheTotali
				End If
			End If

			Return Ritorno
		End If

		Sql = "Select Coalesce(Count(*),0) From ( " &
			"Select " & TipoRicerca & " As TipoRicerca FROM informazioniimmagini A " &
			"Left Join dati B On B.idtipologia = 1 And A.idCategoria = B.idcategoria And A.idMultimedia = B.progressivo " &
			"Where (B.eliminata='N' Or B.eliminata='n') " &
			"Group By " & TipoRicerca & " " &
			"Having Count(*) > 1 " &
			") As A"
		'"Left Join categorie C On A.idCategoria = C.idcategoria And C.idtipologia = B.idTipologia " &
		'"Where (B.eliminata='N' Or B.eliminata='n') And Instr(C.Percorso, '/Videos/') = 0 " &
		'Return Sql
		Rec = db.LeggeQuery(Mp, Sql, ConnessioneSql)
		If TypeOf (Rec) Is String Then
			Ok = False
		End If
		If Ok Then
			Dim QuanteRigheTotali As Integer = Rec(0).Value
			Rec.Close

			'Sql = "Select TipoRicerca From ( " &
			'	"Select ROW_NUMBER() OVER(Order BY " & TipoRicerca & ") As NumeroRiga, " & TipoRicerca & " As TipoRicerca, count(*) FROM informazioniimmagini " &
			'	"where eliminata='N' or eliminata='n' group by " & TipoRicerca & " having count(*) > 1 " &
			'	") As A Where NumeroRiga > " & (Val(Inizio) - 1) & " And NumeroRiga < " & (Inizio + Val(QuanteImmagini))
			Sql = "Select Coalesce(TipoRicerca, '***') As TipoRicerca From ( " &
				"Select ROW_NUMBER() OVER(Order BY B.dimensioni, A.Width, A.Height, B.solonome) As NumeroRiga, " & TipoRicerca & " As TipoRicerca, count(*) " &
				"FROM informazioniimmagini A " &
				"Left Join dati B On B.idtipologia = 1 And A.idCategoria = B.idcategoria And A.idMultimedia = B.progressivo " &
				"Where B.eliminata='N' Or B.eliminata='n' " &
				"Group By " & TipoRicerca & " " &
				"Having Count(*) > 1 " &
				") As A Where NumeroRiga > " & (Val(Inizio) - 1) & " And NumeroRiga < " & (Inizio + Val(QuanteImmagini))

			'"Left Join categorie C On A.idCategoria = C.idcategoria And C.idtipologia = B.idTipologia " &
			'"Where B.eliminata='N' Or B.eliminata='n' And Instr(C.Percorso, '/Videos/') = 0 " &
			Rec = db.LeggeQuery(Mp, Sql, ConnessioneSql)
			If TypeOf (Rec) Is String Then
				Ok = False
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

				For Each l As String In Lista
					Ok = True

					Sql = "Select A.*, b.nomefile, c.percorso, c.protetta, Coalesce(d.progressivo, '') As preferito, Coalesce(e.progressivo, '') As preferitoprot, b.solonome, b.Dimensioni " &
						"From informazioniimmagini A " &
						"left join dati b On b.idtipologia = 1 And A.idCategoria = b.idCategoria And A.idMultimedia=b.progressivo " &
						"left join categorie c On c.idtipologia = 1 And A.idCategoria = c.idcategoria " &
						"left join preferiti d On d.idTipologia = 1 And A.idCategoria = d.idCategoria And A.idMultimedia=d.progressivo " &
						"left join preferitiprot e On e.idTipologia = 1 And A.idCategoria = e.idCategoria And A.idMultimedia=e.progressivo " &
						"Where " & TipoRicerca & " = '" & l & "' And (A.Eliminata='N' Or A.Eliminata='n') And (B.Eliminata='N' Or B.Eliminata='n') And A.idCategoria=" & idCategoria & " " &
						"Order By " & TipoRicerca & ", b.dimensioni, A.Width, A.Height, b.solonome"
					Rec = db.LeggeQuery(Mp, Sql, ConnessioneSql)
					If TypeOf (Rec) Is String Then
						Ok = False
					End If
					If Ok Then
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
							End If


							Rec.MoveNext
						Loop
						Rec.Close
					End If
				Next

				Ritorno = Ritorno & "*" & QuanteRigheTotali
			End If
		End If

		Return Ritorno
	End Function
End Module
