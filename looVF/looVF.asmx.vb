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
	Public Function CreaDirectoryVirtuale(App As String, Path As String) As String
		Dim IIs As New IIS
		' IIs.CreateVirtualDir("LocalHost", "looVF", "Appoggio", "C:\Appoggio")
		IIs.CreateVirtualDir("Default Web Site", "Appoggi2", "D:\Appoggio")
	End Function

	<WebMethod()>
	Public Function RitornaFiles(Aggiorna As String) As String
		Dim gf As New GestioneFilesDirectory

		If Aggiorna = "N" Then
			Dim Ok As Boolean = False

			gf.ScansionaDirectorySingola(Server.MapPath(".") & "\Temp\")
			Dim Filetti() As String = gf.RitornaFilesRilevati
			Dim qFiles As Integer = gf.RitornaQuantiFilesRilevati
			Dim MaxDatella As Date = "01/01/1970 00:00:00"
			For i As Integer = 1 To qFiles
				Dim Nome As String = gf.TornaNomeFileDaPath(Filetti(i))
				Dim Campi() As String = Nome.Split("_")
				Nome = Campi(0) & "/" & Campi(1) & "/" & Campi(2) & " " & Campi(3) & ":" & Campi(4) & ":" & Campi(5)
				Dim Datella As Date = Nome.Replace(".txt", "")
				If Datella > MaxDatella Then
					MaxDatella = Datella
					Ok = True
				End If
			Next

			If Ok Then
				Dim RitornoAggiorna As String = Server.MapPath(".") & "\Temp\" & MaxDatella.ToString.Replace(":", "_").Replace(" ", "_").Replace("\", "_").Replace("/", "_") & ".txt"

				Return RitornoAggiorna
			End If
		End If

		Dim sPathsVideo As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PercorsiVideo.txt")
		'If sPathsVideo = "" Then
		'	Return "ERROR: Nessun path video presente"
		'End If
		Dim sPathsImm As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PercorsiImmagini.txt")
		'If sPathsImm = "" Then
		'	Return "ERROR: Nessun path immagini presente"
		'End If
		Dim PathVideo() As String = sPathsVideo.Split("§")
		Dim PathImmagini() As String = sPathsImm.Split("§")
		Dim FilesVideo As List(Of StrutturaFiles) = New List(Of StrutturaFiles)
		Dim FilesImmagini As List(Of StrutturaFiles) = New List(Of StrutturaFiles)
		Dim Conta As Integer

		Conta = 0
		For Each p As String In PathVideo
			If p.Trim <> "" Then
				Dim pp() As String = p.Split(";")

				If Strings.Right(pp(1), 1) <> "\" Then
					pp(1) &= "\"
				End If
				gf.ScansionaDirectorySingola(pp(1))
				Dim qFiles As Integer = gf.RitornaQuantiFilesRilevati
				Dim Files() As String = gf.RitornaFilesRilevati
				Conta += 1
				For i As Integer = 1 To qFiles
					Dim sf As New StrutturaFiles

					sf.Categoria = Conta
					sf.NomeFile = Files(i).Replace(pp(1), "")
					sf.DimensioniFile = FileLen(Files(i))
					sf.DataFile = FileDateTime(Files(i))

					FilesVideo.Add(sf)
				Next
			End If
		Next

		For Each p As String In PathImmagini
			If p.Trim <> "" Then
				Dim pp() As String = p.Split(";")

				If Strings.Right(pp(1), 1) <> "\" Then
					pp(1) &= "\"
				End If
				gf.ScansionaDirectorySingola(pp(1))
				Dim qFiles As Integer = gf.RitornaQuantiFilesRilevati
				Dim Files() As String = gf.RitornaFilesRilevati
				Conta += 1
				For i As Integer = 1 To qFiles
					Dim sf As New StrutturaFiles

					sf.Categoria = Conta
					sf.NomeFile = Files(i).Replace(pp(1), "")
					sf.DimensioniFile = FileLen(Files(i))
					sf.DataFile = FileDateTime(Files(i))

					FilesImmagini.Add(sf)
				Next
			End If
		Next

		gf.CreaDirectoryDaPercorso(Server.MapPath(".") & "\Temp\")

		Dim NomeFile As String = Server.MapPath(".") & "\Temp\" & Now.ToString.Replace(":", "_").Replace(" ", "_").Replace("\", "_").Replace("/", "_") & ".txt"
		gf.ApreFileDiTestoPerScrittura(NomeFile)

		For Each p As String In PathVideo
			If p.Trim <> "" Then
				Dim pp() As String = p.Split(";")
				Dim Stringa As String = "CategoriaVideo;" & pp(0) & ";§"
				gf.ScriveTestoSuFileAperto(Stringa)
			End If
		Next

		For Each p As String In PathImmagini
			If p.Trim <> "" Then
				Dim pp() As String = p.Split(";")
				Dim Stringa As String = "CategoriaImmagini;" & pp(0) & ";§"
				gf.ScriveTestoSuFileAperto(Stringa)
			End If
		Next

		For Each sf As StrutturaFiles In FilesVideo
			Dim Stringa As String = "Video;" & sf.Categoria & ";" & sf.NomeFile & ";" & sf.DimensioniFile.ToString & ";" & sf.DataFile & ";§"
			gf.ScriveTestoSuFileAperto(Stringa)
		Next

		For Each sf As StrutturaFiles In FilesImmagini
			Dim Stringa As String = "Pic;" & sf.Categoria & ";" & sf.NomeFile & ";" & sf.DimensioniFile.ToString & ";" & sf.DataFile & ";§"
			gf.ScriveTestoSuFileAperto(Stringa)
		Next
		gf.ChiudeFileDiTestoDopoScrittura()

		' Dim Ritorno As String = gf.LeggeFileIntero(NomeFile)

		gf = Nothing

		Dim Ritorno As String = NomeFile.Replace(Server.MapPath(".") & "\Temp\", "")

		Return Ritorno
	End Function

End Class