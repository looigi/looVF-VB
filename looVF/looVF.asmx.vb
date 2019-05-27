Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://looVF.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class looVF
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaFiles() As String
		Dim gf As New GestioneFilesDirectory
		Dim sPathsVideo As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PercorsiVideo.txt")
		'If sPathsVideo = "" Then
		'	Return "ERROR: Nessun path video presente"
		'End If
		Dim sPathsImm As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PercorsiImmagini.txt")
		'If sPathsImm = "" Then
		'	Return "ERROR: Nessun path immagini presente"
		'End If
		Dim PathVideo() As String = sPathsVideo.Split(";")
		Dim PathImmagini() As String = sPathsImm.Split(";")
		Dim FilesVideo As List(Of StrutturaFiles) = New List(Of StrutturaFiles)
		Dim FilesImmagini As List(Of StrutturaFiles) = New List(Of StrutturaFiles)

		For Each p As String In PathVideo
			gf.ScansionaDirectorySingola(p)
			Dim qFiles As Integer = gf.RitornaQuantiFilesRilevati
			Dim Files() As String = gf.RitornaFilesRilevati
			For i As Integer = 1 To qFiles
				Dim sf As New StrutturaFiles

				sf.NomeFile = Files(i)
				sf.DimensioniFile = FileLen(Files(i))
				sf.DataFile = FileDateTime(Files(i))

				FilesVideo.Add(sf)
			Next
		Next

		For Each p As String In PathImmagini
			gf.ScansionaDirectorySingola(p)
			Dim qFiles As Integer = gf.RitornaQuantiFilesRilevati
			Dim Files() As String = gf.RitornaFilesRilevati
			For i As Integer = 1 To qFiles
				Dim sf As New StrutturaFiles

				sf.NomeFile = Files(i)
				sf.DimensioniFile = FileLen(Files(i))
				sf.DataFile = FileDateTime(Files(i))

				FilesImmagini.Add(sf)
			Next
		Next

		gf.CreaDirectoryDaPercorso(Server.MapPath(".") & "\Temp\")

		Dim NomeFile As String = Server.MapPath(".") & "\Temp\" & Now.ToString.Replace(":", "_").Replace(" ", "_").Replace("\", "_").Replace("/", "_") & ".txt"
		gf.ApreFileDiTestoPerScrittura(NomeFile)

		For Each sf As StrutturaFiles In FilesVideo
			Dim Stringa As String = "Video;" & sf.NomeFile & ";" & sf.DimensioniFile.ToString & ";" & sf.DataFile & ";§"
			gf.ScriveTestoSuFileAperto(Stringa)
		Next
		For Each sf As StrutturaFiles In FilesImmagini
			Dim Stringa As String = "Pic;" & sf.NomeFile & ";" & sf.DimensioniFile.ToString & ";" & sf.DataFile & ";§"
			gf.ScriveTestoSuFileAperto(Stringa)
		Next
		gf.ChiudeFileDiTestoDopoScrittura()

		Dim Ritorno As String = gf.LeggeFileIntero(NomeFile)

		gf = Nothing

		Return Ritorno
	End Function

End Class