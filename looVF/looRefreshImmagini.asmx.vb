Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.IO
Imports System.Threading
Imports System.Timers
Imports System.Net
Imports System.Web.Script.Serialization
Imports looVF.GestioneImmagini
Imports System.Drawing

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://looRefresh.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class looRefreshImmagini
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaListaImmagini() As String
		Dim Barra As String

		If TipoDB = "SQLSERVER" Then
			Barra = "\"
		Else
			Barra = "/"
		End If

		Dim gf As New GestioneFilesDirectory
		Dim P As String = gf.LeggeFileIntero(Server.MapPath(".") & Barra & "PercorsiSfondi.txt")
		Dim pp() As String = P.Split(";")
		Dim PathSfondi As String = pp(0).Replace(vbCrLf, "")
		Dim Ritorno As String = ""

		gf.ScansionaDirectorySingola(PathSfondi)
		Dim Filetti() As String = gf.RitornaFilesRilevati
		Dim qFiletti As String = gf.RitornaQuantiFilesRilevati

		For i As Integer = 1 To qFiletti
			Ritorno &= Filetti(i).Replace(PathSfondi & Barra, "").Replace(";", "*PV*") & ";" & gf.TornaDimensioneFile(Filetti(i)) & ";" & gf.TornaDataDiCreazione(Filetti(i)) & "§"
		Next

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ConverteImmagineBase64(Path As String) As String
		Dim Barra As String
		Dim ControBarra As String

		If TipoDB = "SQLSERVER" Then
			Barra = "\"
			ControBarra = "/"
		Else
			Barra = "/"
			ControBarra = "\"
		End If

		Dim gf As New GestioneFilesDirectory
		Dim P As String = gf.LeggeFileIntero(Server.MapPath(".") & Barra & "PercorsiSfondi.txt")
		Dim pp() As String = P.Split(";")
		Dim PathSfondi As String = pp(0).Replace(vbCrLf, "")
		Dim PathImm As String = PathSfondi & Barra & Path.Replace(ControBarra, Barra)

		Return convertImageToBase64(PathImm)
	End Function

	<WebMethod()>
	Public Function ScriveImmagine(Path As String, Base64 As String) As String
		Dim Barra As String
		Dim ControBarra As String

		If TipoDB = "SQLSERVER" Then
			Barra = "\"
			ControBarra = "/"
		Else
			Barra = "/"
			ControBarra = "\"
		End If

		Dim gf As New GestioneFilesDirectory
		Dim P As String = gf.LeggeFileIntero(Server.MapPath(".") & Barra & "PercorsiSfondi.txt")
		Dim pp() As String = P.Split(";")
		Dim PathSfondi As String = pp(0).Replace(vbCrLf, "").Replace("PathSfondi=", "")
		Dim PathImm As String = PathSfondi & Barra & Path.Replace(ControBarra, Barra)
		Dim Cartella As String = gf.TornaNomeDirectoryDaPath(PathImm)
		gf.CreaDirectoryDaPercorso(Cartella & Barra)

		Dim Image1 As System.Drawing.Image = convertByteToImage(Base64)

		gf.EliminaFileFisico(PathImm)
		Dim SaveImage As New System.Drawing.Bitmap(Image1)
		SaveImage.Save(PathImm, System.Drawing.Imaging.ImageFormat.Jpeg)
		SaveImage.Dispose()
		If gf.EsisteFile(PathImm) Then
			Return "OK"
		Else
			Return "ERROR: File non copiato"
		End If
	End Function
End Class