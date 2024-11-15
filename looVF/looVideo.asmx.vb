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
<System.Web.Services.WebService(Namespace:="http://looVideo.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class looVideo
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RefreshVideo() As String
		Dim Ritorno As String = "*"
		Dim gf As New GestioneFilesDirectory
		Dim Db As New clsGestioneDB(TipoDB)
		Dim ConnessioneSQL As String = Db.LeggeImpostazioniDiBaseVideo()
		If ConnessioneSQL <> "" Then
			Dim Path As String = gf.LeggeFileIntero(Server.MapPath(".") & "\PercorsiVideo.txt")
			Dim Sql As String = ""
			Dim Rec As Object

			Sql = "Delete From video"
			Dim Rit As String = Db.EsegueSql(Server.MapPath("."), Sql, ConnessioneSQL, False)
			If Not Ritorno.Contains("ERROR:") Then

			End If
		End If

		Return Ritorno
	End Function

End Class