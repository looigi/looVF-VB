Module mdlLooVF
	Public FaiLog As Boolean = False

	Public ListaImmagini As List(Of String) = New List(Of String)
	Public QuanteImmaginiSfondi As Integer = 0
	Public ContatoreRiletturaImmagini As Integer = 0
	Public StaLeggendoImmagini As Boolean = False
	Public TipoDB As String = "MARIADB"
	Public StringaErrore As String = "ERROR: "
	Public effettuaLog As Boolean = True
	Public nomeFileLogGenerale As String = ""
	Public listaLog As New List(Of String)
	Public timerLog As New Timers.Timer
	Public timerConv As New Timers.Timer
	Public VecchiaRicerca As String = ""
	Public VecchioQuante As Long
	Public UltimoMultimediaImm As Long
	Public UltimoMultimediaVid As Long
	Public StaEffettuandoConversioneAutomatica As Boolean = False
	Public StaEffettuandoConversioneAutomaticaFinale As Boolean = False
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

	Public Function dataAttuale() As String
		Return Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00")
	End Function

	Public Sub ScriveLogGlobale(Path As String, Cosa As String)
		If FaiLog = True Then
			Dim gf As New GestioneFilesDirectory

			gf.ApreFileDiTestoPerScrittura(Path)
			gf.ScriveTestoSuFileAperto(dataAttuale() & ": " & Cosa)
			gf.ChiudeFileDiTestoDopoScrittura()
		End If
	End Sub
End Module
