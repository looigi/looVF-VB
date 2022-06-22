Module mdlLooVF
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
	Public VecchiaRicerca As String = ""
	Public VecchioQuante As Long

	Public Function dataAttuale() As String
		Return Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00")
	End Function
End Module
