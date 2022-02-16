Public Class GestioneDB
    Private ConnessioneMDB As String
    Private ConnessioneSQL As String

    Public Function ProvaConnessione(Connessione As String) As String
        Dim Conn As Object = CreateObject("ADODB.Connection")

        Try
            Conn.Open(Connessione)
            Conn.Close()

            Conn = Nothing
            Return ""
        Catch ex As Exception
            Dim H As HttpApplication = HttpContext.Current.ApplicationInstance
            Dim StringaPassaggio As String

            StringaPassaggio = "?Errore=Apertura DB"
            StringaPassaggio = StringaPassaggio & "&Utente=" & H.Session("idUtente")
            StringaPassaggio = StringaPassaggio & "&Chiamante=" & H.Request.CurrentExecutionFilePath.ToUpper.Trim
            StringaPassaggio = StringaPassaggio & "&Errore=" & ex.Message
            H.Response.Redirect("Errore.aspx" & StringaPassaggio)

            Return ex.Message
        End Try
    End Function

    Public Function LeggeImpostazioniDiBase() As Boolean
        Dim Ritorno As String
        Dim Ok As Boolean = True
        Dim CosaCercare As String
        Dim Conn As String = ""

		'If ModalitaLocale = True Then
		CosaCercare = "SQLConnectionStringLOCALElooVF"
		'Else
		'CosaCercare = "SQLConnectionStringWEB"
		'End If

		' Impostazioni di base
		Dim ListaConnessioni As ConnectionStringSettingsCollection = ConfigurationManager.ConnectionStrings

        If ListaConnessioni.Count <> 0 Then
            ' Get the collection elements. 
            For Each Connessioni As ConnectionStringSettings In ListaConnessioni
                Dim Nome As String = Connessioni.Name
                Dim Provider As String = Connessioni.ProviderName
                Dim connectionString As String = Connessioni.ConnectionString

                If Nome = CosaCercare Then
                    Conn = "Provider=" & Provider & ";" & connectionString

                    Exit For
                End If
            Next
        End If

        If Conn = "" Then
            ' Response.Redirect("errore_ErroreImprevisto.aspx?Errore=Impostazioni di connessione al DB non valide&Chiamante=" & Request.CurrentExecutionFilePath.ToUpper.Trim & "&Sql=")
            Ok = False
        Else
            Ritorno = ProvaConnessione(Conn)
            If Ritorno <> "" Then
                ' Response.Redirect("errore_ErroreImprevisto.aspx?Errore=" & Ritorno & "&Chiamante=" & Request.CurrentExecutionFilePath.ToUpper.Trim & "&Sql=")
                Ok = False
            Else
                ConnessioneSQL = Conn
            End If
            ' Impostazioni di base
        End If

        Return Ok
    End Function

    Public Function ApreDB() As Object
        ' Routine che apre il DB e vede se ci sono errori
        Dim Conn As Object = CreateObject("ADODB.Connection")

        Try
            Conn.Open(ConnessioneSQL)
            Conn.CommandTimeout = 0
        Catch ex As Exception
            Dim H As HttpApplication = HttpContext.Current.ApplicationInstance
            Dim StringaPassaggio As String

            StringaPassaggio = "?Errore=Apertura DB"
            StringaPassaggio = StringaPassaggio & "&Utente=" & H.Session("idUtente")
            StringaPassaggio = StringaPassaggio & "&Chiamante=" & H.Request.CurrentExecutionFilePath.ToUpper.Trim
            StringaPassaggio = StringaPassaggio & "&Sql="
            H.Response.Redirect("Errore.aspx" & StringaPassaggio)
        End Try

        Return Conn
    End Function

    Private Function ControllaAperturaConnessione(ByRef Conn As Object) As Boolean
        Dim Ritorno As Boolean = False

        If Conn Is Nothing Then
            Ritorno = True
            Conn = ApreDB(ConnessioneSQL)
        End If

        Return Ritorno
    End Function

    Public Function EsegueSql(ByVal Conn As Object, ByVal Sql As String) As String
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(Conn)
        Dim Ritorno As String = ""

        ' Routine che esegue una query sul db
        Try
            Conn.Execute(Sql)
        Catch ex As Exception
			'Dim H As HttpApplication = HttpContext.Current.ApplicationInstance
			'Dim StringaPassaggio As String

			'StringaPassaggio = "?Errore=Errore esecuzione query: " & Err.Description
			'StringaPassaggio = StringaPassaggio & "&Utente=" & H.Session("idUtente")
			'StringaPassaggio = StringaPassaggio & "&Chiamante=" & H.Request.CurrentExecutionFilePath.ToUpper.Trim
			'StringaPassaggio = StringaPassaggio & "&Sql=" & Sql
			'H.Response.Redirect("Errore.aspx" & StringaPassaggio)
			Ritorno = ex.Message
		End Try

		ChiudeDB(AperturaManuale, Conn)

        Return Ritorno
    End Function

    Public Function EsegueSqlSenzaTRY(ByVal Conn As Object, ByVal Sql As String) As String
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(Conn)
        Dim Ritorno As String = ""

        Conn.Execute(Sql)

        ChiudeDB(AperturaManuale, Conn)

        Return Ritorno
    End Function

    Private Sub ChiudeDB(ByVal TipoApertura As Boolean, ByRef Conn As Object)
        If TipoApertura = True Then
            Conn.Close()
        End If
    End Sub

    Public Function LeggeQuery(ByVal Conn As Object, ByVal Sql As String) As Object
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(Conn)
        Dim Rec As Object = CreateObject("ADODB.Recordset")

        Try
            Rec.Open(Sql, Conn)
        Catch ex As Exception
            Rec = Nothing

            Dim H As HttpApplication = HttpContext.Current.ApplicationInstance
            Dim StringaPassaggio As String

            StringaPassaggio = "?Errore=Errore query: " & Err.Description
            StringaPassaggio = StringaPassaggio & "&Utente=" & H.Session("idUtente")
            StringaPassaggio = StringaPassaggio & "&Chiamante=" & H.Request.CurrentExecutionFilePath.ToUpper.Trim
            StringaPassaggio = StringaPassaggio & "&Sql=" & Sql
            H.Response.Redirect("Errore.aspx" & StringaPassaggio)
        End Try

        ChiudeDB(AperturaManuale, Conn)

        Return Rec
    End Function
End Class
