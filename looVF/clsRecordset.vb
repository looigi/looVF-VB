Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.IO
Imports System.Runtime.Remoting
Imports System.Threading.Tasks
Imports System.Xml
Imports MySqlConnector

Public Class clsRecordset
	Inherits DataTable

	Private DT As New DataTable
	Private RigaAttuale As Long = 0
	Private Fine As Boolean = False
	Private Inizio As Boolean = False
	Private RigaSelezionata As DataRow
	Private Sql As String

	Public Structure Ritorno
		Dim Value As String
	End Structure

	Default Property myProperty(index As Object) As Ritorno
		Get
			Return PrendeCampo(index)
		End Get

		Set(value As Ritorno)
		End Set
	End Property

	Sub New(d As DataTable, S As String)
		Try
			DT = d
			Fine = False
			Inizio = True
			RigaAttuale = 0
			Sql = S
			If DT.Rows.Count = 0 Then
				Fine = True
			Else
				ImpostaRiga()
			End If
		Catch ex As Exception
			' Return ex.Message 
		End Try
	End Sub

	Public Function PrendeCampo(NomeCampo As Object) As Ritorno
		Dim r As New Ritorno
		Try
			If RigaSelezionata.Item(NomeCampo) Is DBNull.Value Then
				r.Value = ""
			Else
				r.Value = RigaSelezionata.Item(NomeCampo)
			End If
		Catch ex As Exception
			r.Value = StringaErrore & NomeCampo & ". Lunghezza DT: " & DT.Rows.Count & ". Riga Attuale: " & RigaAttuale & ". EOF: " & Fine & ". SQL: " & Sql
		End Try

		Return r
	End Function

	Public Sub Close()
		DT = Nothing
	End Sub

	Private Sub ImpostaRiga()
		RigaSelezionata = DT.Rows(RigaAttuale)
	End Sub

	Public Function Bof() As Boolean
		Return Fine
	End Function

	Public Function Eof() As Boolean
		Return Fine
	End Function

	Public Sub MovePrevious()
		RigaAttuale -= 1
		If RigaAttuale < 0 Then
			Inizio = True
		Else
			ImpostaRiga()
		End If
	End Sub

	Public Sub MoveNext()
		RigaAttuale += 1
		If RigaAttuale > DT.Rows.Count - 1 Then
			Fine = True
		Else
			ImpostaRiga()
		End If
	End Sub

	Public Sub MoveFirst()
		RigaAttuale = 0
		Inizio = True
		ImpostaRiga()
	End Sub

	Public Sub MoveLast()
		RigaAttuale = DT.Rows.Count - 1
		ImpostaRiga()
	End Sub
End Class
