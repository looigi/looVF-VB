Imports System.Reflection
Imports ADOR
Imports MySqlConnector

Public Class clsMariaDB
	Dim Conn As MySqlConnection
	Dim StringaConnessione As String

	Public Function apreConnessione(c) As Object
		StringaConnessione = c
		Conn = New MySqlConnection(c)

		Try
			Conn.Open()
		Catch ex As Exception
			Return ex.Message
		End Try

		Return Conn
	End Function

	Public Function ChiudiConn(Conn)
		Conn.Close()

		Return "OK"
	End Function

	Public Function EsegueSql(Sql As String, ModificaQuery As Boolean) As String
		Dim Errore As String = ""
		Dim Sql2 As String

		If ModificaQuery Then
			If TipoDB = "SQLSERVER" Then
				Sql2 = Sql
			Else
				Sql2 = Sql.ToLower()
				Sql2 = Sql2.Replace("[", "")
				Sql2 = Sql2.Replace("]", "")
				Sql2 = Sql2.Replace("dbo.", "")

				Sql2 = Sql2.Replace("generale", "Generale")
				Sql2 = Sql2.Replace("dbvuoto", "DBvuoto")
			End If
		Else
			Sql2 = Sql
		End If

		' Routine che esegue una query sul db
		Try
			Dim cmd As MySqlCommand = New MySqlCommand(Sql2, Conn)
			cmd.ExecuteNonQuery()

			Errore = "OK"
		Catch ex As MySqlException
			Errore = ex.Message
		End Try

		Return Errore
	End Function

	Public Function Lettura(sql As String, ModificaQuery As Boolean) As Object
		Dim Sql2 As String = ""

		If ModificaQuery Then
			If TipoDB = "SQLSERVER" Then
				Sql2 = sql
			Else
				Sql2 = sql.ToLower()
				Sql2 = Sql2.Replace("[", "")
				Sql2 = Sql2.Replace("]", "")
				Sql2 = Sql2.Replace("dbo.", "")

				Sql2 = Sql2.Replace("generale", "Generale")
				Sql2 = Sql2.Replace("dbvuoto", "DBvuoto")
				Sql2 = Sql2.Replace("dbpieno", "dbPieno")
			End If
		Else
			Sql2 = sql
		End If

		Dim cmd As MySqlCommand = New MySqlCommand(Sql2, Conn)
		Dim Ritorno As MySqlDataReader
		Dim rec As Object = Nothing

		Try
			Ritorno = cmd.ExecuteReader()
			Dim theCommand As New DataTable()
			theCommand.Load(Ritorno)
			rec = New clsRecordset(theCommand, Sql2)
		Catch ex As Exception
			rec = "MDB ERROR:" & ex.Message
		End Try

		Return rec
	End Function

	'Public Function ConvertToRecordSet(inTable As DataTable) As Object
	'	Dim recordSet As Recordset = New Recordset ' (CursorLocationEnum.adUseClient)
	'	Dim recordSetFields = recordSet.Fields
	'	Dim inColumns = inTable.Columns

	'	For Each column As DataColumn In inColumns
	'		recordSetFields.Append(column.ColumnName, TranslateType(column.DataType), column.MaxLength, If(column.AllowDBNull, FieldAttributeEnum.adFldIsNullable, FieldAttributeEnum.adFldUnspecified), Nothing)
	'	Next

	'	recordSet.Open(Missing.Value, Missing.Value, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockOptimistic, 0)

	'	For Each row As DataRow In inTable.Rows
	'		recordSet.AddNew(Missing.Value,
	'						  Missing.Value)
	'		For columnIndex As Integer = 0 To inColumns.Count - 1
	'			recordSetFields(columnIndex).Value = row(columnIndex)
	'		Next
	'	Next

	'	If Not recordSet.Eof() Then
	'		recordSet.MoveFirst()
	'	End If

	'	Return recordSet
	'End Function

	'Private Function TranslateType(ByVal columnDataType As IReflect) As DataTypeEnum
	'	Select Case columnDataType.UnderlyingSystemType.ToString()
	'		Case "System.Boolean"
	'			Return DataTypeEnum.adBoolean
	'		Case "System.Byte"
	'			Return DataTypeEnum.adUnsignedTinyInt
	'		Case "System.Char"
	'			Return DataTypeEnum.adChar
	'		Case "System.DateTime"
	'			Return DataTypeEnum.adDate
	'		Case "System.Decimal"
	'			Return DataTypeEnum.adCurrency
	'		Case "System.Double"
	'			Return DataTypeEnum.adDouble
	'		Case "System.Int16"
	'			Return DataTypeEnum.adSmallInt
	'		Case "System.Int32"
	'			Return DataTypeEnum.adInteger
	'		Case "System.Int64"
	'			Return DataTypeEnum.adBigInt
	'		Case "System.SByte"
	'			Return DataTypeEnum.adTinyInt
	'		Case "System.Single"
	'			Return DataTypeEnum.adSingle
	'		Case "System.UInt16"
	'			Return DataTypeEnum.adUnsignedSmallInt
	'		Case "System.UInt32"
	'			Return DataTypeEnum.adUnsignedInt
	'		Case "System.UInt64"
	'			Return DataTypeEnum.adUnsignedBigInt
	'		Case Else
	'			Return DataTypeEnum.adVarChar
	'	End Select
	'End Function
End Class
