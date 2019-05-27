Public Class StrutturaFiles
	Private sNomeFile As String = ""
	Private lDimeFile As Long = -1
	Private dData As Date = New Date
	Private sCategoria As Integer = -1

	Public Property NomeFile As String
		Get
			Return Me.sNomeFile
		End Get
		Set(ByVal Value As String)
			Me.sNomeFile = Value
		End Set
	End Property

	Public Property Categoria As Integer
		Get
			Return Me.sCategoria
		End Get
		Set(ByVal Value As Integer)
			Me.sCategoria = Value
		End Set
	End Property

	Public Property DimensioniFile As Long
		Get
			Return lDimeFile
		End Get
		Set(ByVal Value As Long)
			lDimeFile = Value
		End Set
	End Property

	Public Property DataFile As Date
		Get
			Return dData
		End Get
		Set(ByVal Value As Date)
			dData = Value
		End Set
	End Property
End Class
