Public Class StrutturaFiles
	Private sNomeFile As String = ""
	Private lDimeFile As Long = -1
	Private dData As Date = New Date

	Public Property NomeFile As String
		Get
			Return Me.sNomeFile
		End Get
		Set(ByVal Value As String)
			Me.sNomeFile = Value
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
