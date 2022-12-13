Imports System.IO
Imports System.Drawing.Imaging
Imports System.Drawing
Imports System.Drawing.Image
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows.Forms
Imports System.Threading
Imports System.Drawing.Drawing2D
Imports LevDan.Exif

Public Class GestioneImmagini
	Private NomeBNRid As String
	Private NomeRid As String
	Private Const qX As Integer = 50
	Private Const qY As Integer = 50
	Private Const quadrettoX As Integer = 3
	Private Const quadrettoY As Integer = 3
	Private Const Divisore As Integer = 32
	Private C(2) As Integer
	Private Colore As Color
	Private r As Integer
	Private g As Integer
	Private b As Integer

	Public NomeTag() As String
	Public idTag() As Integer
	Public QuantiTag As Integer

	Public Structure StrutturaJPG
		Public Sezione1 As String
		Public Punti As String
		Public Width As Integer
		Public Height As Integer
		Public DataOra As String
		Public PuntiDiagonale As String
		Public PuntiCornice As String
		Public Sezione2 As String
		Public Descrizione As String
		Public Commento As String
		Public Software As String
		Public DateTimeTag As String
		Public XRes As String
		Public YRes As String
		Public Hash As String
	End Structure

	Private Nero As Color = Color.FromArgb(0, 0, 0)
	Private GrigioChiaro As Color = Color.FromArgb(186, 186, 186)
	Private Grigio As Color = Color.FromArgb(128, 128, 128)
	Private Bianco As Color = Color.FromArgb(255, 255, 255)

	Public Enum ShadowDirections As Integer
		TOP_RIGHT = 1
		BOTTOM_RIGHT = 2
		BOTTOM_LEFT = 3
		TOP_LEFT = 4
	End Enum

	Private valoreColore As Integer = 127

	Public Function RitagliaBordoDaImmagine(Imm As Image, QuantoBordo As Integer) As Image
		Dim sourceBmp As New Bitmap(Imm)
		Dim dX As Integer = Imm.Width - (QuantoBordo / 2)
		Dim dY As Integer = Imm.Height - (QuantoBordo / 2)
		Dim destinationBmp As New Bitmap(dX, dY)
		Dim gr As Graphics = Graphics.FromImage(destinationBmp)
		Dim selectionRectangle As New Rectangle(QuantoBordo, QuantoBordo, Imm.Width - QuantoBordo, Imm.Height - QuantoBordo)
		Dim destinationRectangle As New Rectangle(0, 0, dX, dY)

		gr.DrawImage(sourceBmp, destinationRectangle, selectionRectangle, GraphicsUnit.Pixel)

		Dim RitornoImage As Image = destinationBmp

		gr.Dispose()
		sourceBmp.Dispose()

		sourceBmp = Nothing
		gr = Nothing

		Return RitornoImage
	End Function

	Public Function RidimensionaMantenendoProporzioni(Path As String, Path2 As String, Larghezza As Integer, Optional Adatta As Boolean = True) As String
		Dim Ritorno As String = ""

		Try
			Dim myEncoder As System.Drawing.Imaging.Encoder
			Dim myEncoderParameters As New Imaging.EncoderParameters(1)
			Dim img2 As Bitmap
			Dim ImmaginePiccola22 As Image
			Dim jgpEncoder2 As Imaging.ImageCodecInfo
			Dim myEncoder2 As System.Drawing.Imaging.Encoder
			Dim myEncoderParameters2 As New Imaging.EncoderParameters(1)

			Dim Dime As String = RitornaDimensioneImmagine(Path)
			If Dime.Contains("ERROR") Then
				File.Delete(Path)
				Return "ERROR: file non valido"
				Exit Function
			End If

			Dim Dimensioni() As String = Dime.Split("x")
			Dim x As Integer = Dimensioni(0)
			Dim y As Integer = Dimensioni(1)

			If x > Larghezza Or y > Larghezza Then
				Dim largh As Integer
				Dim alt As Integer

				If Adatta = True Then
					Dim propX As Single = Larghezza / x
					Dim propY As Single = Larghezza / y

					If propX < propY Then
						largh = x * propX
						alt = y * propX
					Else
						largh = x * propY
						alt = y * propY
					End If
				Else
					largh = 100
					alt = 100
				End If

				' img2 = New Bitmap(Path)
				img2 = LoadBitmapSenzaLock(Path)

				ImmaginePiccola22 = New Bitmap(img2, largh, alt)
				img2.Dispose()
				img2 = Nothing

				myEncoder = System.Drawing.Imaging.Encoder.Quality
				jgpEncoder2 = GetEncoder(Imaging.ImageFormat.Jpeg)
				myEncoder2 = System.Drawing.Imaging.Encoder.Quality
				Dim myEncoderParameter2 As New Imaging.EncoderParameter(myEncoder, 97)
				myEncoderParameters2.Param(0) = myEncoderParameter2
				ImmaginePiccola22.Save(Path2, jgpEncoder2, myEncoderParameters2)

				ImmaginePiccola22.Dispose()

				ImmaginePiccola22 = Nothing
				jgpEncoder2 = Nothing
				myEncoderParameter2 = Nothing

				Ritorno = "OK"
			End If
		Catch ex As Exception
			Ritorno = "ERROR: Ridimensiona " & ex.Message
		End Try

		Return Ritorno
	End Function

	Public Function CentraImmagineNelPannello(PannelloX As Integer, PannelloY As Integer, ImmX As Integer, ImmY As Integer) As String
		Dim Ritorno As String = ""

		'If ImmX < PannelloX - 10 And ImmY < PannelloY - 10 Then
		'    Ritorno = ImmX & "x" & ImmY
		'Else
		Dim PercX As Single
		Dim PercY As Single
		'Dim Perc As Single
		Dim nX As Integer
		Dim nY As Integer

		PercX = (PannelloX - 20) / ImmX
		PercY = (PannelloY - 20) / ImmY

		If PercX < PercY Then
			PercY = PercX
		Else
			PercX = PercY
		End If

		nX = ImmX * PercX
		nY = ImmY * PercY

		Ritorno = nX & "x" & nY
		'End If

		Return Ritorno
	End Function

	Public Function CambiaOpacitaImmagine(imgLight As Image, opacity As Single) As Bitmap
		If imgLight Is Nothing = False Then
			Dim bmp As New Bitmap(imgLight.Width, imgLight.Height)
			Dim graphics__1 As Graphics = Graphics.FromImage(bmp)
			Dim colormatrix As New ColorMatrix()
			colormatrix.Matrix33 = opacity
			Dim imgAttribute As New ImageAttributes()
			imgAttribute.SetColorMatrix(colormatrix, ColorMatrixFlag.[Default], ColorAdjustType.Bitmap)
			graphics__1.DrawImage(imgLight, New Rectangle(0, 0, bmp.Width, bmp.Height), 0, 0, imgLight.Width, imgLight.Height,
				GraphicsUnit.Pixel, imgAttribute)
			graphics__1.Dispose()

			Return bmp
		End If
	End Function

	Public Function CaricaImmagineSenzaLockarla(NomeImmagine As String) As Image
		Dim bmp As Image = Nothing
		Dim fs As System.IO.FileStream = Nothing

		Try
			fs = New System.IO.FileStream(NomeImmagine, IO.FileMode.Open, IO.FileAccess.Read)
			bmp = System.Drawing.Image.FromStream(fs)
		Catch ex As Exception
			'Stop
		End Try

		If fs Is Nothing = False Then
			Try
				fs.Close()
				fs.Dispose()
			Catch ex As Exception

			End Try
		End If
		fs = Nothing

		Return bmp
	End Function

	Public Sub ImpostaPathPerLavoro(Path As String)
		NomeBNRid = Path & "\Thumbs\AppoggioBN.Jpg"
		NomeRid = Path & "\Thumbs\Appoggio.Jpg"
	End Sub

	' Public Sub CreaValoreUnivocoImmagine(idImmagine As Long, Db As ACCESS, Conn As Object, Immagine As String, gf As GestioneFilesDirectory)
	' 	If File.Exists(Immagine) = False Then
	' 		Exit Sub
	' 	End If
	' 
	' 	Dim imgImmagine As Image
	' 	Dim Stringona As String
	' 
	' 	Ridimensiona(Immagine, NomeRid, qX, qY)
	' 	ConverteImmaginInBN(NomeRid, NomeBNRid)
	' 
	' 	Try
	' 		Kill(NomeRid)
	' 	Catch ex As Exception
	' 
	' 	End Try
	' 
	' 	imgImmagine = New Bitmap(NomeBNRid)
	' 
	' 	Stringona = ""
	' 
	' 	Dim Valore As String
	' 
	' 	For I As Integer = 1 To qX Step quadrettoX
	' 		For k As Integer = 1 To qY Step quadrettoY
	' 			Colore = DirectCast(imgImmagine, Bitmap).GetPixel(k, I)
	' 
	' 			r = Colore.R '* 0.49999999999999994
	' 			g = Colore.G '* 0.49000000000000005
	' 			b = Colore.B '* 0.49999999999999595
	' 
	' 			'r = CInt((r \ Divisore)) * Divisore
	' 			'b = CInt((b \ Divisore)) * Divisore
	' 			'g = CInt((g \ Divisore)) * Divisore
	' 
	' 			If r > 128 Then r = 65 Else r = 32
	' 			If g > 128 Then g = 65 Else g = 32
	' 			If b > 128 Then b = 65 Else b = 32
	' 
	' 			'C(0) = r
	' 			'C(1) = b
	' 			'C(2) = g
	' 			'For Z = 0 To 2
	' 			'    For L = Z + 1 To 2
	' 			'        If C(Z) < C(L) Then
	' 			'            A = C(Z)
	' 			'            C(Z) = C(L)
	' 			'            C(L) = A
	' 			'        End If
	' 			'    Next L
	' 			'Next Z
	' 			'r = C(0)
	' 
	' 			Select Case Chr(r) & Chr(g) & Chr(b)
	' 				Case "A  "
	' 					Valore = "1"
	' 				Case " A "
	' 					Valore = "2"
	' 				Case "  A"
	' 					Valore = "3"
	' 				Case "AA "
	' 					Valore = "4"
	' 				Case "A A"
	' 					Valore = "5"
	' 				Case " AA"
	' 					Valore = "6"
	' 				Case "AAA"
	' 					Valore = "7"
	' 				Case "   "
	' 					Valore = "8"
	' 				Case Else
	' 					Valore = "9"
	' 			End Select
	' 			Stringona += Valore
	' 		Next k
	' 	Next I
	' 
	' 	Stringona = Stringona.Replace(" '", "''")
	' 	Stringona = Stringona.Replace(Chr(0), "0")
	' 
	' 	'Dim Numerone As Long = 0
	' 
	' 	'For i As Integer = 0 To Stringona.Length - 1
	' 	'    Numerone += (Val(Stringona.Substring(i, 1))) * (i + 1)
	' 	'Next
	' 
	' 	imgImmagine.Dispose()
	' 	imgImmagine = Nothing
	' 
	' 	Try
	' 		Kill(NomeBNRid)
	' 	Catch ex As Exception
	' 
	' 	End Try
	' 
	' 	Dim Sql As String
	' 
	' 	Sql = "Delete * From CRC Where idImmagine=" & idImmagine
	' 	Db.EsegueSql(Conn, Sql)
	' 
	' 	Sql = "Insert Into CRC Values (" & idImmagine & ", '" & Stringona & "')"
	' 	Db.EsegueSql(Conn, Sql)
	' End Sub

	Public Sub SalvaImmagineDaPictureBox(filename As String, image As Image, Optional dimeX As Integer = -1, Optional dimeY As Integer = -1)
		If image Is Nothing Then
			Exit Sub
		End If

		If dimeX = -1 Or dimeY = -1 Then
			dimeX = image.Width
			dimeY = image.Height
		End If

		Dim bmp As Bitmap = image
		Dim bmpt As New Bitmap(dimeX, dimeY)
		Using g As Graphics = Graphics.FromImage(bmpt)
			g.DrawImage(bmp, 0, 0,
						bmpt.Width + 1,
						bmpt.Height + 1)
		End Using
		bmpt.Save(filename, System.Drawing.Imaging.ImageFormat.Jpeg)
	End Sub

	Public Function RitornaDimensioneImmagine(Immagine As String) As String
		If File.Exists(Immagine) = True Then
			Dim bt As Bitmap

			Try
				bt = LoadBitmapSenzaLock(Immagine) ' Image.FromFile(Immagine)
				Dim w As Integer = bt.Width
				Dim h As Integer = bt.Height

				bt.Dispose()
				bt = Nothing

				Return w & "x" & h
			Catch ex As Exception
				Return "ERRORE: " & ex.Message
			End Try
		Else
			Return "ERRORE: Immagine inesistente"
		End If
	End Function

	Public Function ConverteImmaginInBN(Path As String, Path2 As String, nomeLog As String, Optional RendeTrasparente As Boolean = True) As String
		Dim Ritorno As String = ""
		Dim img As Bitmap
		Dim ImmaginePiccola As Image
		'Dim ImmaginePiccola2 As Image
		Dim jgpEncoder As Imaging.ImageCodecInfo
		Dim myEncoder As System.Drawing.Imaging.Encoder
		Dim myEncoderParameters As New Imaging.EncoderParameters(1)

		Try
			ScriveLogGlobale(nomeLog, "Conversione in BN: Caricamento immagine")

			' img = New Bitmap(Path)
			img = LoadBitmapSenzaLock(Path)

			ImmaginePiccola = New Bitmap(img)

			img.Dispose()
			img = Nothing

			ImmaginePiccola = Converte(ImmaginePiccola)

			'ScriveLogGlobale(nomeLog, "Conversione in BN: Fase 1")
			jgpEncoder = GetEncoder(Imaging.ImageFormat.Png)
			myEncoder = System.Drawing.Imaging.Encoder.Quality
			'ScriveLogGlobale(nomeLog, "Conversione in BN: Fase 2")
			Dim myEncoderParameter As New Imaging.EncoderParameter(myEncoder, 99)
			myEncoderParameters.Param(0) = myEncoderParameter

			'ScriveLogGlobale(nomeLog, "Conversione in BN: Fase 3-1")
			Dim b As Bitmap = New Bitmap(ImmaginePiccola)
			If RendeTrasparente = True Then
				'ScriveLogGlobale(nomeLog, "Conversione in BN: Fase 3-2: " & b.Width & " x " & b.Height)
				b.MakeTransparent(Color.Black)
				''ScriveLogGlobale(nomeLog, "Conversione in BN: Fase 4")
			End If

			b.Save(Path2, jgpEncoder, myEncoderParameters)

			ImmaginePiccola.Dispose()
			b.Dispose()
			'ScriveLogGlobale(nomeLog, "Conversione in BN: Fase 5")

			ImmaginePiccola = Nothing
			'ImmaginePiccola2 = Nothing
			jgpEncoder = Nothing
			myEncoderParameter = Nothing

			Ritorno = "*"
		Catch ex As Exception
			Ritorno = "ERROR: conversione BN " & ex.Message
			ScriveLogGlobale(nomeLog, Ritorno)
		End Try

		Return Ritorno
	End Function

	Public Sub MetteCorniceAImmagine(Immagine As String, Destinazione As String)
		Try
			Dim bm As Bitmap
			Dim originalX As Integer
			Dim originalY As Integer

			bm = New Bitmap(Immagine)

			originalX = bm.Width
			originalY = bm.Height

			Dim thumb As New Bitmap(originalX, originalY)
			Dim g As Graphics = Graphics.FromImage(thumb)

			g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
			g.DrawImage(bm, New Rectangle(0, 0, originalX, originalY), New Rectangle(0, 0, bm.Width, bm.Height), GraphicsUnit.Pixel)

			Dim r As System.Drawing.Rectangle
			Dim Colore As Pen = Pens.White
			Dim ColoreNero As Pen = Pens.Gray
			Dim c As Integer = 0

			r.X = 0
			r.Y = 0
			r.Width = originalX - 1
			r.Height = originalY - 1

			g.DrawRectangle(ColoreNero, r)

			r.X = 1
			r.Y = 1
			r.Width = originalX - 2
			r.Height = originalY - 2

			g.DrawRectangle(ColoreNero, r)

			r.X = 11
			r.Y = 11
			r.Width = originalX - 23
			r.Height = originalY - 23

			g.DrawRectangle(ColoreNero, r)

			r.X = 12
			r.Y = 12
			r.Width = originalX - 25
			r.Height = originalY - 25

			g.DrawRectangle(ColoreNero, r)

			For i As Integer = 3 To 10
				r.X = i
				r.Y = i
				r.Width = originalX - i - (r.X + 1)
				r.Height = originalY - i - (r.Y + 1)

				g.DrawRectangle(Colore, r)
			Next

			'Colore = Pens.Gray

			'r.X = 9
			'r.Y = 9
			'r.Width = originalX - 9 - 1 - r.X
			'r.Height = originalY - 9 - 1 - r.Y

			'g.DrawRectangle(Colore, r)

			thumb.Save(Destinazione, System.Drawing.Imaging.ImageFormat.Jpeg)

			g.Dispose()

			bm.Dispose()
			thumb.Dispose()
		Catch ex As Exception

		End Try
	End Sub

	Public Function Ridimensiona(Path As String, Path2 As String, Larghezza As Integer, Altezza As Integer, Optional Format As ImageFormat = Nothing) As String
		Try
			Dim myEncoder As System.Drawing.Imaging.Encoder
			Dim myEncoderParameters As New Imaging.EncoderParameters(1)
			Dim img2 As Bitmap
			Dim ImmaginePiccola22 As Image
			Dim jgpEncoder2 As Imaging.ImageCodecInfo
			Dim myEncoder2 As System.Drawing.Imaging.Encoder
			Dim myEncoderParameters2 As New Imaging.EncoderParameters(1)

			img2 = New Bitmap(Path)
			ImmaginePiccola22 = New Bitmap(img2, Val(Larghezza), Val(Altezza))
			img2.Dispose()
			img2 = Nothing

			myEncoder = System.Drawing.Imaging.Encoder.Quality
			If Format Is Nothing Then
				jgpEncoder2 = GetEncoder(Imaging.ImageFormat.Jpeg)
			Else
				jgpEncoder2 = GetEncoder(Format)
			End If
			myEncoder2 = System.Drawing.Imaging.Encoder.Quality
			Dim myEncoderParameter2 As New Imaging.EncoderParameter(myEncoder, 97)
			myEncoderParameters2.Param(0) = myEncoderParameter2
			ImmaginePiccola22.Save(Path2, jgpEncoder2, myEncoderParameters2)

			ImmaginePiccola22.Dispose()

			ImmaginePiccola22 = Nothing
			jgpEncoder2 = Nothing
			myEncoderParameter2 = Nothing
		Catch ex As Exception
			Return "ERROR: " & ex.Message
		End Try

		Return "OK"
	End Function

	Private Function Converte(ByVal inputImage As Image) As Image
		Dim outputBitmap As Bitmap = New Bitmap(inputImage.Width, inputImage.Height)
		Dim X As Long
		Dim Y As Long
		Dim currentBWColor As Color

		For X = 0 To outputBitmap.Width - 1
			For Y = 0 To outputBitmap.Height - 1
				currentBWColor = ConverteColore(DirectCast(inputImage, Bitmap).GetPixel(X, Y))
				outputBitmap.SetPixel(X, Y, currentBWColor)
			Next
		Next

		inputImage = Nothing
		Return outputBitmap
	End Function

	Private Function ConverteColore(ByVal InputColor As Color)
		'Dim eyeGrayScale As Integer = (InputColor.R * 0.3 + InputColor.G * 0.59 + InputColor.B * 0.11)
		Dim Rosso As Single = InputColor.R * 0.3
		Dim Verde As Single = InputColor.G * 0.59
		Dim Blu As Single = InputColor.B * 0.41
		Dim eyeGrayScale As Integer = (Rosso + Verde + Blu) ' * 1.7
		If eyeGrayScale > 255 Then eyeGrayScale = 255
		Dim outputColor As Color = Color.FromArgb(eyeGrayScale, eyeGrayScale, eyeGrayScale)

		Return outputColor
	End Function

	Private Function ConverteChiara(ByVal inputImage As Image) As Image
		Dim outputBitmap As Bitmap = New Bitmap(inputImage.Width, inputImage.Height)
		Dim X As Long
		Dim Y As Long
		Dim currentBWColor As Color

		For X = 0 To outputBitmap.Width - 1
			For Y = 0 To outputBitmap.Height - 1
				currentBWColor = ConverteColoreChiaro(DirectCast(inputImage, Bitmap).GetPixel(X, Y))
				outputBitmap.SetPixel(X, Y, currentBWColor)
			Next
		Next

		inputImage = Nothing
		Return outputBitmap
	End Function

	Private Function ConverteColoreChiaro(ByVal InputColor As Color)
		'Dim eyeGrayScale As Integer = (InputColor.R * 0.3 + InputColor.G * 0.59 + InputColor.B * 0.11)
		Dim Rosso As Single = InputColor.R * 0.49999999999999994
		Dim Verde As Single = InputColor.G * 0.49000000000000005
		Dim Blu As Single = InputColor.B * 0.49999999999999595
		Dim eyeGrayScale As Integer = (Rosso + Verde + Blu) '* 4.1000000000000005
		If eyeGrayScale > 250 Then eyeGrayScale = 250
		If eyeGrayScale < 185 Then eyeGrayScale = 185
		Dim outputColor As Color = Color.FromArgb(eyeGrayScale, eyeGrayScale, eyeGrayScale)

		Return outputColor
	End Function

	Private Function GetEncoder(ByVal format As Imaging.ImageFormat) As Imaging.ImageCodecInfo

		Dim codecs As Imaging.ImageCodecInfo() = Imaging.ImageCodecInfo.GetImageDecoders()

		Dim codec As Imaging.ImageCodecInfo
		For Each codec In codecs
			If codec.FormatID = format.Guid Then
				Return codec
			End If
		Next codec
		Return Nothing

	End Function

	Public Sub RidimensionaEArrotondaIcona(ByVal PercorsoImmagine As String)
		Dim bm As Bitmap
		Dim originalX As Integer
		Dim originalY As Integer

		'carica immagine originale
		bm = New Bitmap(PercorsoImmagine)

		originalX = bm.Width
		originalY = bm.Height

		Dim thumb As New Bitmap(originalX, originalY)
		Dim g As Graphics = Graphics.FromImage(thumb)

		g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
		g.DrawImage(bm, New Rectangle(0, 0, originalX, originalY), New Rectangle(0, 0, bm.Width, bm.Height), GraphicsUnit.Pixel)

		Dim r As System.Drawing.Rectangle
		Dim s As System.Drawing.Size
		Dim coloreRosso As Pen = New Pen(Color.Red)
		coloreRosso.Width = 3

		For dimeX = originalX - 15 To originalX * 2
			r.X = (originalX / 2) - (dimeX / 2)
			r.Y = (originalY / 2) - (dimeX / 2)
			s.Width = dimeX
			s.Height = dimeX
			r.Size = s
			g.DrawEllipse(coloreRosso, r)
		Next

		Dim InizioY As Integer = -1
		Dim InizioX As Integer = -1
		Dim FineY As Integer = -1
		Dim FineX As Integer = -1
		Dim pixelColor As Color

		For i As Integer = 1 To originalX - 1
			For k As Integer = 1 To originalY - 1
				pixelColor = thumb.GetPixel(i, k)
				If pixelColor.ToArgb <> Color.Red.ToArgb Then
					InizioX = i
					'g.DrawLine(Pens.Black, i, 0, i, originalY)
					Exit For
				End If
			Next
			If InizioX <> -1 Then
				Exit For
			End If
		Next

		For i As Integer = originalX - 1 To 1 Step -1
			For k As Integer = originalY - 1 To 1 Step -1
				pixelColor = thumb.GetPixel(i, k)
				If pixelColor.ToArgb <> Color.Red.ToArgb Then
					FineX = i
					'g.DrawLine(Pens.Black, i, 0, i, originalY)
					Exit For
				End If
			Next
			If FineX <> -1 Then
				Exit For
			End If
		Next

		For i As Integer = 1 To originalY - 1
			For k As Integer = 1 To originalX - 1
				pixelColor = thumb.GetPixel(k, i)
				If pixelColor.ToArgb <> Color.Red.ToArgb Then
					InizioY = i
					'g.DrawLine(Pens.Black, 0, i, originalX, i)
					Exit For
				End If
			Next
			If InizioY <> -1 Then
				Exit For
			End If
		Next

		For i As Integer = originalY - 1 To 1 Step -1
			For k As Integer = originalX - 1 To 1 Step -1
				pixelColor = thumb.GetPixel(k, i)
				If pixelColor.ToArgb <> Color.Red.ToArgb Then
					FineY = i
					'g.DrawLine(Pens.Black, 0, i, originalX, i)
					Exit For
				End If
			Next
			If FineY <> -1 Then
				Exit For
			End If
		Next

		Dim nDimeX As Integer = FineX - InizioX
		Dim nDimeY As Integer = FineY - InizioY

		r.X = InizioX - 1
		r.Y = InizioY - 1
		r.Width = nDimeX + 1
		r.Height = nDimeY + 1

		Dim bmpAppoggio As Bitmap = New Bitmap(nDimeX, nDimeY)
		Dim g2 As Graphics = Graphics.FromImage(bmpAppoggio)

		g2.DrawImage(thumb, 0, 0, r, GraphicsUnit.Pixel)

		thumb = bmpAppoggio
		g2.Dispose()

		g.Dispose()

		thumb.MakeTransparent(Color.Red)

		thumb.Save(PercorsoImmagine & ".tsz", System.Drawing.Imaging.ImageFormat.Png)
		bm.Dispose()
		thumb.Dispose()

		Try
			Kill(PercorsoImmagine)
		Catch ex As Exception

		End Try

		Rename(PercorsoImmagine & ".tsz", PercorsoImmagine)
	End Sub

	Public Function RuotaFoto(Nome As String, Angolo As Single) As String
		Dim r As RotateFlipType

		Select Case Angolo
			Case 1
				r = RotateFlipType.RotateNoneFlipX
			Case 2
				r = RotateFlipType.RotateNoneFlipY
			Case 90
				r = RotateFlipType.Rotate90FlipNone
			Case -90
				r = RotateFlipType.Rotate270FlipNone
		End Select

		Dim bitmap1 As Bitmap = CType(Bitmap.FromFile(Nome), Bitmap)

		bitmap1.RotateFlip(r)
		bitmap1.Save(Nome & ".ruo", System.Drawing.Imaging.ImageFormat.Jpeg)
		bitmap1.Dispose()
		bitmap1 = Nothing

		Try
			Kill(Nome)

			Rename(Nome & ".ruo", Nome)

			Return "OK"
		Catch ex As Exception
			Return "ERRORE: " & ex.Message
		End Try
	End Function

	' Public Function RitornaDatiExif(Immagine As String) As String()
	' 	Dim imm As Bitmap = CaricaImmagineSenzaLockarla(Immagine)
	' 	Dim Campi() As String = {}
	' 
	' 	Try
	' 		Dim er As Goheer.EXIF.EXIFextractor = New Goheer.EXIF.EXIFextractor(imm, "§")
	' 		Campi = er.ToString.Split("§")
	' 		er = Nothing
	' 	Catch ex As Exception
	' 
	' 	End Try
	' 
	' 	imm.Dispose()
	' 	imm = Nothing
	' 
	' 	Return Campi
	' End Function

	Public Sub ConverteXRay(NomeBmp As String)
		Dim iX As Integer
		Dim iY As Integer

		Dim bmpImage As Bitmap = LoadBitmapSenzaLock(NomeBmp)
		Dim intPrevColor As Integer

		For iX = 0 To bmpImage.Width - 1
			For iY = 0 To bmpImage.Height - 1
				intPrevColor = (CInt(bmpImage.GetPixel(iX, iY).R) +
				   bmpImage.GetPixel(iX, iY).G +
				   bmpImage.GetPixel(iX, iY).B) \ 3

				bmpImage.SetPixel(iX, iY,
				   Color.FromArgb(intPrevColor, intPrevColor,
				   intPrevColor))
			Next iY
		Next iX

		Dim iaAttributes As New ImageAttributes
		Dim cmMatrix As New ColorMatrix

		cmMatrix.Matrix00 = -1
		cmMatrix.Matrix11 = -1
		cmMatrix.Matrix22 = -1
		cmMatrix.Matrix40 = 1
		cmMatrix.Matrix41 = 1
		cmMatrix.Matrix42 = 1

		iaAttributes.SetColorMatrix(cmMatrix)

		Dim rect As New Rectangle(0, 0, bmpImage.Width,
		   bmpImage.Height)

		Dim graph As Graphics = Graphics.FromImage(bmpImage)

		graph.DrawImage(bmpImage, rect, rect.X, rect.Y, rect.Width,
		   rect.Height, GraphicsUnit.Pixel, iaAttributes)

		RinominaBitmap(bmpImage, NomeBmp & ".XR", NomeBmp)
	End Sub

	Private Function PrendeIdDaTag(Tagghetto As String) As Integer
		Dim id As Integer = -1

		For i As Integer = 0 To QuantiTag - 1
			If Tagghetto.Replace(" ", "") = NomeTag(i) Then
				id = idTag(i)
				Exit For
			End If
		Next

		Return id
	End Function

	' ' Public Sub ScriveTag(NomeApplicazione As String, sNomeFile As String, NomeSito As String, Resto() As String)
	' 	Dim DatiExif() As String = RitornaDatiExif(sNomeFile)
	' 
	' 	Dim bmp As Bitmap = Image.FromFile(sNomeFile)
	' 
	' 	Dim er As Goheer.EXIF.EXIFextractor = New Goheer.EXIF.EXIFextractor(bmp, "\n")
	' 	Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Format(Now.Year, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
	' 
	' 	Dim testo As String = NomeSito & ";"
	' 	For i As Integer = 0 To Resto.Length - 2
	' 		testo += Resto(i) & ";"
	' 	Next
	' 
	' 	Dim nomeimm As String = Resto(Resto.Length - 1)
	' 	For i As Integer = nomeimm.Length To 1 Step -1
	' 		If Mid(nomeimm, i, 1) = "." Then
	' 			nomeimm = Mid(nomeimm, 1, i - 1)
	' 			Exit For
	' 		End If
	' 	Next
	' 
	' 	' imposta codici originali
	' 	Dim testina As String
	' 	Dim testone As String
	' 	Dim id As Integer
	' 	Dim ceCommento As Boolean = False
	' 
	' 	For i As Integer = 0 To DatiExif.Length - 1
	' 		If DatiExif(i) <> "" Then
	' 			testina = Mid(DatiExif(i), 1, DatiExif(i).IndexOf(":")).Trim.ToUpper
	' 			testone = Mid(DatiExif(i), DatiExif(i).IndexOf(":") + 2, DatiExif(i).Length).Trim
	' 			id = PrendeIdDaTag(testina)
	' 			If id <> -1 Then
	' 				If id = 270 Then
	' 					testone = testo & "§;" & testone & ";"
	' 					ceCommento = True
	' 				End If
	' 
	' 				er.setTag(id, testone & Chr(0))
	' 			End If
	' 		End If
	' 	Next
	' 	' imposta codici originali
	' 
	' 	If ceCommento = False Then
	' 		er.setTag(270, testo & Chr(0))
	' 	End If
	' 	er.setTag(305, NomeApplicazione & Chr(0))
	' 	er.setTag(306, Datella & Chr(0))
	' 
	' 	Try
	' 		bmp.Save(sNomeFile & ".bbb")
	' 	Catch ex As Exception
	' 		'Stop
	' 	End Try
	' 
	' 	er = Nothing
	' 	bmp.Dispose()
	' 	bmp = Nothing
	' 
	' 	File.Delete(sNomeFile)
	' 	Dim c As Integer = 0
	' 	Do While File.Exists(sNomeFile & ".bbb")
	' 		Rename(sNomeFile & ".bbb", sNomeFile)
	' 		Thread.Sleep(1000)
	' 		c += 1
	' 		If c = 5 Then
	' 			Exit Do
	' 		End If
	' 	Loop
	' End Sub

	Public Sub ConverteInSeppia(NomeBmp As String)
		Dim bmpImage As Bitmap = LoadBitmapSenzaLock(NomeBmp)
		Dim cCurrColor As Color

		For iY As Integer = 0 To bmpImage.Height - 1
			For iX As Integer = 0 To bmpImage.Width - 1
				cCurrColor = bmpImage.GetPixel(iX, iY)

				Dim intAlpha As Integer = cCurrColor.A
				Dim intRed As Integer = cCurrColor.R
				Dim intGreen As Integer = cCurrColor.G
				Dim intBlue As Integer = cCurrColor.B

				Dim intSRed As Integer = CInt((0.393 * intRed +
				   0.769 * intGreen + 0.189 * intBlue))
				Dim intSGreen As Integer = CInt((0.349 * intRed +
				   0.686 * intGreen + 0.168 * intBlue))
				Dim intSBlue As Integer = CInt((0.272 * intRed +
				   0.534 * intGreen + 0.131 * intBlue))

				If intSRed > 255 Then
					intRed = 255
				Else
					intRed = intSRed
				End If

				If intSGreen > 255 Then
					intGreen = 255
				Else
					intGreen = intSGreen
				End If

				If intSBlue > 255 Then
					intBlue = 255
				Else
					intBlue = intSBlue
				End If

				bmpImage.SetPixel(iX, iY, Color.FromArgb(intAlpha,
				   intRed, intGreen, intBlue))
			Next

		Next

		RinominaBitmap(bmpImage, NomeBmp & ".SEP", NomeBmp)
	End Sub

	Public Sub ConverteEdge(NomeBmp As String)
		Dim tmpImage As Bitmap = LoadBitmapSenzaLock(NomeBmp)
		Dim bmpImage As Bitmap = LoadBitmapSenzaLock(NomeBmp)

		Dim intWidth As Integer = tmpImage.Width
		Dim intHeight As Integer = tmpImage.Height

		Dim intOldX As Integer(,) = New Integer(,) {{-1, 0, 1},
		   {-2, 0, 2}, {-1, 0, 1}}
		Dim intOldY As Integer(,) = New Integer(,) {{1, 2, 1},
		   {0, 0, 0}, {-1, -2, -1}}

		Dim intR As Integer(,) = New Integer(intWidth - 1,
		   intHeight - 1) {}
		Dim intG As Integer(,) = New Integer(intWidth - 1,
		   intHeight - 1) {}
		Dim intB As Integer(,) = New Integer(intWidth - 1,
		   intHeight - 1) {}

		Dim intMax As Integer = 16 * 16

		For i As Integer = 0 To intWidth - 1
			For j As Integer = 0 To intHeight - 1
				intR(i, j) = tmpImage.GetPixel(i, j).R
				intG(i, j) = tmpImage.GetPixel(i, j).G
				intB(i, j) = tmpImage.GetPixel(i, j).B
			Next
		Next

		Dim intRX As Integer = 0
		Dim intRY As Integer = 0
		Dim intGX As Integer = 0
		Dim intGY As Integer = 0
		Dim intBX As Integer = 0
		Dim intBY As Integer = 0

		Dim intRTot As Integer
		Dim intGTot As Integer
		Dim intBTot As Integer

		For i As Integer = 1 To tmpImage.Width - 1 - 1
			For j As Integer = 1 To tmpImage.Height - 1 - 1
				intRX = 0
				intRY = 0
				intGX = 0
				intGY = 0
				intBX = 0
				intBY = 0

				intRTot = 0
				intGTot = 0
				intBTot = 0

				For width As Integer = -1 To 2 - 1
					For height As Integer = -1 To 2 - 1
						intRTot = intR(i + height, j + width)
						intRX += intOldX(width + 1, height + 1) * intRTot
						intRY += intOldY(width + 1, height + 1) * intRTot

						intGTot = intG(i + height, j + width)
						intGX += intOldX(width + 1, height + 1) * intGTot
						intGY += intOldY(width + 1, height + 1) * intGTot

						intBTot = intB(i + height, j + width)
						intBX += intOldX(width + 1, height + 1) * intBTot
						intBY += intOldY(width + 1, height + 1) * intBTot
					Next
				Next

				If intRX * intRX + intRY * intRY > intMax OrElse
				   intGX * intGX + intGY * intGY > intMax OrElse
				   intBX * intBX + intBY * intBY > intMax Then

					bmpImage.SetPixel(i, j, Color.Black)
				Else
					bmpImage.SetPixel(i, j, Color.Transparent)
				End If
			Next
		Next

		RinominaBitmap(bmpImage, NomeBmp & ".EDGE", NomeBmp)
	End Sub

	Public Sub RotateImage(NomeBmp As String, angle As Single)
		Dim bmp As Bitmap = LoadBitmapSenzaLock(NomeBmp)

		Dim Height As Single = bmp.Height
		Dim Width As Single = bmp.Width
		Dim hypotenuse As Integer = System.Convert.ToInt32(System.Math.Floor(Math.Sqrt(Height * Height + Width * Width)))
		Dim rotatedImage As Bitmap = New Bitmap(hypotenuse, hypotenuse)
		Using g As Graphics = Graphics.FromImage(rotatedImage)
			g.TranslateTransform(rotatedImage.Width / 2, rotatedImage.Height / 2)
			g.RotateTransform(angle)
			g.TranslateTransform(-rotatedImage.Width / 2, -rotatedImage.Height / 2)
			g.DrawImage(bmp, (hypotenuse - Width) / 2, (hypotenuse - Height) / 2, Width, Height)
			rotatedImage.MakeTransparent(Color.Black)
		End Using

		RinominaBitmap(rotatedImage, NomeBmp & ".ROT", NomeBmp)
	End Sub

	Public Function RuotaImmagineSenzaSalvare(bm_in As Bitmap, angolo As Single) As Bitmap
		' Copy the output bitmap from the source image.
		' Make an array of points defining the
		' image's corners.
		Dim wid As Single = bm_in.Width
		Dim hgt As Single = bm_in.Height
		Dim corners As Point() = {
			New Point(0, 0),
			New Point(wid, 0),
			New Point(0, hgt),
			New Point(wid, hgt)}

		' Translate to center the bounding box at the origin.
		Dim cx As Single = wid / 2
		Dim cy As Single = hgt / 2
		Dim i As Long
		For i = 0 To 3
			corners(i).X -= cx
			corners(i).Y -= cy
		Next i

		' Rotate.
		Dim theta As Single = angolo * Math.PI _
			/ 180.0
		Dim sin_theta As Single = Math.Sin(theta)
		Dim cos_theta As Single = Math.Cos(theta)
		Dim X As Single
		Dim Y As Single
		For i = 0 To 3
			X = corners(i).X
			Y = corners(i).Y
			corners(i).X = X * cos_theta + Y * sin_theta
			corners(i).Y = -X * sin_theta + Y * cos_theta
		Next i

		' Translate so X >= 0 and Y >=0 for all corners.
		Dim xmin As Single = corners(0).X
		Dim ymin As Single = corners(0).Y
		For i = 1 To 3
			If xmin > corners(i).X Then xmin = corners(i).X
			If ymin > corners(i).Y Then ymin = corners(i).Y
		Next i
		For i = 0 To 3
			corners(i).X -= xmin
			corners(i).Y -= ymin
		Next i

		' Create an output Bitmap and Graphics object.
		Dim bm_out As New Bitmap(CInt(-2 * xmin), CInt(-2 *
			ymin))
		Dim gr_out As Graphics = Graphics.FromImage(bm_out)
		' gr_out.Clear(Color.Azure)

		' Drop the last corner lest we confuse DrawImage, 
		' which expects an array of three corners.
		ReDim Preserve corners(2)

		' Draw the result onto the output Bitmap.
		gr_out.DrawImage(bm_in, corners)
		' bm_out.MakeTransparent(Color.Azure)

		' Display the result.
		' RinominaBitmap(bm_out, NomeBmp & ".RUO", NomeBmp)
		Return bm_out
	End Function

	Public Sub RuotaImmagine(NomeBmp As String, angolo As Single)
		Dim bm_in As Bitmap = LoadBitmapSenzaLock(NomeBmp)

		' Copy the output bitmap from the source image.
		' Make an array of points defining the
		' image's corners.
		Dim wid As Single = bm_in.Width
		Dim hgt As Single = bm_in.Height
		Dim corners As Point() = {
			New Point(0, 0),
			New Point(wid, 0),
			New Point(0, hgt),
			New Point(wid, hgt)}

		' Translate to center the bounding box at the origin.
		Dim cx As Single = wid / 2
		Dim cy As Single = hgt / 2
		Dim i As Long
		For i = 0 To 3
			corners(i).X -= cx
			corners(i).Y -= cy
		Next i

		' Rotate.
		Dim theta As Single = angolo * Math.PI _
			/ 180.0
		Dim sin_theta As Single = Math.Sin(theta)
		Dim cos_theta As Single = Math.Cos(theta)
		Dim X As Single
		Dim Y As Single
		For i = 0 To 3
			X = corners(i).X
			Y = corners(i).Y
			corners(i).X = X * cos_theta + Y * sin_theta
			corners(i).Y = -X * sin_theta + Y * cos_theta
		Next i

		' Translate so X >= 0 and Y >=0 for all corners.
		Dim xmin As Single = corners(0).X
		Dim ymin As Single = corners(0).Y
		For i = 1 To 3
			If xmin > corners(i).X Then xmin = corners(i).X
			If ymin > corners(i).Y Then ymin = corners(i).Y
		Next i
		For i = 0 To 3
			corners(i).X -= xmin
			corners(i).Y -= ymin
		Next i

		' Create an output Bitmap and Graphics object.
		Dim bm_out As New Bitmap(CInt(-2 * xmin), CInt(-2 *
			ymin))
		Dim gr_out As Graphics = Graphics.FromImage(bm_out)
		gr_out.Clear(Color.Azure)

		' Drop the last corner lest we confuse DrawImage, 
		' which expects an array of three corners.
		ReDim Preserve corners(2)

		' Draw the result onto the output Bitmap.
		gr_out.DrawImage(bm_in, corners)
		bm_out.MakeTransparent(Color.Azure)

		File.Delete(NomeBmp)
		bm_out.Save(NomeBmp)

		' Display the result.
		' RinominaBitmap(bm_out, NomeBmp & ".RUO", NomeBmp)
	End Sub

	Public Function LoadBitmapSenzaLock(NomeBmp As String) As Bitmap
		Dim bmp As Bitmap

		Using fs As New FileStream(NomeBmp, FileMode.Open, FileAccess.Read)
			bmp = Image.FromStream(fs)
		End Using

		Return bmp
	End Function

	Public Sub RinominaBitmap(bmp As Bitmap, Nome1 As String, Nome2 As String)
		Dim c As Integer = 0
		Dim Ok As Boolean = True

		Do While File.Exists(Nome1)
			File.Delete(Nome1)
			Thread.Sleep(1000)
			c += 1
			If c = 10 Then
				Ok = False
				Exit Do
			End If
		Loop

		bmp.MakeTransparent(Color.Black)
		If Ok Then
			Try
				bmp.Save(Nome1, ImageFormat.Png)
			Catch ex As Exception
				Stop
			End Try
		End If

		bmp.Dispose()
		bmp = Nothing

		If Ok Then
			File.Delete(Nome2)
			Rename(Nome1, Nome2)
		End If
	End Sub

	Public Sub ApplicaOmbraABitmap(ByRef SourceImage As Drawing.Bitmap,
							ByVal ShadowColor As Drawing.Color,
							ByVal BackgroundColor As Drawing.Color,
							Optional ByVal ShadowDirection As ShadowDirections =
												  ShadowDirections.BOTTOM_RIGHT,
							Optional ByVal ShadowOpacity As Integer = 190,
							Optional ByVal ShadowSoftness As Integer = 4,
							Optional ByVal ShadowDistance As Integer = 5,
							Optional ByVal ShadowRoundedEdges As Boolean = True)
		Dim ImgTarget As Bitmap = Nothing
		Dim ImgShadow As Bitmap = Nothing
		Dim g As Graphics = Nothing
		Try
			If SourceImage IsNot Nothing Then
				If ShadowOpacity < 0 Then
					ShadowOpacity = 0
				ElseIf ShadowOpacity > 255 Then
					ShadowOpacity = 255
				End If
				If ShadowSoftness < 1 Then
					ShadowSoftness = 1
				ElseIf ShadowSoftness > 30 Then
					ShadowSoftness = 30
				End If
				If ShadowDistance < 1 Then
					ShadowDistance = 1
				ElseIf ShadowDistance > 50 Then
					ShadowDistance = 50
				End If
				If ShadowColor = Color.Transparent Then
					ShadowColor = Color.Black
				End If
				If BackgroundColor = Color.Transparent Then
					BackgroundColor = Color.White
				End If

				'get shadow
				Dim shWidth As Integer = CInt(SourceImage.Width / ShadowSoftness)
				Dim shHeight As Integer = CInt(SourceImage.Height / ShadowSoftness)
				ImgShadow = New Bitmap(shWidth, shHeight)
				g = Graphics.FromImage(ImgShadow)
				g.Clear(Color.Transparent)
				g.InterpolationMode = InterpolationMode.HighQualityBicubic
				g.SmoothingMode = SmoothingMode.AntiAlias
				Dim sre As Integer = 0
				If ShadowRoundedEdges = True Then sre = 1
				g.FillRectangle(New SolidBrush(Color.FromArgb(ShadowOpacity, ShadowColor)),
									  sre, sre, shWidth, shHeight)
				g.Dispose()

				'draw shadow
				Dim d_shWidth As Integer = SourceImage.Width + ShadowDistance
				Dim d_shHeight As Integer = SourceImage.Height + ShadowDistance
				ImgTarget = New Bitmap(d_shWidth, d_shHeight)
				g = Graphics.FromImage(ImgTarget)
				g.Clear(BackgroundColor)
				g.InterpolationMode = InterpolationMode.HighQualityBicubic
				g.SmoothingMode = SmoothingMode.AntiAlias
				g.DrawImage(ImgShadow, New Rectangle(0, 0, d_shWidth, d_shHeight),
										0, 0, ImgShadow.Width, ImgShadow.Height, GraphicsUnit.Pixel)
				Select Case ShadowDirection
					Case ShadowDirections.BOTTOM_RIGHT
						g.DrawImage(SourceImage,
							New Rectangle(0, 0, SourceImage.Width, SourceImage.Height),
							   0, 0, SourceImage.Width, SourceImage.Height, GraphicsUnit.Pixel)
					Case ShadowDirections.BOTTOM_LEFT
						g.Dispose()
						ImgTarget.RotateFlip(RotateFlipType.RotateNoneFlipX)
						g = Graphics.FromImage(ImgTarget)
						g.DrawImage(SourceImage,
							 New Rectangle(ShadowDistance, 0, SourceImage.Width, SourceImage.Height),
									0, 0, SourceImage.Width, SourceImage.Height, GraphicsUnit.Pixel)
					Case ShadowDirections.TOP_LEFT
						g.Dispose()
						ImgTarget.RotateFlip(RotateFlipType.Rotate180FlipNone)
						g = Graphics.FromImage(ImgTarget)
						g.DrawImage(SourceImage,
									  New Rectangle(ShadowDistance, ShadowDistance,
														SourceImage.Width, SourceImage.Height),
									  0, 0, SourceImage.Width, SourceImage.Height, GraphicsUnit.Pixel)
					Case ShadowDirections.TOP_RIGHT
						g.Dispose()
						ImgTarget.RotateFlip(RotateFlipType.RotateNoneFlipY)
						g = Graphics.FromImage(ImgTarget)
						g.DrawImage(SourceImage,
						   New Rectangle(0, ShadowDistance, SourceImage.Width, SourceImage.Height),
								  0, 0, SourceImage.Width, SourceImage.Height, GraphicsUnit.Pixel)
				End Select

				g.Dispose()
				g = Nothing
				ImgShadow.Dispose()
				ImgShadow = Nothing

				SourceImage = New Bitmap(ImgTarget)
				ImgTarget.Dispose()
				ImgTarget = Nothing
			End If

		Catch ex As Exception
			If g IsNot Nothing Then
				g.Dispose()
				g = Nothing
			End If
			If ImgShadow IsNot Nothing Then
				ImgShadow.Dispose()
				ImgShadow = Nothing
			End If
			If ImgTarget IsNot Nothing Then
				ImgTarget.Dispose()
				ImgTarget = Nothing
			End If
		End Try

	End Sub

	Private Function ContaPunti(MP As String, filename As String, filelog As String, tutto As Boolean) As String
		Dim y As Long = -1

		Static xx As Random = New Random()

		y = xx.Next(1000)

		Dim NomeEsteso As String = Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & "_" & y.ToString.Trim & "_Ptn"
		Dim Punti As String = ""
		Dim PuntiDiagonale As String = ""
		Dim PuntiCornice As String = ""
		Dim Hash As String = ""
		Dim HashColore As String = ""
		'Dim Nome As String = MP & "/Appoggio/RidBN_" & NomeEsteso & ".jpg"

		ScriveLogGlobale(filelog, "Conteggio punti: " & filename)

		Dim NomeFileRidotto As String = MP & "/Appoggio/Rid_" & NomeEsteso & ".jpg"
		Dim rit As String = RidimensionaMantenendoProporzioni(filename, NomeFileRidotto, 100, False)

		Dim NomeFileRidottoBN As String = MP & "/Appoggio/RidBN_" & NomeEsteso & ".jpg"
		rit = ConverteImmaginInBN(NomeFileRidotto, NomeFileRidottoBN, filelog, False)
		'Dim NomeFileRidottoBN2 As String = MP & "/Appoggio/RidBN2_" & NomeEsteso & ".jpg"

		Try
			Dim img As Bitmap = LoadBitmapSenzaLock(NomeFileRidotto)
			Dim width2 As Integer = img.Width
			Dim height2 As Integer = img.Height
			Dim DimensioneDaControllare As Integer = 50 ' width2
			Dim imgBN As Bitmap = LoadBitmapSenzaLock(NomeFileRidottoBN)

			Dim quantoX As Integer = width2 + 1
			Dim quantoY As Integer = height2 + 1
			'If tutto = False Then
			'	quantoX = DimensioneDaControllare + 1
			'	quantoY = DimensioneDaControllare + 1
			'End If
			'Dim img2 As Bitmap = New Bitmap(quantoX, quantoY) ' DimensioneDaControllare + 1, DimensioneDaControllare + 1)
			'Dim img3 As Bitmap = New Bitmap(DimensioneDaControllare + 1, DimensioneDaControllare + 1)
			'Dim img4 As Bitmap = New Bitmap(DimensioneDaControllare + 1, DimensioneDaControllare + 1)
			ScriveLogGlobale(filelog, "DimeX-DimeY: " & quantoX & "-" & quantoY)

			'Dim PuntiNeg As Integer = 0
			'Dim PuntiNeg2 As Integer = 0

			Dim inizioX As Integer = ((width2 - 1) / 2) - (DimensioneDaControllare / 2)
			Dim fineX As Integer = ((width2 - 1) / 2) + (DimensioneDaControllare / 2)
			Dim inizioY As Integer = ((height2 - 1) / 2) - (DimensioneDaControllare / 2)
			Dim fineY As Integer = ((height2 - 1) / 2) + (DimensioneDaControllare / 2)
			'Dim inizioX As Integer = 0
			'Dim fineX As Integer = width2 - 1
			'Dim inizioY As Integer = 0
			'Dim fineY As Integer = height2 - 1
			Dim posX As Integer = 0
			Dim posY As Integer = 0

			If tutto = True Then
				inizioX = 0
				fineX = width2 - 1
				inizioY = 0
				fineY = height2 - 1
			End If
			ScriveLogGlobale(filelog, "RangeX: " & inizioX & "-" & fineX)
			ScriveLogGlobale(filelog, "RangeY: " & inizioY & "-" & fineY)

			'Dim listaPunti As New List(Of Color)
			'Dim quantiPunti As New List(Of Integer)
			'Dim listaPuntiArr As New List(Of String)
			'Dim quantiPuntiArr As New List(Of Integer)
			'Dim Colore As New List(Of Color)
			Dim QQ1 As Integer = 0
			Dim QQ2 As Integer = 0
			Dim QQ3 As Integer = 0
			Dim QQBn As Integer = 0
			Dim QQTot As Integer = 0

			For i As Integer = inizioX To fineX
				posY = 0
				For k As Integer = inizioY To fineY
					Dim pixelColor As Color = img.GetPixel(i, k)

					Dim Alfa As Integer = pixelColor.A ' Int(Val(pixelColor.R.ToString()) / 10) * 10
					Dim Red As Integer = pixelColor.R ' Int(Val(pixelColor.R.ToString()) / 10) * 10
					Dim Green As Integer = pixelColor.G ' Int(Val(pixelColor.G.ToString()) / 10) * 10
					Dim Blue As Integer = pixelColor.B '  Int(Val(pixelColor.B.ToString()) / 10) * 10
					'Dim Colorone As String = Red & "-" & Green & "-" & Blue

					'Dim Alfa2 As Integer = (CInt(Alfa / 10) * 10) ' + ((i + k) * 25)
					'Dim Red2 As Integer = (CInt(Red / 10) * 10) ' + ((i + k) * 25)
					'Dim Green2 As Integer = (CInt(Green / 10) * 10) ' + ((i + k) * 25)
					'Dim Blue2 As Integer = (CInt(Blue / 10) * 10) ' + ((i + k) * 25)

					QQ1 += IIf(Red > 128, 1, 0)
					QQ2 += IIf(Green > 128, 1, 0)
					QQ3 += IIf(Blue > 128, 1, 0)

					'If Red2 > 230 Then
					'	QQ1 += 15
					'Else
					'	If Red2 > 200 Then
					'		QQ1 += 11
					'	Else
					'		If Red2 > 128 Then
					'			QQ1 += 6
					'		Else
					'			If Red2 > 60 Then
					'				QQ1 += 2
					'			Else
					'				If Red2 > 30 Then
					'					QQ1 += 1
					'				End If
					'			End If
					'		End If
					'	End If
					'End If

					'If Green2 > 230 Then
					'	QQ2 += 15
					'Else
					'	If Green2 > 200 Then
					'		QQ2 += 11
					'	Else
					'		If Green2 > 128 Then
					'			QQ2 += 6
					'		Else
					'			If Green2 > 60 Then
					'				QQ2 += 2
					'			Else
					'				If Green2 > 30 Then
					'					QQ2 += 1
					'				End If
					'			End If
					'		End If
					'	End If
					'End If

					'If Blue2 > 230 Then
					'	QQ3 += 15
					'Else
					'	If Blue2 > 200 Then
					'		QQ3 += 11
					'	Else
					'		If Blue2 > 128 Then
					'			QQ3 += 6
					'		Else
					'			If Blue2 > 60 Then
					'				QQ3 += 2
					'			Else
					'				If Blue2 > 30 Then
					'					QQ3 += 1
					'				End If
					'			End If
					'		End If
					'	End If
					'End If

					Dim pixelColor2 As Color = imgBN.GetPixel(i, k)
					Dim BN As Integer = pixelColor2.R ' CInt(pixelColor2.R / 10) * 10

					QQBn += IIf(BN > 128, 1, 0)

					'If BN > 230 Then
					'	QQBn += 15
					'Else
					'	If BN > 200 Then
					'		QQBn += 11
					'		QQTot += 1
					'	Else
					'		If BN > 128 Then
					'			QQBn += 6
					'		Else
					'			If BN > 60 Then
					'				QQBn += 2
					'			Else
					'				If BN > 30 Then
					'					QQBn += 1
					'				End If
					'			End If
					'		End If
					'	End If
					'End If

					'Dim Media As Integer = CInt(((CInt(Red / 100) * 100) + (CInt(Green / 100) * 100) + (CInt(Blue / 100) * 100) + BN) / 4)

					'Media = (CInt(Media / 100) * 100)
					'QQTot += Media

					'Dim Ok3 As Boolean = True
					'Dim Quale As Integer = 0

					'For Each v As Color In listaPunti
					'	If v = pixelColor Then
					'		Ok3 = False
					'		quantiPunti.Item(Quale) += 1
					'		Exit For
					'	End If
					'	Quale += 1
					'Next

					'If Ok3 Then
					'	listaPunti.Add(pixelColor)
					'	quantiPunti.Add(1)
					'	'Dim Colorello As Color = Color.FromArgb(Quale, Quale, Quale)
					'	'Colore.Add(Colorello)
					'	Quale = listaPunti.Count - 1
					'End If

					'Ok3 = True
					'Quale = 0

					'For Each v As String In listaPuntiArr
					'	If v = Colorone Then
					'		Ok3 = False
					'		quantiPuntiArr.Item(Quale) += 1
					'		Exit For
					'	End If
					'	Quale += 1
					'Next

					'If Ok3 Then
					'	listaPuntiArr.Add(Colorone)
					'	quantiPuntiArr.Add(1)
					'	Dim Colorello As Color = Color.FromArgb(Red, Green, Blue)
					'	Colore.Add(Colorello)
					'	Quale = listaPuntiArr.Count - 1
					'End If

					'img2.SetPixel(posX, posY, Colore.Item(Quale))
					' ScriveLogGlobale(fileLog, "Colori " & i & "," & k & ": R. " & pixelColor.R & " G. " & pixelColor.G & " B. " & pixelColor.B)

					'''If tutto = True Then
					'''	img2.SetPixel(posX, posY, pixelColor)

					'''	Punti += 10 ' CInt(System.Drawing.ColorTranslator.ToWin32(pixelColor) / 100000) ' Val(pixelColor.R) + Val(pixelColor.G) + Val(pixelColor.B)
					'''Else
					'''	'ScriveLogGlobale(fileLog, "Controllo Colori")
					'''	If pixelColor.R > valoreColore Or pixelColor.G > valoreColore Or pixelColor.B > valoreColore Then
					'''		'ScriveLogGlobale(fileLog, "Controllo Colori1")
					'''		img2.SetPixel(posX, posY, Bianco)
					'''		'img3.SetPixel(posX, posY, Nero)
					'''		'PuntiNeg += 1
					'''	Else
					'''		If pixelColor.R > (valoreColore / 1.5) Or pixelColor.G > (valoreColore / 1.5) Or pixelColor.B > (valoreColore / 1.5) Then
					'''			'ScriveLogGlobale(fileLog, "Controllo Colori2")
					'''			img2.SetPixel(posX, posY, GrigioChiaro)
					'''			'img3.SetPixel(posX, posY, Grigio)
					'''			Punti += 20 '(3 * (posX + posY)) ' Val(pixelColor.R) + Val(pixelColor.G) + Val(pixelColor.B)
					'''		Else
					'''			If pixelColor.R > (valoreColore / 2.5) Or pixelColor.G > (valoreColore / 2.5) Or pixelColor.B > (valoreColore / 2.5) Then
					'''				'ScriveLogGlobale(fileLog, "Controllo Colori3")
					'''				img2.SetPixel(posX, posY, Grigio)
					'''				'img3.SetPixel(posX, posY, GrigioChiaro)
					'''				Punti += 40 '(5 * (posX + posY)) ' Val(pixelColor.R) + Val(pixelColor.G) + Val(pixelColor.B)
					'''			Else
					'''				'ScriveLogGlobale(fileLog, "Controllo Colori4")
					'''				img2.SetPixel(posX, posY, Nero)
					'''				'img3.SetPixel(posX, posY, Bianco)
					'''				Punti += 80 '(7 * (posX + posY)) ' Val(pixelColor.R) + Val(pixelColor.G) + Val(pixelColor.B)
					'''			End If
					'''		End If
					'''	End If
					'''End If
					'''''ScriveLogGlobale(fileLog, "Controllo Essenziale")
					''''If (pixelColor.R >= valoreColore - 10 And pixelColor.R <= valoreColore + 10) And
					''''	(pixelColor.G >= valoreColore - 10 And pixelColor.G <= valoreColore + 210) And
					''''	(pixelColor.B >= valoreColore - 10 And pixelColor.B <= valoreColore + 10) Then
					''''	'img4.SetPixel(posX, posY, Nero)
					''''	PuntiNeg2 += 1
					''''End If
					''''img2.SetPixel(posX, posY, Nero)
					''''ScriveLogGlobale(filelog, "Fine Controllo")
					''''If pixelColor.R > valoreColore Or pixelColor.G > valoreColore Or pixelColor.B > valoreColore Then
					''''	img2.SetPixel(posX, posY, Nero)
					''''	Punti += 1
					''''End If
					posY += 1
				Next
				posX += 1
			Next

			'Dim Quale2 As Integer = 0

			'For i As Integer = 0 To listaPunti.Count - 1
			'	For K As Integer = i + 1 To listaPunti.Count - 1
			'		If quantiPunti.Item(i) < quantiPunti.Item(K) Then
			'			Dim Appoggio2 As Integer = quantiPunti.Item(i)
			'			quantiPunti.Item(i) = quantiPunti.Item(K)
			'			quantiPunti.Item(K) = Appoggio2
			'			Dim AppoggioC As Color = listaPunti.Item(i)
			'			listaPunti.Item(i) = listaPunti.Item(K)
			'			listaPunti.Item(K) = AppoggioC
			'		End If
			'	Next
			'Next
			'ScriveLogGlobale(filelog, "Quanti Pixel diversi: " & listaPunti.Count - 1)
			'Try
			'	ScriveLogGlobale(filelog, "Pixel 1: " & listaPunti.Item(0).ToString() & " -> " & quantiPunti.Item(0))
			'	ScriveLogGlobale(filelog, "Pixel 2: " & listaPunti.Item(1).ToString() & " -> " & quantiPunti.Item(1))
			'	ScriveLogGlobale(filelog, "Pixel 3: " & listaPunti.Item(2).ToString() & " -> " & quantiPunti.Item(2))
			'	ScriveLogGlobale(filelog, "Pixel 4: " & listaPunti.Item(3).ToString() & " -> " & quantiPunti.Item(3))
			'	ScriveLogGlobale(filelog, "Pixel 5: " & listaPunti.Item(4).ToString() & " -> " & quantiPunti.Item(4))
			'Catch ex As Exception

			'End Try

			QQ1 = CInt(QQ1 / 5) * 5
			QQ2 = CInt(QQ2 / 5) * 5
			QQ3 = CInt(QQ3 / 5) * 5
			QQBn = CInt(QQBn / 5) * 5
			QQTot = ((QQ1 + QQ2 + QQ3 + QQBn) / 4) * 0.7

			PuntiDiagonale = QQ2 '  listaPunti.Item(0).R.ToString() & "-" & listaPunti.Item(0).G.ToString() & "-" & listaPunti.Item(0).B.ToString() & "-" & quantiPunti.Item(0)
			Hash = QQBn ' QQ2 ' listaPunti.Item(1).R.ToString() & "-" & listaPunti.Item(1).G.ToString() & "-" & listaPunti.Item(1).B.ToString() & "-" & quantiPunti.Item(1)

			'For i As Integer = 0 To listaPuntiArr.Count - 1
			'	For K As Integer = i + 1 To listaPuntiArr.Count - 1
			'		If quantiPuntiArr.Item(i) < quantiPuntiArr.Item(K) Then
			'			Dim Appoggio2 As Integer = quantiPuntiArr.Item(i)
			'			quantiPuntiArr.Item(i) = quantiPuntiArr.Item(K)
			'			quantiPuntiArr.Item(K) = Appoggio2
			'			Dim AppoggioS As String = listaPuntiArr.Item(i)
			'			listaPuntiArr.Item(i) = listaPuntiArr.Item(K)
			'			listaPuntiArr.Item(K) = AppoggioS
			'		End If
			'	Next
			'Next
			'Try
			'	ScriveLogGlobale(filelog, "Quanti Pixel diversi arrotondati: " & listaPuntiArr.Count - 1)
			'	ScriveLogGlobale(filelog, "Pixel 1: " & listaPuntiArr.Item(0).ToString() & " -> " & quantiPuntiArr.Item(0))
			'	ScriveLogGlobale(filelog, "Pixel 2: " & listaPuntiArr.Item(1).ToString() & " -> " & quantiPuntiArr.Item(1))
			'	ScriveLogGlobale(filelog, "Pixel 3: " & listaPuntiArr.Item(2).ToString() & " -> " & quantiPuntiArr.Item(2))
			'	ScriveLogGlobale(filelog, "Pixel 4: " & listaPuntiArr.Item(3).ToString() & " -> " & quantiPuntiArr.Item(3))
			'	ScriveLogGlobale(filelog, "Pixel 5: " & listaPuntiArr.Item(4).ToString() & " -> " & quantiPuntiArr.Item(4))
			'Catch ex As Exception

			'End Try
			' Dim Nome As String = MP & "/Appoggio/RidBN_" & NomeEsteso & ".jpg"

			'QQ3 = (CInt(QQ3 / 100) * 100)
			Punti = QQ3 ' listaPunti.Count - 1 & "-" & quantiPunti.Item(0) & "-" & quantiPuntiArr.Item(0)

			'QQ1 = QQ1 * 80 / 100
			PuntiCornice = QQ1 ' listaPuntiArr.Item(0) & "-" & quantiPuntiArr.Item(0)

			HashColore = QQTot ' QQ1 ' listaPuntiArr.Item(1) & "-" & quantiPuntiArr.Item(1)

			'ScriveLogGlobale(filelog, "Salvataggio file")
			'img2.Save(NomeFileRidottoBN2)

			'img2.Dispose()
		Catch ex As Exception
			'NomeFileRidottoBN2 = ""
			Punti = -1
			ScriveLogGlobale(filelog, "ERROR: " & ex.Message)
		End Try

		Dim gf As New GestioneFilesDirectory
		gf.EliminaFileFisico(NomeFileRidotto)
		'gf.EliminaFileFisico(NomeFileRidottoBN)

		Return NomeFileRidottoBN & ";" & Punti & ";" & PuntiDiagonale & ";" & PuntiCornice & ";" & Hash & ";" & HashColore & ";"
	End Function

	Public Function AcquisisceExif(fileLog As String, file_name As String) As StrutturaJPG
		Dim gf As New GestioneFilesDirectory
		Dim Ritorno As New StrutturaJPG

		Ritorno.Descrizione = ""
		Ritorno.Commento = ""
		Ritorno.Software = ""
		Ritorno.DateTimeTag = ""
		Ritorno.XRes = ""
		Ritorno.YRes = ""

		If (file_name.ToUpper.Contains(".JPEG") Or file_name.ToUpper.Contains(".JPG")) And gf.EsisteFile(file_name) Then
			Try
				ScriveLogGlobale(fileLog, "")
				ScriveLogGlobale(fileLog, "Nome file: " & file_name)

				Dim _exif As ExifTagCollection = New ExifTagCollection(file_name)

				For Each tag As ExifTag In _exif
					Select Case tag.FieldName
						Case "ImageDescription"
							Ritorno.Descrizione = tag.Value.ToString().Trim.Replace("'", "''")
						Case "UserComment"
							Ritorno.Commento = tag.Value.ToString().Trim.Replace("'", "''")
						Case "Software"
							Ritorno.Software = tag.Value.ToString().Trim.Replace("'", "''")
						Case "DateTime"
							Ritorno.DateTimeTag = tag.Value.ToString().Trim.Replace("'", "''")
						Case "XResolution"
							Ritorno.XRes = tag.Value.ToString().Trim.Replace("'", "''")
						Case "YResolution"
							Ritorno.YRes = tag.Value.ToString().Trim.Replace("'", "''")
					End Select
				Next

				'Dim ss As System.Type = Ritorno.Descrizione.GetType
				'ScriveLogGlobale(fileLog, "Acquisizione EXIF: >Tipo: " & ss.Name)

				If Ritorno.Descrizione.Length > 100 Then Ritorno.Descrizione = Mid(Ritorno.Descrizione, 1, 97) & "..."
				If Ritorno.Commento.Length > 100 Then Ritorno.Commento = Mid(Ritorno.Commento, 1, 97) & "..."
				If Ritorno.Software.Length > 20 Then Ritorno.Software = Mid(Ritorno.Software, 1, 17) & "..."
				If Ritorno.DateTimeTag.Length > 20 Then Ritorno.DateTimeTag = Mid(Ritorno.DateTimeTag, 1, 17) & "..."
				If Ritorno.XRes.Length > 5 Then Ritorno.XRes = Mid(Ritorno.XRes, 1, 5)
				If Ritorno.YRes.Length > 5 Then Ritorno.YRes = Mid(Ritorno.YRes, 1, 5)

				ScriveLogGlobale(fileLog, "Acquisizione EXIF: Descrizione: " & Ritorno.Descrizione)
				ScriveLogGlobale(fileLog, "Acquisizione EXIF: Commento: " & Ritorno.Commento)
				ScriveLogGlobale(fileLog, "Acquisizione EXIF: Software: " & Ritorno.Software)
				ScriveLogGlobale(fileLog, "Acquisizione EXIF: DateTimeTag: " & Ritorno.DateTimeTag)
				ScriveLogGlobale(fileLog, "Acquisizione EXIF: XRes: " & Ritorno.XRes)
				ScriveLogGlobale(fileLog, "Acquisizione EXIF: YRes: " & Ritorno.YRes)

				ScriveLogGlobale(fileLog, "Acquisizione EXIF completata")
			Catch ex As Exception
				ScriveLogGlobale(fileLog, "Acquisizione EXIF errore: " & ex.Message)
			End Try
		End If

		Return Ritorno
	End Function

	Public Function CalcolaPuntini(Mp As String, ByVal file_name As String, fileLog As String, numeroImmagineConvertita As String,
								   Optional DebugTest As Boolean = False, Optional Proporziona As Boolean = False) As StrutturaJPG
		Dim gf As New GestioneFilesDirectory
		Dim Ritorno As New StrutturaJPG
		Dim s As New StrutturaJPG
		Static x As Random = New Random()

		Dim y As Long = -1

		y = x.Next(1000)

		Dim NomeEsteso As String = Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & "_" & y.ToString.Trim & "_Ptn"

		gf.EliminaFileFisico(Mp & "/Appoggio/BW_" & NomeEsteso & ".jpg")
		gf.EliminaFileFisico(Mp & "/Appoggio/Rid_" & NomeEsteso & ".jpg")
		''gf.EliminaFileFisico(Mp & "/Appoggio/RidColore_" & NomeEsteso & ".jpg")
		'gf.EliminaFileFisico(Mp & "/Appoggio/RidBN_" & NomeEsteso & ".jpg")

		Dim rit As String = "*"

		ScriveLogGlobale(fileLog, "Conversione per calcolo puntini immagine numero: ---" & numeroImmagineConvertita & "---")

		Ritorno.Descrizione = ""
		Ritorno.Commento = ""
		Ritorno.Software = ""
		Ritorno.DateTimeTag = ""
		Ritorno.XRes = ""
		Ritorno.YRes = ""

		Try
			rit = ContaPunti(Mp, file_name, fileLog, True)
			Dim r() As String = rit.Split(";")
			If r(0) <> "" Then
				' ScriveLogGlobale(fileLog, "Punti: " & r(1))
				'Ritorno.Hash = r(2)
				'Ritorno.HashColore = r(3)
				Ritorno.Punti = r(1)
				Ritorno.PuntiDiagonale = r(2)
				Ritorno.PuntiCornice = r(3)
				Ritorno.Sezione1 = r(4)
				Ritorno.Sezione2 = r(5)
				Ritorno.Hash = GeneraHash(file_name, fileLog)

				Dim dataOra As String = gf.TornaDataDiCreazione(file_name)
				Dim imgG As Bitmap = LoadBitmapSenzaLock(file_name)
				Dim width As Integer = imgG.Width
				Dim height As Integer = imgG.Height
				Ritorno.DataOra = dataOra
				Ritorno.Width = width
				Ritorno.Height = height
				ScriveLogGlobale(fileLog, "Data Ora: " & dataOra)
				ScriveLogGlobale(fileLog, "Width: " & width)
				ScriveLogGlobale(fileLog, "Height: " & height)

				'Dim rJpg As StrutturaJPG = AcquisizionePunti(Mp, fileLog, r(0), DebugTest)
				'Ritorno.PuntiCornice = rJpg.PuntiCornice
				'Ritorno.PuntiDiagonale = rJpg.PuntiDiagonale
				'ScriveLogGlobale(fileLog, "Punti Cornice: " & Ritorno.PuntiCornice)
				'ScriveLogGlobale(fileLog, "Punti Diagonale: " & Ritorno.PuntiDiagonale)

				gf.EliminaFileFisico(r(0))

				' Confronto per immagine ridotta
				'Dim NomeFileRidotto As String = Mp & "/Appoggio/Rid_" & NomeEsteso & ".jpg"
				'rit = RidimensionaMantenendoProporzioni(file_name, NomeFileRidotto, 200, Proporziona)
				'If rit <> "OK" Then
				'	ScriveLogGlobale(fileLog, rit)
				'	Ritorno.Hash = "ERROR: " & rit
				'Else
				'	ScriveLogGlobale(fileLog, "Ridimensionamento effettuato")
				'	rit = ContaPunti(Mp, NomeFileRidotto, fileLog, False)
				'	r = rit.Split(";")
				'	rJpg = AcquisizionePunti(Mp, fileLog, r(0), DebugTest)
				'	Ritorno.Hash = CInt((rJpg.PuntiCornice) / 1000000)
				'	Ritorno.HashColore = CInt((rJpg.PuntiDiagonale) / 1000000)
				'	ScriveLogGlobale(fileLog, "Hash: " & Ritorno.Hash)

				'	gf.EliminaFileFisico(r(0))
				'End If
				' Confronto per immagine ridotta
			Else
				Ritorno.Sezione1 = "ERROR: Problemi nel conteggio"
			End If
		Catch ex As Exception
			Ritorno.Sezione1 = "ERROR: " & ex.Message
		End Try

		gf.EliminaFileFisico(Mp & "/Appoggio/Rid_" & NomeEsteso & ".jpg")

		Return Ritorno

		''Try
		''	ScriveLogGlobale(fileLog, "Conversione in BN per calcolo puntini")
		''	rit = ConverteImmaginInBN(file_name, Mp & "/Appoggio/BW_" & NomeEsteso & ".jpg", fileLog, False)

		''	If rit <> "*" Then
		''		ScriveLogGlobale(fileLog, rit)
		''		Ritorno.Hash = "ERROR: " & rit
		''	Else
		''		'ScriveLogGlobale(fileLog, "Ridimensionamento per calcolo puntini Colore")
		''		'rit = RidimensionaMantenendoProporzioni(file_name, Mp & "/Appoggio/RidColore_" & NomeEsteso & ".jpg", 150, Proporziona)
		''		'If rit <> "OK" Then
		''		'	ScriveLogGlobale(fileLog, rit)
		''		'	Ritorno = rit
		''		'Else
		''		rit = RidimensionaMantenendoProporzioni(Mp & "/Appoggio/BW_" & NomeEsteso & ".jpg", Mp & "/Appoggio/Rid_" & NomeEsteso & ".jpg", 200, Proporziona)
		''		If rit <> "OK" Then
		''			ScriveLogGlobale(fileLog, rit)
		''			Ritorno.Hash = "ERROR: " & rit
		''		Else
		''			'ScriveLogGlobale(fileLog, "Acquisizione punti per calcolo puntini")
		''			ScriveLogGlobale(fileLog, "Calcolo puntini")
		''			' Dim img As Bitmap = New Bitmap(Mp & "/Appoggio/Rid_" & NomeEsteso & ".jpg")

		''			rit = ContaPunti(Mp, Mp & "/Appoggio/Rid_" & NomeEsteso & ".jpg", fileLog)
		''			r = rit.Split(";")
		''			Ritorno.Hash = r(2)
		''			Ritorno.HashColore = r(3)
		''			Ritorno.Punti = r(1)
		''			'img3.Save(Mp & "/Appoggio/RidBN_" & NomeEsteso & "_Neg.jpg")
		''			'img4.Save(Mp & "/Appoggio/RidBN_" & NomeEsteso & "_Neg2.jpg")

		''			Ritorno = AcquisizionePunti(Mp, fileLog, Mp & "/Appoggio/RidBN_" & NomeEsteso & ".jpg", DebugTest)

		''			'Ritorno = R.PuntiDiagonale.ToString.Trim & ";" & R.PuntiCornice.ToString.Trim & ";" & R.Punti & ";" & R.Hash & ";" & width2 & ";" & height2 & ";" & R.DataOra & ";" & R.HashColore & ";"

		''			' ScriveLogGlobale(fileLog, "Acquisizione punti per calcolo puntini. Rilevati: " & Ritorno)

		''			'img2.Dispose()
		''		End If
		''	End If
		''	'End If
		''Catch ex As Exception
		''	Ritorno.Hash = "ERROR: " & ex.Message
		''End Try

		'If DebugTest = False Then
		gf.EliminaFileFisico(Mp & "/Appoggio/BW_" & NomeEsteso & ".jpg")
		'If DebugTest = False Then
		gf.EliminaFileFisico(Mp & "/Appoggio/Rid_" & NomeEsteso & ".jpg")
		'gf.EliminaFileFisico(Mp & "/Appoggio/RidColore_" & NomeEsteso & ".jpg")
		gf.EliminaFileFisico(Mp & "/Appoggio/RidBN_" & NomeEsteso & ".jpg")
		'End If

		Return Ritorno
	End Function

	'Public Function CalcolaMD5(Mp As String, ByVal file_name As String, fileLog As String, numeroImmagineConvertita As String, Debug As Boolean) As StrutturaJPG
	'	'Dim gf As New GestioneFilesDirectory
	'	'Static x As Random = New Random()

	'	'Dim y As Long = -1
	'	Dim s As New StrutturaJPG

	'	'y = x.Next(1000)

	'	'Dim NomeEsteso As String = Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & "_" & y.ToString.Trim

	'	'Try
	'	'	ScriveLogGlobale(fileLog, "Conversione immagine numero: ---" & numeroImmagineConvertita & "---")
	'	'	Dim dataOra As String = gf.TornaDataDiCreazione(file_name)
	'	'	ScriveLogGlobale(fileLog, "Data immagine: " & dataOra)

	'	'	' Dim imgG As New Bitmap(file_name)
	'	'	Dim imgG As Bitmap = LoadBitmapSenzaLock(file_name)
	'	'	Dim width As Integer = imgG.Width
	'	'	Dim height As Integer = imgG.Height
	'	'	'Dim propItems As PropertyItem() = imgG.PropertyItems
	'	'	'For Each p As PropertyItem In propItems
	'	'	'	Dim pp = p
	'	'	'Next
	'	'	imgG.Dispose()
	'	'	ScriveLogGlobale(fileLog, "Dimensioni immagine: " & width & " x " & height)

	'	'	gf.CreaDirectoryDaPercorso(Mp & "/Appoggio/")
	'	'	gf.EliminaFileFisico(Mp & "/Appoggio/BW_" & NomeEsteso & ".jpg")
	'	'	gf.EliminaFileFisico(Mp & "/Appoggio/Rid_" & NomeEsteso & ".jpg")
	'	'	gf.EliminaFileFisico(Mp & "/Appoggio/RidColore_" & NomeEsteso & ".jpg")
	'	'	gf.EliminaFileFisico(Mp & "/Appoggio/RidBN_" & NomeEsteso & ".jpg")

	'	'	Dim rit As String = "*"

	'	'	ScriveLogGlobale(fileLog, "Conversione in BN")
	'	'	rit = ConverteImmaginInBN(file_name, Mp & "/Appoggio/BW_" & NomeEsteso & ".jpg", fileLog, False)

	'	'	If rit <> "*" Then
	'	'		ScriveLogGlobale(fileLog, rit)
	'	'		s.Hash = rit
	'	'	Else
	'	'		ScriveLogGlobale(fileLog, "Ridimensionamento BN")
	'	'		rit = RidimensionaMantenendoProporzioni(Mp & "/Appoggio/BW_" & NomeEsteso & ".jpg", Mp & "/Appoggio/Rid_" & NomeEsteso & ".jpg", 75, False)
	'	'		If rit <> "OK" Then
	'	'			ScriveLogGlobale(fileLog, rit)
	'	'			s.Hash = rit
	'	'		Else
	'	'			ScriveLogGlobale(fileLog, "Ridimensionamento Colore file originale")
	'	'			rit = RidimensionaMantenendoProporzioni(file_name, Mp & "/Appoggio/RidColore_" & NomeEsteso & ".jpg", 75, False)
	'	'			If rit <> "OK" Then
	'	'				ScriveLogGlobale(fileLog, rit)
	'	'				s.Hash = rit
	'	'			Else
	'	'				ScriveLogGlobale(fileLog, "Acquisizione punti validi")
	'	'				Dim img As Bitmap = LoadBitmapSenzaLock(Mp & "/Appoggio/Rid_" & NomeEsteso & ".jpg")
	'	'				' Dim img As Bitmap = New Bitmap(Mp & "/Appoggio/Rid_" & NomeEsteso & ".jpg")
	'	'				Dim width2 As Integer = img.Width
	'	'				Dim height2 As Integer = img.Height
	'	'				Dim Punti = 0
	'	'				Dim img2 As Bitmap = New Bitmap(width2, height2)
	'	'				Dim Nero As Color = New Color()
	'	'				Nero = Color.FromArgb(0, 0, 0)
	'	'				Dim GrigioChiaro As Color = New Color()
	'	'				GrigioChiaro = Color.FromArgb(186, 186, 186)
	'	'				Dim Grigio As Color = New Color()
	'	'				Grigio = Color.FromArgb(108, 108, 108)
	'	'				Dim Bianco As Color = New Color()
	'	'				Bianco = Color.FromArgb(255, 255, 255)
	'	'				For i As Integer = 0 To width2 - 1
	'	'					For k As Integer = 0 To height2 - 1
	'	'						Dim pixelColor As Color = img.GetPixel(i, k)

	'	'						'ScriveLogGlobale(fileLog, "Colori " & i & "," & k & ": R. " & pixelColor.R & " G. " & pixelColor.G & " B. " & pixelColor.B)
	'	'						If pixelColor.R > valoreColore Or pixelColor.G > valoreColore Or pixelColor.B > valoreColore Then
	'	'							img2.SetPixel(i, k, Bianco)
	'	'						Else
	'	'							If pixelColor.R > (valoreColore / 1.5) Or pixelColor.G > (valoreColore / 1.5) Or pixelColor.B > (valoreColore / 1.5) Then
	'	'								img2.SetPixel(i, k, GrigioChiaro)
	'	'								Punti += 27
	'	'							Else
	'	'								If pixelColor.R > (valoreColore / 2.5) Or pixelColor.G > (valoreColore / 2.5) Or pixelColor.B > (valoreColore / 2.5) Then
	'	'									img2.SetPixel(i, k, Grigio)
	'	'									Punti += 72
	'	'								Else
	'	'									img2.SetPixel(i, k, Nero)
	'	'									Punti += 2727
	'	'								End If

	'	'							End If
	'	'						End If
	'	'					Next
	'	'				Next
	'	'				ScriveLogGlobale(fileLog, "Salvataggio file bn")
	'	'				img2.Save(Mp & "/Appoggio/RidBN_" & NomeEsteso & ".jpg")

	'	'				Dim PuntiVari As String = AcquisizionePunti(Mp, fileLog, Mp & "/Appoggio/RidBN_" & NomeEsteso & ".jpg", Punti, dataOra, width, height, Debug)
	'	'				ScriveLogGlobale(fileLog, "Acquisizione punti validi. Ritorno: " & PuntiVari)
	'	'				If PuntiVari.Contains(";") Then
	'	'					Dim Pv() As String = PuntiVari.Split(";")

	'	'					s.PuntiDiagonale = Pv(0)
	'	'					s.PuntiCornice = Pv(1)
	'	'					s.Punti = Pv(2)
	'	'					s.Hash = Pv(3)
	'	'					s.HashColore = Pv(7)
	'	'				End If

	'	'				img2.Dispose()

	'	'				s.Width = width
	'	'				s.Height = height
	'	'				s.DataOra = dataOra
	'	'			End If
	'	'		End If
	'	'	End If

	'	'Catch ex As Exception
	'	'	s.Hash = "ERROR: " & ex.Message
	'	'End Try

	'	'gf.EliminaFileFisico(Mp & "/Appoggio/BW_" & NomeEsteso & ".jpg")
	'	'gf.EliminaFileFisico(Mp & "/Appoggio/Rid_" & NomeEsteso & ".jpg")
	'	'gf.EliminaFileFisico(Mp & "/Appoggio/RidColore_" & NomeEsteso & ".jpg")
	'	'gf.EliminaFileFisico(Mp & "/Appoggio/RidBN_" & NomeEsteso & ".jpg")

	'	Return s
	'End Function

	Private Function ControllaPunto(pixelColor As Color, Incremento As Integer) As Long
		Dim Ritorno As Long = 0

		'Ritorno = Val(pixelColor.R) + Val(pixelColor.G) + Val(pixelColor.B)
		'If pixelColor.R > valoreColore Or pixelColor.G > valoreColore Or pixelColor.B > valoreColore Then
		'	'Ritorno = 1
		'	Ritorno = 1 ' 2727 ' + (75 * Incremento)
		'Else
		'	If pixelColor.R > (valoreColore / 1.5) Or pixelColor.G > (valoreColore / 1.5) Or pixelColor.B > (valoreColore / 1.5) Then
		'		Ritorno = 72 ' + (75 * Incremento)
		'	Else
		'		If pixelColor.R > (valoreColore / 2.5) Or pixelColor.G > (valoreColore / 2.5) Or pixelColor.B > (valoreColore / 2.5) Then
		'			Ritorno = 27 ' + (75 * Incremento)
		'		End If
		'	End If
		'End If

		'If pixelColor = Bianco Then
		'	Ritorno = 0
		'Else
		If pixelColor = GrigioChiaro Then
			Ritorno = 3 '* Incremento ' Val(pixelColor.R) + Val(pixelColor.G) + Val(pixelColor.B)
		Else
			If pixelColor = Grigio Then
				Ritorno = 5 '* Incremento ' Val(pixelColor.R) + Val(pixelColor.G) + Val(pixelColor.B)
			Else
				Ritorno = 10 '* Incremento ' Val(pixelColor.R) + Val(pixelColor.G) + Val(pixelColor.B)
			End If

		End If
		'End
		'Ritorno = pixelColor.G ' CInt(System.Drawing.ColorTranslator.ToWin32(pixelColor) / 1000000) * 10 ' Val(pixelColor.R) + Val(pixelColor.G) + Val(pixelColor.B)

		'If pixelColor.R > 0 Or pixelColor.R > 0 Or pixelColor.B > 0 Then
		'	Ritorno = (pixelColor.R + pixelColor.R + pixelColor.B)
		'End If

		Return Ritorno
	End Function

	Private Function AcquisizionePunti(Mp As String, fileLog As String, NomeFileIMG As String, Debug As Boolean) As StrutturaJPG
		Dim Ritorno As String = ""

		Dim img As Bitmap = LoadBitmapSenzaLock(NomeFileIMG)
		Dim width2 As Integer = img.Width
		Dim height2 As Integer = img.Height

		Dim imgCopia As Bitmap
		If Debug = True Then
			imgCopia = LoadBitmapSenzaLock(NomeFileIMG)
		End If

		ScriveLogGlobale(fileLog, "Acquisizione punti validi. Dimensione immagine: " & width2 & " x " & height2)
		'ScriveLogGlobale(fileLog, "Acquisizione punti validi per corpo.")

		'Dim PuntiCorpo As Long = 0

		''Dim Nero As Color = New Color()
		''Nero = Color.FromArgb(0, 0, 0)
		''Dim Bianco As Color = New Color()
		''Bianco = Color.FromArgb(255, 255, 255)
		'For i As Integer = 0 To width2 - 1
		'	For k As Integer = 0 To height2 - 1
		'		Dim pixelColor As Color = img.GetPixel(i, k)
		'		Dim Colore As Long = (CInt(pixelColor.R) * 100) + (CInt(pixelColor.G) * 10) + CInt(pixelColor.B)

		'		PuntiCorpo += Colore
		'	Next
		'Next

		'ScriveLogGlobale(fileLog, "Acquisizione punti validi. Dimensione immagine: " & width2 & " x " & height2)

		Dim PassoX As Single
		Dim PassoY As Single
		Dim Fine As Single

		If width2 < height2 Then
			PassoX = 1
			PassoY = (height2 - 1) / (width2 - 1)
			Fine = height2
		Else
			If width2 = height2 Then
				PassoX = 1
				PassoY = 1
				Fine = width2
			Else
				PassoX = (width2 - 1) / (height2 - 1)
				PassoY = 1
				Fine = width2
			End If
		End If

		ScriveLogGlobale(fileLog, "Acquisizione punti validi per diagonale. Passo: " & PassoX & " x " & PassoY & " - Fine: " & Fine)

		Dim cX As Single = 0
		Dim cY As Single = 0
		Dim PuntiDiagonale As Long = 0
		Dim PuntiCornice As Long = 0

		'Dim Bianco As Color = New Color()
		'Bianco = Color.FromArgb(255, 255, 255)

		Dim Rosso As Color = New Color()
		Rosso = Color.FromArgb(255, 0, 0)
		Dim Verde As Color = New Color()
		Verde = Color.FromArgb(0, 255, 0)
		Dim Blu As Color = New Color()
		Blu = Color.FromArgb(0, 0, 255)

		' Acquisizione punti in diagonale
		While cX <= Fine
			If (Int(cX) < width2 - 1 And Int(cY) < height2 - 1) Then
				'Dim pixelColor As Color = img.GetPixel(Int(cX), Int(cY))
				'Dim Colore As Long = (CInt(pixelColor.R) * 100) + (CInt(pixelColor.G) * 10) + CInt(pixelColor.B)

				' If pixelColor.R > 0 Then
				PuntiDiagonale += ControllaPunto(img.GetPixel(Int(cX), Int(cY)), Int(cY))
				'PuntiDiagonale += ControllaPunto(img.GetPixel(Int(cX) + 1, Int(cY)), Int(cY))

				' End If
				'ScriveLogGlobale(fileLog, "1: " & CInt(cX) & " - " & CInt(cY) & " Colore: R. " & pixelColor.R & "-" & pixelColor.G & "-" & pixelColor.B & " -> " & PuntiDiagonale)
				''ScriveLogGlobale(fileLog, pixelColor.R > 0)
				'ScriveLogGlobale(fileLog, "Colore: " & Colore & " - PuntiDiagonale: " & PuntiDiagonale)

				Dim cX2 As Single = (width2 - 1) - cX
				Dim cY2 As Single = cY ' (height2 - 1) - cY

				If (Int(cX2) < width2 - 1 And Int(cY2) < height2 - 1) Then
					'pixelColor = img.GetPixel(Int(cX2), Int(cY2))
					''Colore = (CInt(pixelColor.R) * 100) + (CInt(pixelColor.G) * 10) + CInt(pixelColor.B)

					'If pixelColor.R > valoreColore Or pixelColor.G > valoreColore Or pixelColor.B > valoreColore Then
					'	PuntiDiagonale += 2
					'Else
					'	If pixelColor.R > valoreColore / 2 Or pixelColor.G > valoreColore / 2 Or pixelColor.B > valoreColore / 2 Then
					'		PuntiDiagonale += 1
					'	End If
					'End If
					PuntiDiagonale += ControllaPunto(img.GetPixel(Int(cX2), Int(cY2)), Int(cY2))
					'PuntiDiagonale += ControllaPunto(img.GetPixel(Int(cX2) - 1, Int(cY2)), Int(cY2))

					'If pixelColor2.R > 0 Then
					'PuntiDiagonale += Colore
					'End If
					'ScriveLogGlobale(fileLog, pixelColor2 = Bianco)
					'ScriveLogGlobale(fileLog, "2: " & CInt(cX2) & " - " & CInt(cY2) & " Colore: R. " & pixelColor2.R & " -> " & PuntiDiagonale)

					If Debug = True Then
						imgCopia.SetPixel(CInt(cX), CInt(cY), Rosso)
						imgCopia.SetPixel(CInt(cX2), CInt(cY2), Rosso)
						'imgCopia.SetPixel(CInt(cX) + 1, CInt(cY), Rosso)
						'imgCopia.SetPixel(CInt(cX2) - 1, CInt(cY2), Rosso)
					End If
				End If
			End If

			cX += PassoX
			cY += PassoY
		End While
		' Acquisizione punti in diagonale

		'Acquisizione punti cornice
		ScriveLogGlobale(fileLog, "Acquisizione punti validi per cornice 1")
		cY = (height2 - 1) / 4
		For cX = 0 To width2 - 1
			'Dim pixelColor As Color = img.GetPixel(cX, cY)
			''Dim Colore As Long = (CInt(pixelColor.R) * 100) + (CInt(pixelColor.G) * 10) + CInt(pixelColor.B)

			'' If pixelColor.R > 0 Then
			'If pixelColor.R > valoreColore Or pixelColor.G > valoreColore Or pixelColor.B > valoreColore Then
			'	PuntiCornice += 2
			'Else
			'	If pixelColor.R > valoreColore / 2 Or pixelColor.G > valoreColore / 2 Or pixelColor.B > valoreColore / 2 Then
			'		PuntiCornice += 1
			'	End If
			'End If

			PuntiCornice += ControllaPunto(img.GetPixel(cX, cY), Int(cX))
			'PuntiCornice += ControllaPunto(img.GetPixel(cX, cY + 1), Int(cX))

			' End If

			Dim cY2 As Integer = (height2 - 1) - cY

			'pixelColor = img.GetPixel(cX, cY2)
			'' Colore = (CInt(pixelColor.R) * 100) + (CInt(pixelColor.G) * 10) + CInt(pixelColor.B)

			''If pixelColor2.R > 0 Then
			'If pixelColor.R > valoreColore Or pixelColor.G > valoreColore Or pixelColor.B > valoreColore Then
			'	PuntiCornice += 2
			'Else
			'	If pixelColor.R > valoreColore / 2 Or pixelColor.G > valoreColore / 2 Or pixelColor.B > valoreColore / 2 Then
			'		PuntiCornice += 1
			'	End If
			'End If
			''End If
			PuntiCornice += ControllaPunto(img.GetPixel(cX, cY2), Int(cX))
			'PuntiCornice += ControllaPunto(img.GetPixel(cX, cY2 - 1), Int(cX))

			If Debug = True Then
				imgCopia.SetPixel(CInt(cX), CInt(cY), Verde)
				imgCopia.SetPixel(CInt(cX), CInt(cY2), Verde)
				'imgCopia.SetPixel(CInt(cX), CInt(cY) + 1, Verde)
				'imgCopia.SetPixel(CInt(cX), CInt(cY2) - 1, Verde)
			End If
		Next

		ScriveLogGlobale(fileLog, "Acquisizione punti validi per cornice 2")
		cX = (width2 - 1) / 4
		For cY = 0 To height2 - 1
			'Dim pixelColor As Color = img.GetPixel(cX, cY)
			''Dim Colore As Long = CInt(pixelColor.R) + CInt(pixelColor.G) + CInt(pixelColor.B)

			'' If pixelColor.R > 0 Then
			'If pixelColor.R > valoreColore Or pixelColor.G > valoreColore Or pixelColor.B > valoreColore Then
			'	PuntiCornice += 2
			'Else
			'	If pixelColor.R > valoreColore / 2 Or pixelColor.G > valoreColore / 2 Or pixelColor.B > valoreColore / 2 Then
			'		PuntiCornice += 1
			'	End If
			'End If
			PuntiCornice += ControllaPunto(img.GetPixel(cX, cY), Int(cY))
			'PuntiCornice += ControllaPunto(img.GetPixel(cX + 1, cY), Int(cY))

			' End If

			Dim cX2 As Integer = (width2 - 1) - cX

			'pixelColor = img.GetPixel(cX2, cY)
			''Colore = CInt(pixelColor.R) + CInt(pixelColor.G) + CInt(pixelColor.B)

			'' If pixelColor2.R > 0 Then
			'If pixelColor.R > valoreColore Or pixelColor.G > valoreColore Or pixelColor.B > valoreColore Then
			'	PuntiCornice += 2
			'Else
			'	If pixelColor.R > valoreColore / 2 Or pixelColor.G > valoreColore / 2 Or pixelColor.B > valoreColore / 2 Then
			'		PuntiCornice += 1
			'	End If
			'End If
			' End If
			PuntiCornice += ControllaPunto(img.GetPixel(cX2, cY), Int(cY))
			'PuntiCornice += ControllaPunto(img.GetPixel(cX2 - 1, cY), Int(cY))

			If Debug = True Then
				imgCopia.SetPixel(CInt(cX), CInt(cY), Blu)
				imgCopia.SetPixel(CInt(cX2), CInt(cY), Blu)
				'imgCopia.SetPixel(CInt(cX) + 1, CInt(cY), Blu)
				'imgCopia.SetPixel(CInt(cX2) - 1, CInt(cY), Blu)
			End If
		Next
		'Acquisizione punti cornice

		'PuntiDiagonale = (Math.Ceiling(PuntiDiagonale / 10) * 10)
		'PuntiCornice = (Math.Ceiling(PuntiCornice / 10) * 10)
		'PuntiCorpo = (Math.Ceiling(PuntiCorpo / 10) * 10)

		'ScriveLogGlobale(fileLog, "Acquisizione hash BN")
		'Dim fs As FileStream = New FileStream(NomeFileIMG, FileMode.Open, FileAccess.Read)
		'Dim arr(fs.Length) As Byte
		'fs.Read(arr, 0, fs.Length)
		'fs.Close()

		'Dim tmpHash() As Byte
		'tmpHash = New MD5CryptoServiceProvider().ComputeHash(arr)

		'Dim Hash As String = ""
		'For Each b As Byte In tmpHash
		'	Dim bb As String = b.ToString
		'	For i As Integer = bb.Length + 1 To 3
		'		bb = "0" & bb
		'	Next
		'	Hash &= bb
		'Next

		'ScriveLogGlobale(fileLog, "Acquisizione hash Colore")
		'Dim fsC As FileStream = New FileStream(NomeFileColore, FileMode.Open, FileAccess.Read)
		'Dim arrC(fsC.Length) As Byte
		'fsC.Read(arrC, 0, fsC.Length)
		'fsC.Close()

		'Dim tmpHashC() As Byte
		'tmpHashC = New MD5CryptoServiceProvider().ComputeHash(arrC)

		'Dim HashColore As String = ""
		'For Each b As Byte In tmpHashC
		'	Dim bb As String = b.ToString
		'	For i As Integer = bb.Length + 1 To 3
		'		bb = "0" & bb
		'	Next
		'	HashColore &= bb
		'Next

		Dim Hash As String = "0"
		Dim HashColore As String = "0"

		Dim s As New StrutturaJPG
		s.PuntiDiagonale = PuntiDiagonale
		s.PuntiCornice = PuntiCornice
		'Ritorno = PuntiDiagonale.ToString.Trim & ";" & PuntiCornice.ToString.Trim & ";" & PuntiCorpo.Trim & ";" & Hash & ";" & Width & ";" & Height & ";" & DataOra & ";" & HashColore & ";"

		If Debug = True Then
			Dim NomeEsteso As String = Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & Format(Now.Millisecond, "00")
			imgCopia.Save(Mp & "/Appoggio/Finale_" & NomeEsteso & ".jpg")
			imgCopia.Dispose()
		End If

		Return s
	End Function

	Public Function GeneraHash(NomeFileIMG As String, fileLog As String) As String
		ScriveLogGlobale(fileLog, "Acquisizione hash BN")
		Dim fs As FileStream = New FileStream(NomeFileIMG, FileMode.Open, FileAccess.Read)
		Dim arr(fs.Length) As Byte
		fs.Read(arr, 0, fs.Length)
		fs.Close()

		Dim tmpHash() As Byte
		tmpHash = New MD5CryptoServiceProvider().ComputeHash(arr)

		Dim Hash As String = ""
		For Each b As Byte In tmpHash
			Dim bb As String = b.ToString
			For i As Integer = bb.Length + 1 To 3
				bb = "0" & bb
			Next
			Hash &= bb
		Next
		ScriveLogGlobale(fileLog, "Hash BN Acquisito")

		Return Hash
	End Function
End Class
