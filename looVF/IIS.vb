Imports System.DirectoryServices

Public Class IIS
	Public Function CreateVDir(nomeDir As String, pathDir As String)
		Try
			Dim procStartInfo As New ProcessStartInfo
			Dim procExecuting As New Process

			With procStartInfo
				.UseShellExecute = True
				.FileName = "C:\Windows\system32\inetsrv\APPCMD"
				.Arguments = " ADD vdir /app.name:""Default Web Site/"" /path:/looVF\" & nomeDir & " /physicalPath:""" & pathDir & """"
				.WindowStyle = ProcessWindowStyle.Normal
				.Verb = "runas" 'add this to prompt for elevation
			End With

			procExecuting = Process.Start(procStartInfo)
			procExecuting.WaitForExit()

			Dim gf As New GestioneFilesDirectory
			gf.CreaAggiornaFile("C:\Appoggio\out.txt", "C:\Windows\system32\inetsrv\APPCMD ADD vdir /app.name:""Default Web Site/"" /path:/looVF\" & nomeDir & " /physicalPath:""" & pathDir & """")
			gf = Nothing

			Return "*"
		Catch ex As Exception
			Return "ERROR: " & ex.Message
		End Try
	End Function
End Class
