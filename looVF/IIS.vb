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

	Public Sub CreateVirtualDir(ByVal WebSite As String, ByVal AppName As String, ByVal Path As String)
		Dim IISSchema As New System.DirectoryServices.DirectoryEntry("IIS://" & WebSite & "/Schema/AppIsolated")
		Dim CanCreate As Boolean = Not IISSchema.Properties("Syntax").Value.ToString.ToUpper() = "BOOLEAN"
		IISSchema.Dispose()

		If CanCreate Then
			Dim PathCreated As Boolean

			Try
				Dim IISAdmin As New System.DirectoryServices.DirectoryEntry("IIS://" & WebSite & "/W3SVC/1/Root")

				'make sure folder exists
				If Not System.IO.Directory.Exists(Path) Then
					System.IO.Directory.CreateDirectory(Path)
					PathCreated = True
				End If

				'If the virtual directory already exists then delete it
				For Each VD As System.DirectoryServices.DirectoryEntry In IISAdmin.Children
					If VD.Name = AppName Then
						IISAdmin.Invoke("Delete", New String() {VD.SchemaClassName, AppName})
						IISAdmin.CommitChanges()
						Exit For
					End If
				Next VD

				'Create and setup new virtual directory
				Dim VDir As System.DirectoryServices.DirectoryEntry = IISAdmin.Children.Add(AppName, "IIsWebVirtualDir")
				VDir.Properties("Path").Item(0) = Path
				VDir.Properties("AppFriendlyName").Item(0) = AppName
				VDir.Properties("EnableDirBrowsing").Item(0) = False
				VDir.Properties("AccessRead").Item(0) = True
				VDir.Properties("AccessExecute").Item(0) = True
				VDir.Properties("AccessWrite").Item(0) = False
				VDir.Properties("AccessScript").Item(0) = True
				VDir.Properties("AuthNTLM").Item(0) = True
				VDir.Properties("EnableDefaultDoc").Item(0) = True
				VDir.Properties("DefaultDoc").Item(0) = "default.htm,default.aspx,default.asp"
				VDir.Properties("AspEnableParentPaths").Item(0) = True
				VDir.CommitChanges()

				'the following are acceptable params
				'INPROC = 0
				'OUTPROC = 1
				'POOLED = 2
				VDir.Invoke("AppCreate", 1)

			Catch Ex As Exception
				If PathCreated Then
					System.IO.Directory.Delete(Path)
				End If
				Throw Ex
			End Try
		End If
	End Sub
End Class
