
' Module1.vb
'
' Subject: Download file using ftp.
'
' Description: Call "downloadFileBW" or 
' "downloadFile" to download
' the most recent file in the directory of the 
' server identified by ftpUrl.
'
' downloadFileBW uses a BackgroundWorker to 
' download the file--it is the preferred
' way to download the file.
'
' downloadFile downloads the file without
' using a BackgroundWorker--it's use
' is discouraged.
'
' Notes: Most variable values in FtpDownload can
' be set using either the constructor or the
' properties. _filenamePattern can only be changed
' by using the property "FilenamePattern"--not through
' the constructor
'
' Usage:
' downloadFileBW("ftp://ftp.server.com", "anonymous", "me@gmail.com", "C:\temp\")
'      or 
' downloadFileNoBW("ftp://ftp.server.com", "anonymous", "me@gmail.com", "C:\temp\")
'
' Written by:   cgeier 07/18/2014
' Modified by:  cgeier 07/19/2014   renamed downloadFile to downloadFileBW
'                                   added downloadFileNoBW

Imports System.Threading
Imports System.ComponentModel
Imports test_ftp.Form1

Module Module1

	Private myFtpDownload As FtpDownload = Nothing
	Private cancelOperation As Boolean = False
	Private totalBytesDownloaded As Integer = 0
	Friend WithEvents BackgroundWorker2 As New BackgroundWorker


	Public Sub downloadFileNoBW(ByVal ftpUrl As String, ByVal username As String, ByVal password As String, ByVal downloadDirectory As String)
		myFtpDownload = New FtpDownload(ftpUrl, username, password, downloadDirectory)
		totalBytesDownloaded = myFtpDownload.getFileFromFtpServer()
		Console.WriteLine("Total Bytes Downloaded: " & totalBytesDownloaded)
	End Sub

	Public Sub downloadFileBW(ByVal ftpUrl As String, ByVal username As String, ByVal password As String, ByVal downloadDirectory As String)
		'reset cancelOperation
		cancelOperation = False

		'create new instance of FtpDownload
		'and setting values using the constructor
		'Alternatively, the following could be used:
		'
		'myFtpDownload = New FtpDownload()
		'
		'and the variables could be set using
		'the properties.
		'
		'ex: myFtpDownload.FtpUrl = "ftp://ftp.server.com"

		myFtpDownload = New FtpDownload(ftpUrl, username, password, downloadDirectory)

		'add support for cancellation and reporting progress
		BackgroundWorker2.WorkerSupportsCancellation = True
		BackgroundWorker2.WorkerReportsProgress = True

		If BackgroundWorker2.IsBusy <> True Then
			BackgroundWorker2.RunWorkerAsync(myFtpDownload)
		End If
	End Sub


	Private Sub BackgroundWorker2_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
		Dim worker As BackgroundWorker
		worker = DirectCast(sender, BackgroundWorker)


		Dim myFtpDownload As FtpDownload = DirectCast(e.Argument, FtpDownload)

		'call getFileFromFtpServerBW
		totalBytesDownloaded = myFtpDownload.getFileFromFtpServerBW(worker, e)

		Console.WriteLine("Total Bytes Downloaded: " & totalBytesDownloaded)
	End Sub

	Private Sub BackgroundWorker2_ProgressChanged(sender As System.Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker2.ProgressChanged
		' This event occurs every time
		' "worker.ReportProgress(0, state)" is
		' called in FtpDownload.getFileFromFtpServerBW

		Dim state As CurrentState = e.UserState
		Console.WriteLine(state.status & " (" & state.percentDone & "%)")
	End Sub

	Private Sub BackgroundWorker2_RunWorkerCompleted(sender As System.Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
		If Not e.Error Is Nothing Then
			'Handle the errors here.
			'Write error message to console and/or
			'a log file.
			'
			'if using a Form, add a statusStrip. 
			'Then update StatusLabel here.

			Console.WriteLine("Error: " + e.Error.Message)
		ElseIf e.Cancelled = True Then
			'If using a Form, add a statusStrip. 
			'Then update StatusLabel here.
			Console.WriteLine("Cancelled by user.")
		Else
			'If using a Form, add a statusStrip. 
			'Then update StatusLabel here.
			Console.WriteLine("Successfully completed.")
		End If
	End Sub



End Module


