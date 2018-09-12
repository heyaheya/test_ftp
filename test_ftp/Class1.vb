' FtpDownload.vb
'
' Subject: Download file using ftp.
'
' Description: This class contains the majority of the code
' that downloads a file from an ftp server. It searches
' for the newest file, and then downloads it.
'
' getFileFromFtpServerBW is for use with a BackgroundWorker
' and is the preferred way to download the file.
'
' getFileFromFtpServer is for use without a BackgroundWorker--
' it's use is not recommended.
'
' Notes: Values may be set using the constructor or the Properties.
'        _filenamePattern is already defined in this file and isn't
'        listed in the constructor. If necessary, it can be changed
'        using the Property "FilenamePattern"
'
' Written by:   cgeier 07/18/2014
' Modified by:  cgeier 07/19/2014    Changed getFileFromFtpServerBW
'                                    from a Sub to a Function.
'
'                                    Added and removed code from
'                                    getFileFromFtpServerBW so that 
'                                    exceptions are passed to e.Error
'                                    in RunWorkerCompleted for the
'                                    BackgroundWorker. Added code to
'                                    set value of state.percentDone
'           
'                                    Added code in getFileFromFtpServer
'                                    to reduce the number of times
'                                    Application.DoEvents is called.
'                                    Added exception handling.
'                                    Added code to show percent done.
'
'                                    Added additional comments.

Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Net
Imports System.Threading
Imports System.ComponentModel

Public Class CurrentState
	Public status As String = String.Empty
	Public percentDone As Integer = 0

	'add any additional variables you want
	'to pass back to form
End Class

Public Class FtpDownload

	'Private _filenamePattern As String = "[-ldrwx]{10}\s+\d\s+\d\s+\d\s+\d+\s+(?<monthName>[A-Z][a-z]+)\s+(?<monthNumber>\d+)\s+(?<timeOfDay>\d{2}:\d{2})\s+(?<filenameDateYear>\d{4})-(?<filenameDateMonth>\d{2})-(?<filenameDateDay>\d{2})-(?<filenameTimePart>\d{4}).txt.tgz.*"
	'Private _filenamePattern As String = "[-ldrwx]{10}\s+\d\s+\d\s+\d\s+\d+\s+(?<monthName>[A-Z][a-z]+)\s+(?<monthNumber>\d+)\s+(?<timeOfDay>\d{2}:\d{2})\s+(?<filename>\d{4}-\d{2}-\d{2}-\d{4}.txt.tgz).*"
	'Private _filenamePattern As String = ".*(?<monthName>[A-Z][a-z]+)\s+(?<monthNumber>\d+)\s+(?<timeOfDay>\d{2}:\d{2})\s+(?<filename>\d{4}-\d{2}-\d{2}-\d{4}.txt.tgz).*"

	'works
	'Private _filenamePattern As String = ".*(?<filename>\d{4}-\d{2}-\d{2}-\d{4}.txt.tgz).*"

	'test
	Private _filenamePattern As String = ".*(?<filename>samba-\d.\d.\d+.tar.gz).*"

	Private _ftpRequest As FtpWebRequest = Nothing

	Private _downloadDirectory As String = String.Empty
	Private _ftpUrl As String = String.Empty
	Private _password As String = "user@name.com"
	Private _username As String = "anonymous"


	'Enum
	Private Enum FtpMethod
		ListFiles = 0
		Download = 1
	End Enum

	'Properties
	Public Property DownloadDirectory As String
		Get
			Return _downloadDirectory
		End Get

		Set(value As String)
			_downloadDirectory = value
		End Set
	End Property 'downloadDirectory

	Public Property FilenamePattern As String
		Get
			Return _filenamePattern
		End Get

		Set(value As String)
			_filenamePattern = value
		End Set
	End Property 'filenamePattern

	Public Property FtpUrl As String
		Get
			Return _ftpUrl
		End Get

		Set(value As String)
			_ftpUrl = value
		End Set
	End Property 'ftpUrl

	Public Property Password As String
		Get
			Return _password
		End Get

		Set(value As String)
			_password = value
		End Set
	End Property 'password

	Public Property Username As String
		Get
			Return _username
		End Get

		Set(value As String)
			_username = value
		End Set
	End Property 'username

	'constructor
	Public Sub New()

	End Sub

	Public Sub New(ByVal ftpUrl As String,
					ByVal username As String,
					ByVal password As String,
					ByVal downloadDirectory As String)

		_ftpUrl = ftpUrl
		_username = username
		_password = password
		_downloadDirectory = downloadDirectory

	End Sub

	Private Function getNewestFilename(ByVal directoryData As String) As String

		Dim output As String = String.Empty
		Dim lastFilename As String = String.Empty
		Dim newestFilename As String = String.Empty

		'create new instance of RegEx
		Dim myRegex As New Regex(_filenamePattern)

		'create new instance of Match
		Dim myMatch As Match = myRegex.Match(directoryData)

		'keep looking for matches until no more are found
		While myMatch.Success

			'get groups
			Dim myGroups As GroupCollection = myMatch.Groups

			'For Each groupName As String In myRegex.GetGroupNames
			'output += String.Format("Group: '{0}' Value: {1}", groupName, myGroups(groupName).Value)
			'output += System.Environment.NewLine
			'Next

			'Console.WriteLine(output)

			lastFilename = myGroups("filename").Value

			If Not String.IsNullOrEmpty(newestFilename) Then
				If lastFilename > newestFilename Then
					newestFilename = lastFilename
				End If
			Else
				newestFilename = myGroups("filename").Value
			End If

			'Console.WriteLine("lastFilename: " + lastFilename + " newestFilename: " + newestFilename)


			'get next match
			myMatch = myMatch.NextMatch()
		End While

		If String.IsNullOrEmpty(newestFilename) Then
			newestFilename = "No filenames found using specified FilenamePattern."
		End If

		Return newestFilename
	End Function

	Private Function CreateFtpWebRequest(ByVal ftpUrl As String, ByVal username As String, ByVal password As String, ByVal ftpMethod As FtpMethod, ByVal keepAlive As Boolean) As Stream
		'Don't place any exception handling here
		'the error will get passed to the calling
		'sub/function

		'defined as Private
		'_ftpRequest = DirectCast(WebRequest.Create(New Uri(ftpUrl)), FtpWebRequest)
		_ftpRequest = WebRequest.Create(New Uri(ftpUrl))


		'either download the file or
		'list the files in the directory
		If ftpMethod = 0 Then
			'list files in directory
			_ftpRequest.Method = WebRequestMethods.Ftp.ListDirectoryDetails
		ElseIf ftpMethod = 1 Then
			'download file
			_ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile
		End If

		'set username and password
		_ftpRequest.Credentials = New NetworkCredential(username, password)

		'return Stream
		Return _ftpRequest.GetResponse().GetResponseStream

	End Function

	Public Function getFileFromFtpServer() As String
		Dim loopCount As Integer = 0

		'set bufferSize to 256 kb
		Dim bufferSize As Integer = 256 * 1024

		'allocate buffer of bufferSize
		Dim buffer(bufferSize) As Byte

		'number of bytes read to buffer
		Dim bytesIn As Integer = 0


		Dim ftpResponseStream As Stream = Nothing
		Dim ftpReader As StreamReader = Nothing
		Dim directoryListing As String = String.Empty
		Dim filename As String = String.Empty
		Dim fqFilename As String = String.Empty
		Dim fileLength As Integer = 0
		Dim output As IO.Stream = Nothing
		Dim outputFilename As String = String.Empty
		Dim percentDoneDbl As Double = 0.0
		Dim percentDoneInt As Integer = 0

		'total bytes received
		Dim totalBytesIn As Integer = 0

		Dim errMsg As String = String.Empty

		If String.IsNullOrEmpty(_ftpUrl) Then
			errMsg = "FtpUrl not set."
			Throw New Exception(errMsg)
		End If

		If String.IsNullOrEmpty(_downloadDirectory) Then
			errMsg = "DownloadDirectory not set."
			Throw New Exception(errMsg)
		End If

		Try

			'create new request to get list of files
			'FtpMethod is an Enum created above
			ftpResponseStream = CreateFtpWebRequest(FtpUrl, Username, Password, FtpMethod.ListFiles, False)

			ftpReader = New StreamReader(ftpResponseStream)

			'get list of files
			directoryListing = ftpReader.ReadToEnd

			'close StreamReader
			ftpReader.Close()

			'close Stream
			ftpResponseStream.Close()

			'get newest filename
			filename = getNewestFilename(directoryListing)

			If filename.StartsWith("No filenames found") Then
				errMsg = filename
			End If

			If filename.Length > 0 And Not filename.StartsWith("No filenames found") Then

				'need to append filename to ftpUrl
				'in order to download the file

				If _ftpUrl.EndsWith("/") Then
					fqFilename = FtpUrl + filename
				Else
					fqFilename = FtpUrl + "/" + filename
				End If

				Console.WriteLine("fqFilename: " & fqFilename)

				'create new request to download fqFilename
				'FtpMethod is an Enum created above
				ftpResponseStream = CreateFtpWebRequest(fqFilename, Username, Password, FtpMethod.Download, True)

				'get file length
				fileLength = _ftpRequest.GetResponse().ContentLength

				Console.WriteLine("file length: " & fileLength)

				'file to save to
				If DownloadDirectory.EndsWith("\") Then
					outputFilename = DownloadDirectory + filename
				Else
					outputFilename = DownloadDirectory + "\" + filename
				End If

				'create filename on local computer
				output = System.IO.File.Create(outputFilename)

				Do
					loopCount += 1

					'read bytes to buffer
					'and get actual # of bytes read
					bytesIn = ftpResponseStream.Read(buffer, 0, bufferSize)

					'Console.WriteLine("bytesIn: " & bytesIn)

					If bytesIn > 0 Then
						'write buffer to file
						output.Write(buffer, 0, bytesIn)

						totalBytesIn += bytesIn

						percentDoneDbl = (Convert.ToDouble(totalBytesIn) / Convert.ToDouble(fileLength)) * 100.0
						percentDoneInt = Convert.ToInt32(percentDoneDbl)

						Console.WriteLine("Downloaded: " & totalBytesIn & " of " & fileLength & " (" & percentDoneInt & "%)")

						'Prevent unresponsiveness.
						'It is better to use the
						'version that uses a BackgroundWorker.
						'When using a BackgroundWorker, 
						'Application.DoEvents() is not needed.
						'
						'Change 5 to an appropriate value for your code
						'this will significantly reduce performance
						'impact that occurs from calling ReportProgress 
						'and checking cancellationPending too frequently.
						'
						'This makes it so that the code inside the if
						'statement only runs every 5th iteration.
						If loopCount Mod 5 = 0 Then
							Application.DoEvents()
						End If

					End If
				Loop Until bytesIn < 1

			Else
				'Console.WriteLine(errMsg)
				Throw New Exception("No filenames found using specified FilenamePattern = " & _filenamePattern)
			End If

		Catch ex As WebException
			Console.WriteLine("Error: [getFileFromFtpServer]: " & ex.Message)
		Catch ex As Exception
			Console.WriteLine("Error: [getFileFromFtpServer]: " & ex.Message)
		Finally
			'make sure ftpReader is closed
			If Not ftpReader Is Nothing Then
				ftpReader.Close()
				ftpReader = Nothing
			End If

			'close output
			If Not output Is Nothing Then
				output.Close()
				output = Nothing
			End If

			'close ftpResponseStream
			If Not ftpResponseStream Is Nothing Then
				ftpResponseStream.Close()
				ftpResponseStream = Nothing
			End If

			buffer = Nothing
			_ftpRequest = Nothing
		End Try

		Return totalBytesIn.ToString()
	End Function

	Public Function getFileFromFtpServerBW(ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs) As String

		Dim state As CurrentState = New CurrentState()
		Dim lastReportedDateTime As DateTime = DateTime.MinValue
		Dim loopCount As Integer = 0

		'set bufferSize to 256 kb
		Dim bufferSize As Integer = 256 * 1024

		'allocate buffer of bufferSize
		Dim buffer(bufferSize) As Byte

		'number of bytes read to buffer
		Dim bytesIn As Integer = 0

		Dim ftpResponseStream As Stream = Nothing
		Dim ftpReader As StreamReader = Nothing
		Dim directoryListing As String = String.Empty
		Dim filename As String = String.Empty
		Dim fqFilename As String = String.Empty
		Dim fileLength As Integer = 0
		Dim output As IO.Stream = Nothing
		Dim outputFilename As String = String.Empty
		Dim percentDoneDbl As Double = 0.0

		'total bytes received
		Dim totalBytesIn As Integer = 0

		Dim errMsg As String = String.Empty

		'check if FtpUrl and DownloadDirectory
		'values have been set
		If String.IsNullOrEmpty(_ftpUrl) Then
			errMsg = "FtpUrl not set."
			e.Cancel = True
		ElseIf String.IsNullOrEmpty(_downloadDirectory) Then
			errMsg = "DownloadDirectory not set."
			e.Cancel = True
		End If

		'exit if FtpUrl and/or DownloadDirectory
		'values have not been set
		If e.Cancel = True Then
			state.status = errMsg

			'report progress back to form
			worker.ReportProgress(0, state)

			'update last reported time
			lastReportedDateTime = DateTime.Now

			'if debugging, the debugger will 
			'stop here. Press F5 to continue
			'or Click "Debug", then select "Continue".
			'
			'This throws an exception that
			'will end up in e.Error in
			'BackgroundWorker2_RunWorkerCompleted
			Throw New Exception(errMsg)

			'exit function
			Return errMsg
		End If

		Try

			'create new request to get list of files
			'FtpMethod is an Enum created above
			ftpResponseStream = CreateFtpWebRequest(FtpUrl, Username, Password, FtpMethod.ListFiles, False)

			ftpReader = New StreamReader(ftpResponseStream)

			'get list of files
			directoryListing = ftpReader.ReadToEnd

			'close ftpReader
			ftpReader.Close()
			ftpReader = Nothing

			'close ftpResponseStream
			ftpResponseStream.Close()
			ftpResponseStream = Nothing

			'get newest filename
			filename = getNewestFilename(directoryListing)

			If filename.StartsWith("No filenames found") Then
				errMsg = filename
			End If

			If filename.Length > 0 And Not filename.StartsWith("No filenames found") Then

				'need to append filename to ftpUrl
				'in order to download the file
				fqFilename = FtpUrl + "/" + filename

				'Console.WriteLine("fqFilename: " & fqFilename)

				'create new request to download fqFilename
				'FtpMethod is an Enum created above
				ftpResponseStream = CreateFtpWebRequest(fqFilename, Username, Password, FtpMethod.Download, True)

				'get file length
				fileLength = _ftpRequest.GetResponse().ContentLength

				'Console.WriteLine("file length: " & fileLength)

				'file to save to
				If DownloadDirectory.EndsWith("\") Then
					outputFilename = DownloadDirectory + filename
				Else
					outputFilename = DownloadDirectory + "\" + filename
				End If

				'create filename on local computer
				output = System.IO.File.Create(outputFilename)

				Do
					loopCount += 1

					'change 9 to an appropriate value for your code
					'this will significantly reduce performance
					'impact that occurs from calling ReportProgress 
					'and checking cancellationPending too frequently.
					'
					'This makes it so that the code inside the if
					'statement only runs every 9th iteration.
					If loopCount Mod 9 = 0 Then
						If worker.CancellationPending Then
							e.Cancel = True
							Exit Do 'exit loop
						End If

						state.status = "Status: Downloaded " & totalBytesIn & " of " & fileLength
						percentDoneDbl = (Convert.ToDouble(totalBytesIn) / Convert.ToDouble(fileLength)) * 100.0
						state.percentDone = Convert.ToInt32(percentDoneDbl)

						'report progress back to form
						worker.ReportProgress(0, state)

						'update last reported time
						lastReportedDateTime = DateTime.Now
					End If

					'read bytes to buffer
					'and get actual # of bytes read
					bytesIn = ftpResponseStream.Read(buffer, 0, bufferSize)

					'Console.WriteLine("bytesIn: " & bytesIn)

					If bytesIn > 0 Then
						'write buffer to file
						output.Write(buffer, 0, bytesIn)

						totalBytesIn += bytesIn

						'Console.WriteLine("Downloaded: " & totalBytesIn & " of " & fileLength)

					End If
				Loop Until bytesIn < 1

			Else
				'Console.WriteLine(errMsg)
				Throw New Exception("No filenames found using specified FilenamePattern = " & _filenamePattern)
			End If

			If e.Cancel = True Then
				state.status = "Status: Cancelled by user."
			Else
				state.status = "Status: Complete."
				percentDoneDbl = (Convert.ToDouble(totalBytesIn) / Convert.ToDouble(fileLength)) * 100.0
				state.percentDone = Convert.ToInt32(percentDoneDbl)
			End If

			'report progress back to form
			worker.ReportProgress(0, state)

			'update last reported time
			lastReportedDateTime = DateTime.Now

			'don't put any Catch statements here so
			'the errors will end up in e.Error in
			'BackgroundWorker2_RunWorkerCompleted
			'Handle the errors in 
			'BackgroundWorker2_RunWorkerCompleted

		Finally
			'make sure ftpReader is closed
			If Not ftpReader Is Nothing Then
				ftpReader.Close()
				ftpReader = Nothing
			End If

			'close output
			If Not output Is Nothing Then
				output.Close()
				output = Nothing
			End If

			'close ftpResponseStream
			If Not ftpResponseStream Is Nothing Then
				ftpResponseStream.Close()
				ftpResponseStream = Nothing
			End If

			buffer = Nothing
			_ftpRequest = Nothing
		End Try

		Return totalBytesIn.ToString()
	End Function

End Class
