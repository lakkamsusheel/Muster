Module UploadToFTP

    Sub Main()
        Dim cmdArgs() As String = Command.Split(",")
        Dim strServerName As String = "muster.deq.state.ms.us"
        Dim strPort As String = "21"
        Dim strFtpDirName As String = "muster"
        Dim strUserID As String = "anonymous"
        Dim strUserPass As String = ""
        Dim strFilePaths() As Object



        If Not cmdArgs Is Nothing AndAlso cmdArgs.GetUpperBound(0) = "-1" Then

            Console.WriteLine("Please Select a Job parameter")

            End

        End If

        If cmdArgs(0).ToUpper = "BACKUPS" Then

            ' for test purpose only
            ReDim strFilePaths(0)

            strFilePaths(0) = Dir("E:\CIBER\MUSTER\DBSync\*.ZIP")

            strFilePaths(0) = String.Format("{0}{1}", "E:\CIBER\MUSTER\DBSync\", strFilePaths(0))

            Console.WriteLine(String.Format("Found In Directory List: {0}", strFilePaths(0)))

        ElseIf cmdArgs(0).ToUpper = "PROHIBITION" Then

            Dim list As New ArrayList

            Try

                If Dir("E:\CIBER\MUSTER\DBSync\FTPZipDB\Reports\Prohibition*.PDF") <> String.Empty Then

                    list.Add(Dir("E:\CIBER\MUSTER\DBSync\FTPZipDB\Reports\Prohibition*.PDF"))

                    list.Item(list.Count - 1) = String.Format("{0}{1}", "E:\CIBER\MUSTER\DBSync\FTPZipDB\reports\", list.Item(list.Count - 1))

                    Console.WriteLine(String.Format("Found In Directory List: {0}", list.Item(list.Count - 1)))

                    Dim nextFile As Object = Dir()

                    While nextFile <> String.Empty

                        list.Add(nextFile)

                        list.Item(list.Count - 1) = String.Format("{0}{1}", "E:\CIBER\MUSTER\DBSync\FTPZipDB\reports\", list.Item(list.Count - 1))

                        Console.WriteLine(String.Format("Found In Directory List: {0}", list.Item(list.Count - 1)))

                        nextFile = Dir()

                    End While

                    strFilePaths = list.ToArray

                End If

            Catch ex As Exception
                Console.WriteLine(ex.ToString)
            End Try

        ElseIf cmdArgs(0).ToUpper = "PSI" Then

            Dim list As New ArrayList

            Try

                If Dir("E:\CIBER\MUSTER\DBSync\FTPZipDB\Reports\WebPSI.PDF") <> String.Empty Then

                    list.Add(Dir("E:\CIBER\MUSTER\DBSync\FTPZipDB\Reports\WebPSI*.PDF"))

                    list.Item(list.Count - 1) = String.Format("{0}{1}", "E:\CIBER\MUSTER\DBSync\FTPZipDB\reports\", list.Item(list.Count - 1))

                    Console.WriteLine(String.Format("Found In Directory List: {0}", list.Item(list.Count - 1)))

                    Dim nextFile As Object = Dir()

                    While nextFile <> String.Empty

                        list.Add(nextFile)

                        list.Item(list.Count - 1) = String.Format("{0}{1}", "E:\CIBER\MUSTER\DBSync\FTPZipDB\reports\", list.Item(list.Count - 1))

                        Console.WriteLine(String.Format("Found In Directory List: {0}", list.Item(list.Count - 1)))

                        nextFile = Dir()

                    End While

                    strFilePaths = list.ToArray

                End If

            Catch ex As Exception
                Console.WriteLine(ex.ToString)
            End Try


        Else

            Console.WriteLine("Command not recognized.")

            End

        End If


        'strFilePaths(1) = "C:\MDEQ\UploadToFTP\2.txt"

        If strServerName = String.Empty Then
            Console.WriteLine("Invalid Server name")

        Else

            Console.WriteLine("beginning FTP program")

            TestFTP(strServerName, strPort, strFtpDirName, strUserID, strUserPass, strFilePaths)

        End If


    End Sub

    '
    ' Copy and paste the code below into a VB WebForm or WinForm
    '  application and then do the following:
    '
    '       1).  From within the ASP.NET or WinForm or Console app set a
    '            reference to the FTP.dll and BitOperators.dll
    '            files.
    '       2).  At the top of the application code file 
    '            (E.g WebForm1.aspx.vb or Form1.vb) type in
    '               Imports FTP
    '       3).  Compile the application and run.
    '       4).  Have fun.

    Private Sub TestFTP(ByVal serverName As String, ByVal portNum As Integer, ByVal ftpDirName As String, ByVal ftpUser As String, ByVal ftpPass As String, ByVal strFilePaths() As Object)
        Dim ff As clsFTP

        Try
            '-------------------------------------------
            ' OPTION 1
            ' --------
            '
            ' Create an instance of the FTP Class.
            'ff = New clsFTP()

            ' Setup the appropriate properties.
            'ff.RemoteHost = "microsoft"
            'ff.RemoteUser = "ftpuser"
            'ff.RemotePassword = "password"
            '-------------------------------------------

            '-------------------------------------------
            ' OPTION 2
            ' --------
            '  Pass the values into the constructor 
            '  instead.  These can be overridden by simply
            '  setting the appropriate properties on the
            '  instance of the clsFTP Class.
            ff = New clsFTP(serverName, _
                            ftpDirName, _
                            ftpUser, _
                            ftpPass, _
                            portNum)

            ' Attempt to log into the FTP Server.
            If (ff.Login()) Then
                '
                ' Move the to Area1\Section1\Subby1\ directory.
                'ff.ChangeDirectory("Area1")
                'ff.ChangeDirectory("Section1")
                'ff.CreateDirectory("Subby1")
                'ff.ChangeDirectory("Subby1")
                ff.SetBinaryMode(True)

                ' Upload a file.
                For Each strFile As Object In strFilePaths
                    ff.UploadFile(CStr(strFile))
                    Console.WriteLine()
                    Console.WriteLine(CStr(strFile) + " has been uploaded")
                Next

                ' Download a file.
                'ff.DownloadFile("secureapps.pdf", "d:\general\secureapps.pdf")

                ' Remove a file from the FTP Site.
                'If (ff.DeleteFile("secureapps.pdf")) Then
                'Console.WriteLine("File has been removed from FTP Site")
                'MessageBox.Show("File has been removed from FTP Site")
                'Else
                'Console.WriteLine("Unable to remove file from FTP Site.  Message from server: " & ff.MessageString & "<br>")
                'MessageBox.Show("Unable to remove file from FTP Site")
                'End If

                ' Rename a file on the FTP Site.
                'If (ff.RenameFile("secureapps.pdf", "newapp.pdf")) Then
                '    Response.Write("File has been renamed")
                '    MessageBox.Show("File has been renamed")
                'End If

                'ff.ChangeDirectory("..")
                'If (ff.RemoveDirectory("Subby1")) Then
                '    Response.Write("Directory has been removed<br>")
                '    ' MessageBox.Show("Directory has been removed")
                'Else
                '    Response.Write("Unable to remove the directory.  Message from server: " & ff.MessageString & "<br>")
                '    ' MessageBox.Show("Unable to remove the directory.")
                'End If
            End If

        Catch ex As System.Exception
            ' Console App
            Console.WriteLine(ex.Message)
            Console.WriteLine("Message from FTP Server was: " & ff.MessageString)

            ' WinForms
            'Messagebox.Show(ex.Message)
            'MessageBox.show("Message from FTP Server was: " & ff.MessageString)
        Finally
            '
            ' Always close down the connection to ensure that
            '  there are no "stray" Fido's Fetching data.  In
            '  other words, no stray/limbo/not-in-use FTP
            '  connections.
            ff.CloseConnection()
        End Try
    End Sub

End Module
