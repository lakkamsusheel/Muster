Imports ICSharpCode.SharpZipLib.Zip
Imports System.Data.SqlClient

Module ZipDB


    Sub PerformTask()

        Dim cmdArgs() As String = Command.Split(",")
        Dim strServerName As String = "GARD-PROD"
        Dim strDBName As String = "Muster_Prd"
        Dim strZipDirName As String = String.Empty
        Dim strDateofBackup As String = ""
        Dim strTimeofBackup As String = "0030"
        Dim strDBUserID As String = "MusterApp"
        Dim strDBUserPass As String = "8f1-4c9A"

        For i As Integer = LBound(cmdArgs) To UBound(cmdArgs)
            Select Case i
                Case 0
                    ' strServerName = cmdArgs(i).Trim
                Case 1
                    'strDBName = cmdArgs(i).Trim
                Case 2
                    'strZipDirName = cmdArgs(i).Trim
                Case 3
                    'strDateofBackup = cmdArgs(i).Trim
                Case 4
                    'strTimeofBackup = cmdArgs(i).Trim
                Case 5
                    'strDBUserID = cmdArgs(i).Trim
                Case 6
                    'strDBUserPass = cmdArgs(i).Trim
            End Select
            Console.WriteLine(i.ToString + ":" + cmdArgs(i).Trim)
        Next

        If strServerName = String.Empty Then
            strServerName = System.Net.Dns.GetHostByName("localhost").HostName()
            Console.WriteLine(strServerName)
        End If
        Dim strSQLPath, strFileDirName As String
        If strZipDirName = String.Empty Then
            Dim ds As New DataSet
            Dim sqlConn As SqlConnection
            Dim sqlDataAdapter As SqlDataAdapter

            Try
                sqlConn = New SqlConnection("Data Source=" + strServerName + ";Initial Catalog=" + strDBName + ";User ID=" + strDBUserID + ";Password=" + strDBUserPass + ";")
                sqlDataAdapter = New SqlDataAdapter("SELECT PROFILE_VALUE FROM [tblSYS_PROFILE_INFO] WHERE PROFILE_KEY = 'COMMON_PATHS' AND PROFILE_MODIFIER_1 = 'DBSync'", sqlConn)
                sqlDataAdapter.Fill(ds, "PROFILE_VALUE")
                strZipDirName = ds.Tables(0).Rows(0)("PROFILE_VALUE")
            Catch ex As Exception
                'strZipDirName = "\\" + strServerName + "\MUSTER\DBSync"
                ' write entry to error log - ex.message
                Console.WriteLine(ex.Message + vbCrLf + "when getting zip dir name from: " + strServerName + " - " + strDBName)
                Exit Sub
            Finally
                sqlConn.Close()
                sqlDataAdapter = Nothing
                sqlConn = Nothing
            End Try
        ElseIf strZipDirName.EndsWith(IO.Path.DirectorySeparatorChar) Then
            strZipDirName = strZipDirName.TrimEnd(IO.Path.DirectorySeparatorChar)
        End If

        Console.WriteLine("ZipDirName: " + strZipDirName)

        Dim strZipFileName As String = strDBName + "_db_" + IIf(strDateofBackup = String.Empty, Now.Year.ToString + IIf(Now.Month < 10, "0", "") + Now.Month.ToString + IIf(Now.Day < 10, "0", "") + Now.Day.ToString, strDateofBackup) + strTimeofBackup + ".ZIP"
        Dim strFileName As String = strDBName + "_db_" + IIf(strDateofBackup = String.Empty, Now.Year.ToString + IIf(Now.Month < 10, "0", "") + Now.Month.ToString + IIf(Now.Day < 10, "0", "") + Now.Day.ToString, strDateofBackup) + strTimeofBackup + ".BAK"

        Console.WriteLine("ZipFileName: " + strZipFileName)
        Console.WriteLine("FileName: " + strFileName)


        Try
            strSQLPath = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE", False).OpenSubKey("Microsoft", False).OpenSubKey("MSSQLServer", False).OpenSubKey("Setup", False).GetValue("SQLPath")
        Catch ex As Exception
            ' write entry to error log - sql server path not found in registry
            Console.WriteLine(ex.Message + vbCrLf + "SQL Server Path not found in registry for: " + strServerName)
            Exit Sub
        End Try

        strFileDirName += strSQLPath + IO.Path.DirectorySeparatorChar + "BACKUP" + IO.Path.DirectorySeparatorChar + strDBName

        Console.WriteLine("FileDirName: " + strFileDirName)


        If Not IO.Directory.Exists(strZipDirName) Then
            IO.Directory.CreateDirectory(strZipDirName)
            Console.WriteLine("Created Directory: " + strZipDirName)
        End If

        Console.WriteLine("Check exists: " + strZipDirName + IO.Path.DirectorySeparatorChar + strZipFileName)

        If IO.File.Exists(strZipDirName + IO.Path.DirectorySeparatorChar + strZipFileName) Then
            Console.WriteLine("File Exists: " + strZipDirName + IO.Path.DirectorySeparatorChar + strZipFileName)
        End If

        ' delete zip file from sql server backup folder
        Console.WriteLine(String.Format("Deleting zip files in source folder {0}", strFileDirName))
        Dim strZipFiles() As String = IO.Directory.GetFiles(strFileDirName)
        For Each strZipFile As String In strZipFiles
            If Not strZipFile.ToUpper.EndsWith(strZipFileName.ToUpper) And strZipFile.ToUpper.EndsWith(".ZIP") Then
                Console.WriteLine("Directory Clean-up: Deleting file: " + strZipFile)
                IO.File.Delete(strZipFile)
            End If
        Next

        ' delete all the files in the zip directory - data cleanup
        Console.WriteLine(String.Format("Deleting zip files in source folder {0}", strZipDirName))
        strZipFiles = IO.Directory.GetFiles(strZipDirName)
        For Each strZipFile As String In strZipFiles
            If Not strZipFile.ToUpper.EndsWith(strZipFileName.ToUpper) And Not strZipFile.ToUpper.EndsWith(".TXT") Then
                Console.WriteLine("Directory Clean-up: Deleting file: " + strZipFile)
                IO.File.Delete(strZipFile)
            End If
        Next


        If Not IO.Directory.Exists(strFileDirName) Then
            ' write entry to error log - sql server backup path not found
            Console.WriteLine("SQL Server Backup Path not found for: " + strServerName + " - " + strFileDirName)
            Exit Sub
        Else
            If IO.File.Exists(strFileDirName + IO.Path.DirectorySeparatorChar + strZipFileName) Then
                ' write entry to error log - zip exists
                Console.WriteLine("File Exists: " + strFileDirName + IO.Path.DirectorySeparatorChar + strZipFileName)

                'Exit Sub


            End If

            If IO.File.Exists(strFileDirName + IO.Path.DirectorySeparatorChar + strFileName) AndAlso Not IO.File.Exists(strFileDirName + IO.Path.DirectorySeparatorChar + strZipFileName) Then
                IO.Directory.SetCurrentDirectory(strFileDirName)

                'Dim baseOutputStream As IO.Stream = New IO.FileStream(strFileDirName + IO.Path.DirectorySeparatorChar + strZipFileName, IO.FileMode.Create)
                Dim baseOutputStream As IO.Stream = New IO.FileStream(strZipFileName, IO.FileMode.Create)

                Console.WriteLine("Creating zip")

                Dim zipEntry As ZipEntry
                zipEntry = New ZipEntry(strFileName) ' strFileDirName + IO.Path.DirectorySeparatorChar + strFileName
                Dim zip As ZipFile
                zip = New ZipFile(baseOutputStream)
                zip.BeginUpdate()
                zip.Add(strFileName, CompressionMethod.Deflated) ' strFileDirName + IO.Path.DirectorySeparatorChar + strFileName
                zip.CommitUpdate()
                zip.Close()
                baseOutputStream.Close()
                zip = Nothing
                baseOutputStream = Nothing
                zipEntry = Nothing
            ElseIf Not IO.File.Exists(strFileDirName + IO.Path.DirectorySeparatorChar + strFileName) Then
                ' write entry to error log - db backup does not exists
                Console.WriteLine("DB Backup File does not Exists: " + strFileDirName + IO.Path.DirectorySeparatorChar + strFileName)
                Exit Sub

            Else
                Console.WriteLine("Using Old Zip file to copy Over : " + strZipFileName)

            End If
        End If


        Console.WriteLine("Copying zip to destination folder")
        ' copy zip file to destination folder
        IO.File.Copy(strFileDirName + IO.Path.DirectorySeparatorChar + strZipFileName, strZipDirName + IO.Path.DirectorySeparatorChar + strZipFileName, True)



    End Sub
    Sub Main()

        PerformTask()

    End Sub
End Module
