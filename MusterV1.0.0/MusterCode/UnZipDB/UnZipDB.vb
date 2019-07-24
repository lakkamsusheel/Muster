Imports ICSharpCode.SharpZipLib.Zip
Imports System.Data.SqlClient

Module UnZipDB
    Sub Main()
        Dim cmdArgs() As String = Command.Split(",")
        Dim strServerName As String = String.Empty
        Dim strDBName As String = "Muster_Prd"
        Dim strZipDirName As String = String.Empty
        Dim strDateofBackup As String = ""
        Dim strTimeofBackup As String = "0030"
        Dim strDBUserID As String = "sa"
        Dim strDBUserPass As String = "4b3dD60w"

        For i As Integer = LBound(cmdArgs) To UBound(cmdArgs)
            Select Case i
                Case 0
                    strServerName = cmdArgs(i).Trim
                Case 1
                    strDBName = cmdArgs(i).Trim
                Case 2
                    strZipDirName = cmdArgs(i).Trim
                Case 3
                    strDateofBackup = cmdArgs(i).Trim
                Case 4
                    strTimeofBackup = cmdArgs(i).Trim
                Case 5
                    strDBUserID = cmdArgs(i).Trim
                Case 6
                    strDBUserPass = cmdArgs(i).Trim
            End Select
            Console.WriteLine(i.ToString + ":" + cmdArgs(i).Trim)
        Next

        If strServerName = String.Empty Then
            strServerName = System.Net.Dns.GetHostByName("localhost").HostName()
            Console.WriteLine(strServerName)
        End If
        Dim strSQLPath, strFileDirName As String
        If strZipDirName = String.Empty Then
            Console.WriteLine("Invalid ZipDirName: " + strZipDirName)
            Exit Sub
        ElseIf strZipDirName.EndsWith(IO.Path.DirectorySeparatorChar) Then
            strZipDirName = strZipDirName.TrimEnd(IO.Path.DirectorySeparatorChar)
        End If

        Console.WriteLine("ZipDirName: " + strZipDirName)

        Dim strZipFileName As String = strDBName + "_db_" + IIf(strDateofBackup = String.Empty, Now.Year.ToString + IIf(Now.Month < 10, "0", "") + Now.Month.ToString + IIf(Now.Day < 10, "0", "") + Now.Day.ToString, strDateofBackup) + strTimeofBackup + ".ZIP"
        Dim strFileName As String = strDBName + "_db_" + IIf(strDateofBackup = String.Empty, Now.Year.ToString + IIf(Now.Month < 10, "0", "") + Now.Month.ToString + IIf(Now.Day < 10, "0", "") + Now.Day.ToString, strDateofBackup) + strTimeofBackup + ".BAK"

        Console.WriteLine("ZipFileName: " + strZipFileName)
        Console.WriteLine("FileName: " + strFileName)

        Console.WriteLine("Check exists: " + strZipDirName + IO.Path.DirectorySeparatorChar + strFileName)

        Dim bolZipExists As Boolean = True
        Dim bolBAKFileExists As Boolean = False

        If Not IO.File.Exists(strZipDirName + IO.Path.DirectorySeparatorChar + strZipFileName) Then
            Console.WriteLine("Zip File does not Exist: " + strZipDirName + IO.Path.DirectorySeparatorChar + strZipFileName)
            bolZipExists = False
        End If

        If IO.File.Exists(strZipDirName + IO.Path.DirectorySeparatorChar + strFileName) Then
            Console.WriteLine("File Exists: " + strZipDirName + IO.Path.DirectorySeparatorChar + strFileName)
            bolBAKFileExists = True

            '     IO.File.Delete(strZipDirName + IO.Path.DirectorySeparatorChar + strZipFileName)
            '     bolBAKFileExists = False


        End If

        If bolZipExists And Not bolBAKFileExists Then
            Console.WriteLine("Extracting zip")

            Dim unZip As FastZip
            Try
                unZip = New FastZip(New FastZipEvents)
                unZip.ExtractZip(strZipDirName + IO.Path.DirectorySeparatorChar + strZipFileName, strZipDirName, "")
                bolBAKFileExists = True
            Finally
                unZip = Nothing
            End Try
        End If

        If bolZipExists Then
            IO.File.Delete(strZipDirName + IO.Path.DirectorySeparatorChar + strZipFileName)
            Console.WriteLine("File: " + strZipFileName + " deleted.")
        End If
        Try
            strSQLPath = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE", False).OpenSubKey("Microsoft", False).OpenSubKey("MSSQLServer", False).OpenSubKey("Setup", False).GetValue("SQLDataRoot")
        Catch ex As Exception
            ' write entry to error log - sql server path not found in registry
            Console.WriteLine(ex.Message + vbCrLf + "SQL Server Path not found in registry for: " + strServerName)
            Exit Sub
        End Try

        Console.WriteLine("SQL Server PATH: " + strSQLPath)

        ' restore db
        If bolBAKFileExists Then
            Dim ds As DataSet
            Dim sqlConn As SqlConnection
            Dim sqlCmd As New SqlCommand

            Try
                ds = New DataSet
                sqlConn = New SqlConnection("Data Source=" + strServerName + ";Initial Catalog=master;User ID=" + strDBUserID + ";Password=" + strDBUserPass + ";")
                Dim strSQL As String = "USE MASTER; " + _
                                        "exec sp_KillAllUsersOnMusterPrd; " + _
                                        "IF NOT EXISTS(SELECT * FROM SYSDATABASES WHERE NAME = 'Muster_Prd') " + _
                                        "BEGIN CREATE DATABASE [Muster_Prd] " + _
                                        "END; " + _
                                        "RESTORE DATABASE [Muster_Prd] " + _
                                        "FROM DISK = '" + strZipDirName + IO.Path.DirectorySeparatorChar + strFileName + "' " + _
                                        "WITH  FILE = 1, NOUNLOAD, STATS = 10, RECOVERY, REPLACE, " + _
                                        "MOVE 'Muster_Prd_Data' TO '" + strSQLPath + "\data\Muster_Prd_Data.MDF' "

                'if the provided connection is not open, we will open it
                If sqlConn.State <> ConnectionState.Open Then
                    sqlConn.Open()
                End If

                'Set the Command Timeout to 100 minutes
                sqlCmd.CommandTimeout = 6000

                'associate the connection with the command
                sqlCmd.Connection = sqlConn

                'set the command text (stored procedure name or SQL statement)
                sqlCmd.CommandText = strSQL

                'set the command type
                sqlCmd.CommandType = CommandType.Text

                sqlCmd.ExecuteNonQuery()
                sqlCmd.Parameters.Clear()
            Catch ex As Exception
                ' write entry to error log - ex.message
                Console.WriteLine(ex.Message + vbCrLf + "when restoring db: " + strServerName + " - " + strDBName)
                Exit Sub
            Finally
                If Not sqlCmd Is Nothing Then
                    sqlCmd.Dispose()
                    sqlCmd = Nothing
                End If
                If Not sqlConn Is Nothing Then
                    sqlConn.Close()
                    sqlConn = Nothing
                End If
            End Try
        End If

        If IO.Directory.Exists(strZipDirName) Then
            ' delete files in the zip directory - data cleanup
            Dim strFiles() As String = IO.Directory.GetFiles(strZipDirName)
            For Each strFile As String In strFiles
                If strFile.EndsWith(strFileName) Then
                    Console.WriteLine("Skip Deleting file: " + strFile)
                Else
                    Console.WriteLine("Deleting file: " + strFile)
                    IO.File.Delete(strFile)
                End If
            Next
        End If
    End Sub

End Module
