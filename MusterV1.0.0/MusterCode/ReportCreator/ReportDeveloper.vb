Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.CrystalReports
Imports CrystalDecisions.ReportSource
Imports CrystalDecisions.Windows.Forms
Imports CrystalDecisions.Shared
Imports System


Public Class ReportDeveloper

    Public Class ReportDeveloperDB

        Public DBSource
        Public userID
        Public Password
        Public Server

        Private Function getValue(ByVal key As String, ByVal str As String)

            str = str.Replace(" =", "=").Replace("= ", "=")
            key = String.Format("{0}=", key)

            str = str.Substring(str.ToUpper.IndexOf(key.ToUpper) + key.Length)

            If str.IndexOf(";") > -1 Then
                str = str.Substring(0, str.IndexOf(";"))
            End If

            Return str
        End Function


        Sub New(ByVal connStr As String)

            Server = getValue("Data Source", connStr)
            DBSource = getValue("initial catalog", connStr)
            userID = getValue("user id", connStr)
            Password = getValue("password", connStr)

        End Sub

    End Class

    Private strFileName As String = String.Empty
    Private rDocument As ReportDocument
    Private strConn As String

    Sub New(ByVal reportFile As String, ByVal connStr As String)

        strFileName = reportFile
        strConn = connStr

    End Sub

    Sub dispose()

        If Not rDocument Is Nothing Then
            rDocument.Dispose()
        End If

    End Sub



    Public Sub InitializeReport()

        Dim CrConnInfo As ConnectionInfo
        Dim CrTableLogon As TableLogOnInfo
        Dim crTable As CrystalDecisions.CrystalReports.Engine.Table
        Dim crTables As CrystalDecisions.CrystalReports.Engine.Tables
        Dim crDatabase As CrystalDecisions.CrystalReports.Engine.Database
        Dim reportDbAccess As ReportDeveloperDB


        Try
            If System.IO.File.Exists(strFileName) Then

                rDocument = New ReportDocument
                rDocument.Load(strFileName) ', OpenReportMethod.OpenReportByTempCopy

            Else
                Throw New Exception(String.Format("report Template {0} not found. ", strFileName))

            End If

        Catch ex As Exception

            Throw New Exception("Error Loading report. {0}", ex)
            Exit Sub

        End Try

        Try
            crDatabase = rDocument.Database

            crTables = crDatabase.Tables
            CrConnInfo = New ConnectionInfo

            reportDbAccess = New ReportDeveloperDB(strConn)

            With CrConnInfo

                Try
                    .ServerName = reportDbAccess.Server
                    .DatabaseName = reportDbAccess.DBSource
                    .UserID = reportDbAccess.userID
                    .Password = reportDbAccess.Password

                Catch ex As Exception

                    Throw New Exception("Error initializing data Source to report", ex)
                End Try


            End With


            For Each crTable In crTables
                CrTableLogon = crTable.LogOnInfo
                CrTableLogon.ConnectionInfo = CrConnInfo

                crTable.ApplyLogOnInfo(CrTableLogon)

                If crTable.Location = "Command" Then
                    crTable.Location = CrConnInfo.DatabaseName

                Else
                    crTable.Location = CrConnInfo.DatabaseName & ".dbo." & _
                        crTable.Location.Substring(crTable.Location.LastIndexOf(".") + 1)
                End If

            Next

            'set logon info to subreports 
            Dim crSections As Sections
            Dim crSection As Section

            crSections = rDocument.ReportDefinition.Sections

            Dim crReportObjects As ReportObjects
            Dim crReportObject As ReportObject
            Dim crSubreportObject As SubreportObject
            Dim subRepDoc As New ReportDocument
            Dim nSubReportParamCount As Integer = 0

            For Each crSection In crSections

                crReportObjects = crSection.ReportObjects

                For Each crReportObject In crReportObjects

                    If crReportObject.Kind = ReportObjectKind.SubreportObject Then

                        'If you find a subreport, typecast the reportobject to a subreport object 
                        crSubreportObject = CType(crReportObject, SubreportObject)

                        'Open the subreport 
                        subRepDoc = crSubreportObject.OpenSubreport(crSubreportObject.SubreportName)
                        nSubReportParamCount = nSubReportParamCount + subRepDoc.DataDefinition.ParameterFields.Count

                        crDatabase = subRepDoc.Database
                        crTables = crDatabase.Tables

                        'Loop through each table and set the connection info 
                        'Pass the connection info to the logoninfo object then apply the 
                        'logoninfo to the subreport 
                        CrConnInfo = New ConnectionInfo

                        With CrConnInfo

                            Try
                                .ServerName = reportDbAccess.Server
                                .DatabaseName = reportDbAccess.DBSource
                                .UserID = reportDbAccess.userID
                                .Password = reportDbAccess.Password

                            Catch ex As Exception
                                Throw New Exception("Error initializing data Source to Sub report", ex)
                            End Try

                        End With

                        For Each crTable In crTables
                            CrTableLogon = crTable.LogOnInfo
                            CrTableLogon.ConnectionInfo = CrConnInfo
                            crTable.ApplyLogOnInfo(CrTableLogon)
                            If crTable.Location = "Command" Then
                                crTable.Location = CrConnInfo.DatabaseName
                            Else
                                crTable.Location = CrConnInfo.DatabaseName & ".dbo." & _
                                    crTable.Location.Substring(crTable.Location.LastIndexOf(".") + 1)
                            End If
                        Next

                    End If
                Next
            Next

            Console.WriteLine("report has Created and set to user database")

#If DEBUG Then
            Debugger.Log(0, "Application", "report has been Created")
#End If

        Catch ex As Exception

            Throw New Exception(String.Format("{0}{1}{2}", "Error Detected: ", vbCrLf, ex.Message))
        End Try
    End Sub

    Sub PushParametersInReport(ByVal params As ArrayList)

    End Sub

    Sub ExportFile()

        Dim nLen As Integer
        Dim strPhysicalPath As String
        Dim fInfo As New System.IO.FileInfo(strFileName)
        nLen = (fInfo.Name.Length - 4)
        Dim strPdfFileName As String = fInfo.Name.Remove(nLen, 4)
        Dim strDirName As String = fInfo.DirectoryName
        Dim CrystalExportOptions As ExportOptions
        Dim CrystalDiskFileDestinationOptions As DiskFileDestinationOptions

        Try

            CrystalDiskFileDestinationOptions = New DiskFileDestinationOptions

            strPhysicalPath = String.Format("{0}\{1}.PDF", strDirName, strPdfFileName)

            CrystalDiskFileDestinationOptions.DiskFileName = strPhysicalPath

            CrystalExportOptions = rDocument.ExportOptions

            With CrystalExportOptions

                .DestinationOptions = CrystalDiskFileDestinationOptions
                .ExportDestinationType = ExportDestinationType.DiskFile
                .ExportFormatType = ExportFormatType.PortableDocFormat

            End With

            rDocument.Export()

            Console.WriteLine("report has been exported")

#If DEBUG Then
            Debugger.Log(0, "Application", "report has been Exported")
#End If

        Catch ex As Exception
            Throw New Exception("Error Exporting report: {0}", ex)
        End Try


        fInfo = Nothing
        CrystalDiskFileDestinationOptions = Nothing


    End Sub

End Class
