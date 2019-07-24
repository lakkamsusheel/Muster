Imports Microsoft.ApplicationBlocks.ExceptionManagement
Imports System.Text
Imports System.IO
Imports System.Collections.Specialized
Imports System.Data.SqlClient
Namespace MUSTER.Exceptions
    Public Class MusterExceptions
        Implements IExceptionPublisher
        '-------------------------------------------------------------------------------
        ' MUSTER.MusterExceptions
        '   Provides the interface to publish errors to the database
        '
        ' Copyright (C) 2004 CIBER, Inc.
        ' All rights reserved.
        '
        ' Release   Initials    Date        Description
        '  1.0        EN      10/??/04    Original class definition.
        '  1.1        JC      01/03/05    Replaced dump of sbAddInfo with
        '                                   dump of exception.stacktrace info.
        '  1.2        Manju   07/24/07    Set command / connection to nothing to release resources
        ' Function          Description
        '-------------------------------------------------------------------------------
        '
        ' TODO - update operations and attributes list.
        '
        Private mstrTable As String = "tblException_Log"
        Private LocalUserSettings As Microsoft.Win32.Registry

        Public Sub Publish(ByVal exception As System.Exception, ByVal additionalInfo As System.Collections.Specialized.NameValueCollection, _
        ByVal configSettings As System.Collections.Specialized.NameValueCollection) Implements Microsoft.ApplicationBlocks.ExceptionManagement.IExceptionPublisher.Publish

            Dim conSQLConnection As New SqlConnection
            Dim cmdSQLCommand As New SqlCommand
            Dim sbSQL As New StringBuilder
            Dim sbAddInfo As New StringBuilder
            conSQLConnection.ConnectionString = LocalUserSettings.CurrentUser.GetValue("MusterSQLConnection")
            If Not additionalInfo Is Nothing Then
                For Each i As String In additionalInfo
                    sbAddInfo.AppendFormat("{0}: {1};", i, additionalInfo.Get(i))
                Next
            End If
            conSQLConnection.Open()
            cmdSQLCommand.Connection = conSQLConnection
            sbSQL.AppendFormat("insert into {0}(Message, Source,Module,AddInfo,created_date) values ", mstrTable)
            sbSQL.AppendFormat("('{0}','{1}','{2}','{3}','{4}')", exception.Message, exception.Source, exception.TargetSite, exception.StackTrace, DateTime.Now)
            cmdSQLCommand.CommandText = sbSQL.ToString
            cmdSQLCommand.ExecuteNonQuery()
            conSQLConnection.Close()

            cmdSQLCommand.Dispose()
            conSQLConnection.Dispose()

            If Not sbSQL Is Nothing Then
                sbSQL = Nothing
            End If
            If Not sbAddInfo Is Nothing Then
                sbAddInfo = Nothing
            End If
            If Not cmdSQLCommand Is Nothing Then
                cmdSQLCommand = Nothing
            End If
            If Not conSQLConnection Is Nothing Then
                conSQLConnection = Nothing
            End If
        End Sub
    End Class
End Namespace

