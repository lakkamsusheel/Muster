'-------------------------------------------------------------------------------
' MUSTER.DataAccess.TicklerMessageDB
'   Provides the means for marshalling tickler message state to/from the repository
'
' Copyright (C) 2009 CIBER, Inc.
' All rights reserved.
'
' Release   Initials           Date        Description
'  1.0       Thomas Franey     05/29/09    Original class definition.
'
' Function                  Description
' DBGetByID(msgID)         Returns an LustRemediation Object indicated by Lust Remediation ID
' DBGetDS(SQL)          Returns a resultant Dataset by running query specified by the string arg SQL
' Put(oTicklerMessage)        Saves the LustRemediation passed as an argument, to the DB
'-------------------------------------------------------------------------------
'
'

Imports System.Data.SqlClient

Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class TicklerMessageDB


#Region "Private Member Variables"
        Private _strConn As Object
        Private MusterException As Exceptions.MusterExceptions = New MUSTER.Exceptions.MusterExceptions
#End Region

#Region "ENUMS"

        Public Enum MessageBOOLEnum

            No = 0
            Yes = 1
            BothYesNo = 2

        End Enum
#End Region
#Region "Exposed Methods"
        Public Sub New(Optional ByVal strDBConn As String = "", Optional ByRef MusterXCEP As MUSTER.Exceptions.MusterExceptions = Nothing)
            Try
                If MusterXCEP Is Nothing Then
                    MusterException = New MUSTER.Exceptions.MusterExceptions
                Else
                    MusterException = MusterXCEP
                End If
                If strDBConn = String.Empty Then
                    Dim oCnn As New ConnectionSettings
                    _strConn = oCnn.cnString
                    oCnn = Nothing
                Else
                    _strConn = strDBConn
                End If
            Catch ex As Exception
                Throw ex
            End Try
        End Sub

        Public Function DBGetDS(ByVal strSQL As String) As DataSet
            Dim dsData As DataSet = Nothing
            Try
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            Return dsData

        End Function

        Public Function DbGetTicklerList(ByVal userid As String, Optional ByVal read As MessageBOOLEnum = MessageBOOLEnum.BothYesNo, Optional ByVal completed As MessageBOOLEnum = MessageBOOLEnum.No) As DataTable

            Dim strSQL As String = String.Format("exec spGetTICKLERList '{0}',{1},{2}", userid, Convert.ToInt16(read), Convert.ToInt16(completed))
            Dim dsData As DataSet
            Dim dsTable As DataTable

            Try

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)


                If Not dsData Is Nothing AndAlso dsData.Tables.Count > 0 Then

                    dsTable = dsData.Tables(0)
                    dsData.Dispose()

                    dsTable.Columns.Add("IsNew", (1 = 1).GetType)
                    dsTable.Columns("IsNew").ReadOnly = False

                End If

                Return dsTable

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)

                Throw Ex

                Return Nothing
            End Try


        End Function


        Public Function DbGetTicklerListSent(ByVal userid As String, Optional ByVal read As MessageBOOLEnum = MessageBOOLEnum.BothYesNo, Optional ByVal completed As MessageBOOLEnum = MessageBOOLEnum.No) As DataTable

            Dim strSQL As String = String.Format("exec spGetTICKLERList '{0}',{1},{2},1", userid, Convert.ToInt16(read), Convert.ToInt16(completed))
            Dim dsData As DataSet
            Dim dsTable As DataTable

            Try

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)


                If Not dsData Is Nothing AndAlso dsData.Tables.Count > 0 Then

                    dsTable = dsData.Tables(0)
                    dsData.Dispose()

                    dsTable.Columns.Add("IsNew", (1 = 1).GetType)
                    dsTable.Columns("IsNew").ReadOnly = False

                End If

                Return dsTable

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex

                Return Nothing

            End Try


        End Function

        Public Function DBGetByID(ByVal nID As String, Optional ByVal setRead As Boolean = False, Optional ByVal setCompleted As Boolean = False) As MUSTER.Info.TicklerMessageInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nID = String.Empty Then
                    Return New MUSTER.Info.TicklerMessageInfo
                End If
                strSQL = "spGetTicklerMessage"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@msgID").Value = nID
                Params("@read").Value = IIf(setRead, 1, 0)
                Params("@completed").Value = IIf(setCompleted, 1, 0)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.TicklerMessageInfo(drSet.Item("MsgID"), _
                            AltIsDBNull(drSet.Item("FromID"), String.Empty), _
                            AltIsDBNull(drSet.Item("ToID"), String.Empty), _
                            AltIsDBNull(drSet.Item("MsgRead"), False), _
                            AltIsDBNull(drSet.Item("MsgCompleted"), False), _
                            AltIsDBNull(drSet.Item("IsIssue"), False), _
                            AltIsDBNull(drSet.Item("Subject"), String.Empty), _
                            AltIsDBNull(drSet.Item("Msg"), String.Empty), _
                            AltIsDBNull(drSet.Item("ModuleGroupID"), 0), _
                            AltIsDBNull(drSet.Item("objectID"), String.Empty), _
                            AltIsDBNull(drSet.Item("Keyword"), String.Empty), _
                            AltIsDBNull(drSet.Item("ImageFile"), String.Empty), _
                            AltIsDBNull(drSet.Item("Date_Created"), Now), _
                            AltIsDBNull(drSet.Item("PostDate"), Now), _
                            AltIsDBNull(drSet.Item("DateRead"), Nothing), _
                            AltIsDBNull(drSet.Item("DateCompleted"), Nothing))

                Else

                    Return New MUSTER.Info.TicklerMessageInfo
                End If
            Catch ex As Exception

                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function


        Public Sub Put(ByRef oTicklerMessageInfo As MUSTER.Info.TicklerMessageInfo, ByRef returnVal As String)
            Try

                returnVal = String.Empty

                Dim Params() As SqlParameter
                Dim valID As String = "-1"

                If oTicklerMessageInfo.ID > String.Empty Then
                    valID = oTicklerMessageInfo.ID
                End If

                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutTicklerMessage")

                With oTicklerMessageInfo

                    Params(0).Value = .FromID
                    Params(1).Value = .toID
                    Params(2).Value = .Read
                    Params(3).Value = .Completed
                    Params(4).Value = .IsIssue
                    Params(5).Value = .Subject
                    Params(6).Value = .Message
                    Params(7).Value = .ModuleID
                    Params(8).Value = .ObjectID
                    Params(9).Value = .Keyword
                    Params(10).Value = .ImageFile
                    Params(11).Value = IIf(.PostDate < New Date(1910, 1, 1), Nothing, .PostDate)
                    Params(12).Value = valID


                End With


                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutTicklerMessage", Params)

                oTicklerMessageInfo.ID = valID

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
    End Class
End Namespace
