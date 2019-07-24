'-------------------------------------------------------------------------------
' MUSTER.DataAccess.LustEventActivityDB
'   Provides the means for marshalling LustEventActivity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC     03/02/05    Original class definition.
'  1.1        MNR     04/22/05    Added condition to check if value is 0, 
'                                   returns new info/collection instead of accessing db
'                                   as there are no records with primary id 0
'
' Function                  Description
' GetAllInfo()      Returns an LustActivityCollection containing all LustActivity objects in the repository
' DBGetByID(ID)     Returns an LustActivity object indicated by int arg ID
' DBGetDS(SQL)      Returns a resultant Dataset by running query specified by the string arg SQL
' Put(LustEvent)       Saves the LustActivity passed as an argument, to the DB
'-------------------------------------------------------------------------------
'
'

Imports System.Data.SqlClient
Imports Utils.DBUtils
Namespace MUSTER.DataAccess
    Public Class LustEventActivityDB
        Private _strConn As String
        Private MusterException As Exceptions.MusterExceptions = New MUSTER.Exceptions.MusterExceptions

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

        Public Function DBGetByID(ByVal nVal As Int64) As MUSTER.Info.LustActivityInfo
            ' #Region "XDEOperation" ' Begin Template Expansion{7BAED731-2749-4C9B-9830-48689EE8C96C}
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nVal = 0 Then
                    Return New MUSTER.Info.LustActivityInfo
                End If
                strSQL = "spGetTecEventActivity"
                strVal = nVal

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                '
                ' Repeat the following line as many times as necessary
                '
                Params("@EVENT_ACTIVITY_ID").Value = nVal

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.LustActivityInfo(AltIsDBNull(drSet.Item("Event_Activity_ID"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_ID"), 0), _
                            AltIsDBNull(drSet.Item("START_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("1ST_GWS_BELOW"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("2ND_GWS_BELOW"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("TECH_COMPLETED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("CLOSED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("ACTIVITY_TYPE_ID"), 0), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DELETED"), 0), _
                            AltIsDBNull(drSet.Item("REM_SYSTEM_ID"), 0) _
                    )

                Else

                    Return New MUSTER.Info.LustActivityInfo
                End If
            Catch ex As Exception

                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
            ' #End Region ' XDEOperation End Template Expansion{7BAED731-2749-4C9B-9830-48689EE8C96C}
        End Function

        Public Function DBGetDS(ByVal strSQL As String) As DataSet
            ' #Region "XDEOperation" ' Begin Template Expansion{074F9FF7-4DD4-4F41-A9C1-F134E46BC49B}
            Dim dsData As DataSet
            Try
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            ' #End Region ' XDEOperation End Template Expansion{074F9FF7-4DD4-4F41-A9C1-F134E46BC49B}
        End Function
        Public Function GetAllInfo() As MUSTER.Info.LustActivityCollection
            ' #Region "XDEOperation" ' Begin Template Expansion{995B7456-17FF-498E-A293-B0175978420B}
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetTecEventActivity"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                '
                ' Repeat the following line as many times as necessary
                '
                Params("@EVENT_ACTIVITY_ID").Value = String.Empty

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colEntities As New MUSTER.Info.LustActivityCollection
                While drSet.Read
                    '
                    ' Modify following statement as necessary to generate new info object
                    '
                    Dim oLustActivityInfo As New MUSTER.Info.LustActivityInfo(AltIsDBNull(drSet.Item("Event_Activity_ID"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_ID"), 0), _
                            AltIsDBNull(drSet.Item("START_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("[1ST_GWS_BELOW]"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("[2ND_GWS_BELOW]"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("TECH_COMPLETED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("CLOSED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("ACTIVITY_TYPE_ID"), 0), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DELETED"), 0), _
                            AltIsDBNull(drSet.Item("REM_SYSTEM_ID"), 0) _
                    )
                    colEntities.Add(oLustActivityInfo)
                End While

                Return colEntities
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
            ' #End Region ' XDEOperation End Template Expansion{995B7456-17FF-498E-A293-B0175978420B}
        End Function
        Public Sub Put(ByRef oLustActivity As MUSTER.Info.LustActivityInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolHasOpenDocs As Boolean = False)
            ' #Region "XDEOperation" ' Begin Template Expansion{CE44BE18-660C-4873-BF52-C01555F731C8}
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.LustActivity, Integer))) Then
                    returnVal = "You do not have rights to save LustEvent Activity."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutTecEventActivity")

                'If activity is being initially added skip this otherwise when adding the activity and 
                'also selecting a technically completed date the closed date will be populated.
                If oLustActivity.ActivityID > 0 Or oLustActivity.ActivityID < -100 Then
                    If oLustActivity.Completed = "01/01/0001" And oLustActivity.Closed <> "01/01/0001" Then
                        oLustActivity.Completed = oLustActivity.Closed
                    ElseIf oLustActivity.Completed <> "01/01/0001" And oLustActivity.Closed = "01/01/0001" And Not bolHasOpenDocs Then
                        oLustActivity.Closed = oLustActivity.Completed
                    End If
                End If
                If oLustActivity.ActivityID = 0 Or oLustActivity.ActivityID = -1 Then
                    Params(0).Value = System.DBNull.Value
                Else
                    Params(0).Value = oLustActivity.ActivityID
                End If
                If oLustActivity.EventID = 0 Or oLustActivity.EventID = -1 Then
                    Params(1).Value = System.DBNull.Value
                Else
                    Params(1).Value = oLustActivity.EventID
                End If

                Params(2).Value = IIFIsDateNull(oLustActivity.Started, DBNull.Value)
                Params(3).Value = IIFIsDateNull(oLustActivity.First_GWS_Below, DBNull.Value)
                Params(4).Value = IIFIsDateNull(oLustActivity.Second_GWS_Below, DBNull.Value)
                Params(5).Value = IIFIsDateNull(oLustActivity.Completed, DBNull.Value)
                Params(6).Value = IIFIsDateNull(oLustActivity.Closed, DBNull.Value)

                If oLustActivity.Type = 0 Then
                    Params(7).Value = System.DBNull.Value
                Else
                    Params(7).Value = oLustActivity.Type
                End If

                Params(8).Value = oLustActivity.Deleted
                Params(9).Value = oLustActivity.RemSystemID
                If oLustActivity.ActivityID <= 0 Then
                    Params(10).Value = oLustActivity.CreatedBy
                Else
                    Params(10).Value = oLustActivity.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutTecEventActivity", Params)

                'Perform check for New ID and assign, if necessary
                If Params(0).Value <> oLustActivity.ActivityID Then
                    oLustActivity.ActivityID = Params(0).Value
                End If


            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            ' #End Region ' XDEOperation End Template Expansion{CE44BE18-660C-4873-BF52-C01555F731C8}
        End Sub
    End Class
End Namespace



