'-------------------------------------------------------------------------------
' MUSTER.DataAccess.LustEventDocumentDB
'   Provides the means for marshalling LustDocument state to/from the repository
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
' GetAllInfo()      Returns an LustDocumentCollection containing all LustDocument objects in the repository
' DBGetByID(ID)     Returns an LustDocument object indicated by int arg ID
' DBGetDS(SQL)      Returns a resultant Dataset by running query specified by the string arg SQL
' Put(LustDocument)       Saves the LustDocument passed as an argument, to the DB
'-------------------------------------------------------------------------------
'
'

Imports System.Data.SqlClient
Imports Utils.DBUtils
Namespace MUSTER.DataAccess
    Public Class LustEventDocumentsDB
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

        Public Function DBGetByID(ByVal nVal As Int64) As MUSTER.Info.LustDocumentInfo
            ' #Region "XDEOperation" ' Begin Template Expansion{7BAED731-2749-4C9B-9830-48689EE8C96C}
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nVal = 0 Then
                    Return New MUSTER.Info.LustDocumentInfo
                End If
                strSQL = "spGetTecEventDocuments"
                'strSQL = "spGetTecEvent"
                strVal = nVal

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                '
                ' Repeat the following line as many times as necessary
                '
                Params("@EVENT_ACTIVITY_DOCUMENT_ID").Value = nVal
                Params("@EVENT_ID").Value = DBNull.Value
                Params("@EVENT_ACTIVITY_ID").Value = DBNull.Value

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.LustDocumentInfo(AltIsDBNull(drSet.Item("EVENT_ACTIVITY_DOCUMENT_ID"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_CLOSED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_SENT_TO_FINANCE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("RECEIVED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV1_RECEIVED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV1_EXTENSION_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV2_RECEIVED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV2_EXTENSION_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("START_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EXTENSION_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DUE_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("ISSUE_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EVENT_ID"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_ACTIVITY_ID"), 0), _
                            AltIsDBNull(drSet.Item("DOCUMENT_PROPERTY_ID"), 0), _
                            AltIsDBNull(drSet.Item("DOCUMENT_ID"), 0), _
                            AltIsDBNull(drSet.Item("DOCUMENT_PROPERTY_ID"), 0), _
                            AltIsDBNull(drSet.Item("Commitment_ID"), 0), _
                            AltIsDBNull(drSet.Item("Paid"), False), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DELETED"), 0) _
                    )

                Else

                    Return New MUSTER.Info.LustDocumentInfo
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
        Public Function DBGetByActivityIDAndDocClass(ByVal ActivityID As Int64, ByVal DocClass As Int64) As MUSTER.Info.LustDocumentInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            Try
                If ActivityID = 0 Or DocClass = 0 Then
                    Return New MUSTER.Info.LustDocumentInfo
                End If
                strSQL = "spGetTecEventDocuments_ByActivityAndDocClass"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                '
                ' Repeat the following line as many times as necessary
                '
                Params("@Document_Property_ID").Value = DocClass
                Params("@EVENT_ACTIVITY_ID").Value = ActivityID

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.LustDocumentInfo(AltIsDBNull(drSet.Item("EVENT_ACTIVITY_DOCUMENT_ID"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_CLOSED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_SENT_TO_FINANCE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("RECEIVED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV1_RECEIVED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV1_EXTENSION_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV2_RECEIVED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV2_EXTENSION_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("START_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EXTENSION_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DUE_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("ISSUE_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EVENT_ID"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_ACTIVITY_ID"), 0), _
                            AltIsDBNull(drSet.Item("DOCUMENT_PROPERTY_ID"), 0), _
                            AltIsDBNull(drSet.Item("DOCUMENT_ID"), 0), _
                            AltIsDBNull(drSet.Item("DOCUMENT_PROPERTY_ID"), 0), _
                            AltIsDBNull(drSet.Item("Commitment_ID"), 0), _
                            AltIsDBNull(drSet.Item("Paid"), False), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DELETED"), 0) _
                    )

                Else

                    Return New MUSTER.Info.LustDocumentInfo
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
        Public Function DBExeNonQuery(ByVal strSQL As String)
            Dim dsData As DataSet
            Try
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.Text, strSQL)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetAllInfo() As MUSTER.Info.LustDocumentCollection
            ' #Region "XDEOperation" ' Begin Template Expansion{995B7456-17FF-498E-A293-B0175978420B}
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetTecEventDocuments"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                '
                ' Repeat the following line as many times as necessary
                '
                Params("@EVENT_ACTIVITY_ID").Value = String.Empty
                Params("@EVENT_ACTIVITY_DOCUMENT_ID").Value = DBNull.Value
                Params("@EVENT_ID").Value = DBNull.Value

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colEntities As New MUSTER.Info.LustDocumentCollection
                While drSet.Read
                    '
                    ' Modify following statement as necessary to generate new info object
                    '
                    Dim oLustDocumentsInfo As New MUSTER.Info.LustDocumentInfo(AltIsDBNull(drSet.Item("EVENT_ACTIVITY_DOCUMENT_ID"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_CLOSED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_SENT_TO_FINANCE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("RECEIVED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV1_RECEIVED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV1_EXTENSION_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV2_RECEIVED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV2_EXTENSION_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("START_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EXTENSION_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DUE_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("ISSUE_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EVENT_ID"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_ACTIVITY_ID"), 0), _
                            AltIsDBNull(drSet.Item("DOCUMENT_PROPERTY_ID"), 0), _
                            AltIsDBNull(drSet.Item("DOCUMENT_ID"), 0), _
                            AltIsDBNull(drSet.Item("DOCUMENT_PROPERTY_ID"), 0), _
                            AltIsDBNull(drSet.Item("Commitment_ID"), 0), _
                            AltIsDBNull(drSet.Item("Paid"), False), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DELETED"), 0) _
                    )
                    colEntities.Add(oLustDocumentsInfo)
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
        Public Function GetAllInfoByEventID(ByVal EventID As Integer) As MUSTER.Info.LustDocumentCollection
            ' #Region "XDEOperation" ' Begin Template Expansion{995B7456-17FF-498E-A293-B0175978420B}
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Dim colEntities As New MUSTER.Info.LustDocumentCollection
            Try
                If EventID = 0 Then
                    Return colEntities
                End If
                strSQL = "spGetTecEventDocuments"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                '
                ' Repeat the following line as many times as necessary
                '
                Params("@EVENT_ID").Value = EventID
                Params("@EVENT_ACTIVITY_DOCUMENT_ID").Value = DBNull.Value
                Params("@EVENT_ACTIVITY_ID").Value = DBNull.Value

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                While drSet.Read
                    '
                    ' Modify following statement as necessary to generate new info object
                    '
                    Dim oLustDocumentsInfo As New MUSTER.Info.LustDocumentInfo(AltIsDBNull(drSet.Item("EVENT_ACTIVITY_DOCUMENT_ID"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_CLOSED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_SENT_TO_FINANCE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("RECEIVED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV1_RECEIVED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV1_EXTENSION_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV2_RECEIVED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV2_EXTENSION_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("START_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EXTENSION_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DUE_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("ISSUE_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EVENT_ID"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_ACTIVITY_ID"), 0), _
                            AltIsDBNull(drSet.Item("DOCUMENT_PROPERTY_ID"), 0), _
                            AltIsDBNull(drSet.Item("DOCUMENT_ID"), 0), _
                            AltIsDBNull(drSet.Item("DOCUMENT_PROPERTY_ID"), 0), _
                            AltIsDBNull(drSet.Item("Commitment_ID"), 0), _
                            AltIsDBNull(drSet.Item("Paid"), False), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DELETED"), 0) _
                    )
                    colEntities.Add(oLustDocumentsInfo)
                End While

                Return colEntities
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
            ' #End Region ' XDEOperation End Template Expansion{995B7456-17FF-498E-A293-B0175978420B}
        End Function
        Public Function GetAllInfoByActivityID(ByVal ActivityID As Integer) As MUSTER.Info.LustDocumentCollection
            ' #Region "XDEOperation" ' Begin Template Expansion{995B7456-17FF-498E-A293-B0175978420B}
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Dim colEntities As New MUSTER.Info.LustDocumentCollection
            Try
                If ActivityID = 0 Then
                    Return colEntities
                End If
                strSQL = "spGetTecEventDocuments"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                '
                ' Repeat the following line as many times as necessary
                '
                Params("@EVENT_ACTIVITY_ID").Value = ActivityID
                Params("@EVENT_ACTIVITY_DOCUMENT_ID").Value = DBNull.Value
                Params("@EVENT_ID").Value = DBNull.Value

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                While drSet.Read
                    '
                    ' Modify following statement as necessary to generate new info object
                    '
                    Dim oLustDocumentsInfo As New MUSTER.Info.LustDocumentInfo(AltIsDBNull(drSet.Item("EVENT_ACTIVITY_DOCUMENT_ID"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_CLOSED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_SENT_TO_FINANCE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("RECEIVED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV1_RECEIVED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV1_EXTENSION_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV2_RECEIVED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV2_EXTENSION_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("START_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EXTENSION_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DUE_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("ISSUE_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EVENT_ID"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_ACTIVITY_ID"), 0), _
                            AltIsDBNull(drSet.Item("DOCUMENT_PROPERTY_ID"), 0), _
                            AltIsDBNull(drSet.Item("DOCUMENT_ID"), 0), _
                            AltIsDBNull(drSet.Item("DOCUMENT_PROPERTY_ID"), 0), _
                            AltIsDBNull(drSet.Item("Commitment_ID"), 0), _
                            AltIsDBNull(drSet.Item("Paid"), False), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DELETED"), 0) _
                    )
                    colEntities.Add(oLustDocumentsInfo)
                End While

                Return colEntities
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
            ' #End Region ' XDEOperation End Template Expansion{995B7456-17FF-498E-A293-B0175978420B}
        End Function
        Public Function GetAllInfoByActivityANDEventID(ByVal EventID As Integer, ByVal ActivityID As Integer) As MUSTER.Info.LustDocumentCollection
            ' #Region "XDEOperation" ' Begin Template Expansion{995B7456-17FF-498E-A293-B0175978420B}
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Dim colEntities As New MUSTER.Info.LustDocumentCollection
            Try
                If EventID = 0 Or ActivityID = 0 Then
                    Return colEntities
                End If
                strSQL = "spGetTecEventDocuments"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                '
                ' Repeat the following line as many times as necessary
                '
                Params("@EVENT_ACTIVITY_ID").Value = ActivityID
                Params("@EVENT_ID").Value = EventID
                Params("@EVENT_ACTIVITY_DOCUMENT_ID").Value = DBNull.Value

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                While drSet.Read
                    '
                    ' Modify following statement as necessary to generate new info object
                    '
                    Dim oLustDocumentsInfo As New MUSTER.Info.LustDocumentInfo(AltIsDBNull(drSet.Item("EVENT_ACTIVITY_DOCUMENT_ID"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_CLOSED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_SENT_TO_FINANCE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("RECEIVED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV1_RECEIVED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV1_EXTENSION_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV2_RECEIVED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("REV2_EXTENSION_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("START_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EXTENSION_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DUE_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("ISSUE_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EVENT_ID"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_ACTIVITY_ID"), 0), _
                            AltIsDBNull(drSet.Item("DOCUMENT_PROPERTY_ID"), 0), _
                            AltIsDBNull(drSet.Item("DOCUMENT_ID"), 0), _
                            AltIsDBNull(drSet.Item("DOCUMENT_PROPERTY_ID"), 0), _
                            AltIsDBNull(drSet.Item("Commitment_ID"), 0), _
                            AltIsDBNull(drSet.Item("Paid"), False), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DELETED"), 0) _
                    )
                    colEntities.Add(oLustDocumentsInfo)
                End While

                Return colEntities
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
            ' #End Region ' XDEOperation End Template Expansion{995B7456-17FF-498E-A293-B0175978420B}
        End Function
        Public Sub Put(ByRef oLustDocument As MUSTER.Info.LustDocumentInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            ' #Region "XDEOperation" ' Begin Template Expansion{CE44BE18-660C-4873-BF52-C01555F731C8}
            Dim tstDate As Date
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.LustDocument, Integer))) Then
                    returnVal = "You do not have rights to save a LustEvent Document."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutTecEventDocument")


                If oLustDocument.ID = 0 Then
                    Params(0).Value = System.DBNull.Value
                Else
                    Params(0).Value = oLustDocument.ID
                End If

                If oLustDocument.EventId = 0 Then
                    Params(1).Value = System.DBNull.Value
                Else
                    Params(1).Value = oLustDocument.EventId
                End If

                If oLustDocument.AssocActivity = 0 Then
                    Params(2).Value = System.DBNull.Value
                Else
                    Params(2).Value = oLustDocument.AssocActivity
                End If

                If oLustDocument.DocumentType = 0 Then
                    Params(3).Value = System.DBNull.Value
                Else
                    Params(3).Value = oLustDocument.DocumentType
                    'Params(3).Value = oLustDocument.DocClass
                End If
                If oLustDocument.DocumentID = 0 Then
                    Params(4).Value = System.DBNull.Value
                Else
                    Params(4).Value = oLustDocument.DocumentID
                End If

                If oLustDocument.STARTDATE = tstDate Then
                    Params(5).Value = System.DBNull.Value '@START_DATE
                Else
                    Params(5).Value = oLustDocument.STARTDATE
                End If

                If oLustDocument.IssueDate = tstDate Then
                    Params(6).Value = System.DBNull.Value
                Else
                    Params(6).Value = oLustDocument.IssueDate
                End If

                If oLustDocument.DueDate = tstDate Then
                    Params(7).Value = System.DBNull.Value
                Else
                    Params(7).Value = oLustDocument.DueDate
                End If

                If oLustDocument.DocRcvDate = tstDate Then
                    Params(8).Value = System.DBNull.Value
                Else
                    Params(8).Value = oLustDocument.DocRcvDate
                End If

                If oLustDocument.EXTENSIONDATE = tstDate Then
                    Params(9).Value = System.DBNull.Value
                Else
                    Params(9).Value = oLustDocument.EXTENSIONDATE
                End If

                If oLustDocument.REV1RECEIVEDDATE = tstDate Then
                    Params(10).Value = System.DBNull.Value
                Else
                    Params(10).Value = oLustDocument.REV1RECEIVEDDATE
                End If

                If oLustDocument.REV1EXTENSIONDATE = tstDate Then
                    Params(11).Value = System.DBNull.Value
                Else
                    Params(11).Value = oLustDocument.REV1EXTENSIONDATE
                End If
                If oLustDocument.REV2RECEIVEDDATE = tstDate Then
                    Params(12).Value = System.DBNull.Value
                Else
                    Params(12).Value = oLustDocument.REV2RECEIVEDDATE
                End If

                If oLustDocument.REV2EXTENSIONDATE = tstDate Then
                    Params(13).Value = System.DBNull.Value
                Else
                    Params(13).Value = oLustDocument.REV2EXTENSIONDATE
                End If



                If oLustDocument.DocFinancialDate = tstDate Then
                    Params(14).Value = System.DBNull.Value
                Else
                    Params(14).Value = oLustDocument.DocFinancialDate
                End If

                If oLustDocument.DocClosedDate = tstDate Then
                    Params(15).Value = System.DBNull.Value
                Else
                    Params(15).Value = oLustDocument.DocClosedDate
                End If

                Params(16).Value = oLustDocument.Deleted
                Params(17).Value = oLustDocument.CommitmentId
                Params(18).Value = oLustDocument.Paid

                If oLustDocument.ID <= 0 Then
                    Params(19).Value = oLustDocument.CreatedBy
                Else
                    Params(19).Value = oLustDocument.ModifiedBy
                End If

                Params(20).Value = moduleID

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutTecEventDocument", Params)
                '
                ' Perform check for New ID and assign, if necessary
                If Params(0).Value <> oLustDocument.ID Then
                    oLustDocument.ID = Params(0).Value
                End If

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            ' #End Region ' XDEOperation End Template Expansion{CE44BE18-660C-4873-BF52-C01555F731C8}
        End Sub
    End Class
End Namespace



