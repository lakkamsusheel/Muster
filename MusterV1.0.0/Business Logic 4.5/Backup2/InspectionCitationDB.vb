'-------------------------------------------------------------------------------
' MUSTER.DataAccess.InspectionCitationDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MNR        06/15/05    Original class definition
'
'
' Function                  Description
' GetAllInfo()        Returns an EntityCollection containing all Entity objects in the repository.
' DBGetByName(NAME)   Returns an EntityInfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an EntityInfo object indicated by arg ID.
'
' NOTE: This file to be used as InspectionCitation to build other objects.
'       Replace keyword "InspectionCitation" with respective Object name.
'       In addition, change SQL strings and data stream Items as necessary
'       Don't forget to update the history information above!!!
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class InspectionCitationDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function DBGetByID(ByVal id As Int64, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.InspectionCitationInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Try
                If id <= 0 Then
                    Return New MUSTER.Info.InspectionCitationInfo
                End If
                strSQL = "spGetInspectionCitation"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@INS_CIT_ID").Value = id
                Params("@INSPECTION_ID").Value = DBNull.Value
                Params("@FCE_ID").Value = DBNull.Value
                Params("@OCE_ID").Value = DBNull.Value
                Params("@DELETED").Value = showDeleted
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.InspectionCitationInfo(drSet.Item("INS_CIT_ID"), _
                                                                drSet.Item("INSPECTION_ID"), _
                                                                AltIsDBNull(drSet.Item("QUESTION_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("FCE_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("OCE_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("CITATION_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("CCAT"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("RESCINDED"), False), _
                                                                AltIsDBNull(drSet.Item("CITATION_DUE_DATE"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("CITATION_RECEIVED_DATE"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("NFA_DATE"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("DELETED"), False), _
                                                                AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                Else
                    Return New MUSTER.Info.InspectionCitationInfo
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
        Public Function DBGetByOtherID(Optional ByVal inspID As Int64 = 0, Optional ByVal fceID As Int64 = 0, Optional ByVal oceID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.InspectionCitationsCollection
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Dim citationCollection As New MUSTER.Info.InspectionCitationsCollection

            If inspID <= 0 And fceID <= 0 And oceID <= 0 Then
                Return citationCollection
            End If

            Try
                strSQL = "spGetInspectionCitation"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@INS_CIT_ID").Value = DBNull.Value
                Params("@INSPECTION_ID").Value = IIf(inspID <= 0, DBNull.Value, inspID)
                Params("@FCE_ID").Value = IIf(fceID <= 0, DBNull.Value, fceID)
                Params("@OCE_ID").Value = IIf(oceID <= 0, DBNull.Value, oceID)
                Params("@DELETED").Value = showDeleted
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                While drSet.Read
                    Dim citationInfo As New MUSTER.Info.InspectionCitationInfo(drSet.Item("INS_CIT_ID"), _
                                                                drSet.Item("INSPECTION_ID"), _
                                                                AltIsDBNull(drSet.Item("QUESTION_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("FCE_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("OCE_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("CITATION_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("CCAT"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("RESCINDED"), False), _
                                                                AltIsDBNull(drSet.Item("CITATION_DUE_DATE"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("CITATION_RECEIVED_DATE"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("NFA_DATE"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("DELETED"), False), _
                                                                AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                    citationCollection.Add(citationInfo)
                End While
                Return citationCollection
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBGetDS(ByVal strSQL As String) As DataSet
            Dim dsData As DataSet
            Try
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        Public Sub Put(ByRef oInspectionCitationInfo As MUSTER.Info.InspectionCitationInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal OverrideRights As Boolean = False)
            Try
                If Not OverrideRights Then
                    If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Citation, Integer))) Then
                        returnVal = "You do not have rights to save Inspection Citation."
                        Exit Sub
                    Else
                        returnVal = String.Empty
                    End If
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Dim strSQL As String = "spPutInspectionCitation"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(oInspectionCitationInfo.ID <= 0, 0, oInspectionCitationInfo.ID)
                Params(1).Value = oInspectionCitationInfo.InspectionID
                Params(2).Value = IIf(oInspectionCitationInfo.FacilityID <= 0, DBNull.Value, oInspectionCitationInfo.FacilityID)
                Params(3).Value = IIf(oInspectionCitationInfo.FCEID <= 0, DBNull.Value, oInspectionCitationInfo.FCEID)
                Params(4).Value = IIf(oInspectionCitationInfo.QuestionID = 0, DBNull.Value, oInspectionCitationInfo.QuestionID)
                Params(5).Value = IIf(oInspectionCitationInfo.CitationID <= 0, DBNull.Value, oInspectionCitationInfo.CitationID)
                Params(6).Value = IIf(oInspectionCitationInfo.CCAT = String.Empty, DBNull.Value, oInspectionCitationInfo.CCAT)
                Params(7).Value = oInspectionCitationInfo.Rescinded
                If Date.Compare(oInspectionCitationInfo.CitationDueDate, CDate("01/01/0001")) = 0 Then
                    Params(8).Value = DBNull.Value
                Else
                    Params(8).Value = oInspectionCitationInfo.CitationDueDate
                End If
                If Date.Compare(oInspectionCitationInfo.CitationReceivedDate, CDate("01/01/0001")) = 0 Then
                    Params(9).Value = DBNull.Value
                Else
                    Params(9).Value = oInspectionCitationInfo.CitationReceivedDate
                End If
                If Date.Compare(oInspectionCitationInfo.NFADate, CDate("01/01/0001")) = 0 Then
                    Params(10).Value = DBNull.Value
                Else
                    Params(10).Value = oInspectionCitationInfo.NFADate
                End If
                Params(11).Value = oInspectionCitationInfo.Deleted
                Params(12).Value = DBNull.Value
                Params(13).Value = DBNull.Value
                Params(14).Value = DBNull.Value
                Params(15).Value = DBNull.Value
                Params(16).Value = IIf(oInspectionCitationInfo.OCEID <= 0, DBNull.Value, oInspectionCitationInfo.OCEID)

                If oInspectionCitationInfo.ID <= 0 Then
                    Params(17).Value = oInspectionCitationInfo.CreatedBy
                Else
                    Params(17).Value = oInspectionCitationInfo.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Params(0).Value <> oInspectionCitationInfo.ID Then
                    oInspectionCitationInfo.ID = Params(0).Value
                End If
                oInspectionCitationInfo.CreatedBy = AltIsDBNull(Params(12).Value, String.Empty)
                oInspectionCitationInfo.CreatedOn = AltIsDBNull(Params(13).Value, CDate("01/01/0001"))
                oInspectionCitationInfo.ModifiedBy = AltIsDBNull(Params(14).Value, String.Empty)
                oInspectionCitationInfo.ModifiedOn = AltIsDBNull(Params(15).Value, CDate("01/01/0001"))
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        Public Function DBCheckCitationExists(ByVal onDate As Date, Optional ByVal facID As Integer = 0, Optional ByVal showDeleted As Boolean = False, Optional ByVal citationID As Integer = 0, Optional ByVal fceCreated As Int16 = -1, Optional ByVal oceCreated As Int16 = -1, Optional ByVal rescinded As Int16 = -1, Optional ByVal strExcludeOCE As String = "") As Boolean
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            ' when onDate has valid date (not "01/01/0001", it will check only the citations which have/had fce
            Try
                strSQL = "spCheckCitationExists"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@FACILITY_ID").Value = IIf(facID = 0, DBNull.Value, facID)
                Params("@DELETED").Value = showDeleted
                Params("@CITATION_ID").Value = IIf(citationID = 0, DBNull.Value, citationID)
                Params("@FCE_CREATED").Value = fceCreated
                Params("@OCE_CREATED").Value = oceCreated
                Params("@RESCINDED").Value = rescinded
                If Date.Compare(onDate, CDate("01/01/0001")) = 0 Then
                    Params("@ONDATE").Value = DBNull.Value
                Else
                    Params("@ONDATE").Value = onDate.Date
                End If
                Params("@EXCLUDEOCES").Value = strExcludeOCE

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    drSet.Read()
                    Return drSet.Item("EXISTS")
                Else
                    Return False
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
    End Class
End Namespace
