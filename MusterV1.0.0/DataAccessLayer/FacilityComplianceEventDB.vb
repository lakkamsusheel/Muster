'-------------------------------------------------------------------------------
' MUSTER.DataAccess.FacilityComplianceEventDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0                              Original class definition
'
'
' Function                  Description
' GetAllInfo()        Returns an EntityCollection containing all Entity objects in the repository.
' DBGetByName(NAME)   Returns an EntityInfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an EntityInfo object indicated by arg ID.
'
' NOTE: This file to be used as FacilityComplianceEvent to build other objects.
'       Replace keyword "FacilityComplianceEvent" with respective Object name.
'       In addition, change SQL strings and data stream Items as necessary
'       Don't forget to update the history information above!!!
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils
Imports System.Data.SqlTypes

Namespace MUSTER.DataAccess
    Public Class FacilityComplianceEventDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
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
#End Region

        Public Function DBGetByID(Optional ByVal id As Integer = 0, Optional ByVal inspID As Int64 = 0, Optional ByVal ownerID As Int64 = 0, Optional ByVal facID As Int64 = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FacilityComplianceEventsCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Dim colFCE As New MUSTER.Info.FacilityComplianceEventsCollection

            If id = 0 And inspID = 0 And ownerID = 0 And facID = 0 Then
                Return colFCE
            End If

            Try
                strSQL = "spGetCAEFCE"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@FCE_ID").Value = IIf(id = 0, DBNull.Value, id)
                Params("@INSPECTION_ID").Value = IIf(inspID = 0, DBNull.Value, inspID)
                Params("@FACILITY_ID").Value = IIf(facID = 0, DBNull.Value, facID)
                Params("@OWNER_ID").Value = IIf(ownerID = 0, DBNull.Value, ownerID)
                Params("@FACILITY_ID").Value = IIf(facID = 0, DBNull.Value, facID)
                Params("@DELETED").Value = showDeleted

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                While drSet.Read
                    Dim oFCEInfo As New MUSTER.Info.FacilityComplianceEventInfo(drSet.Item("FCE_ID"), _
                        AltIsDBNull(drSet.Item("INSPECTION_ID"), 0), _
                        AltIsDBNull(drSet.Item("OWNER_ID"), 0), _
                        AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                        AltIsDBNull(drSet.Item("FCE_DATE"), CDate("01/01/0001")), _
                        AltIsDBNull(drSet.Item("SOURCE"), String.Empty), _
                        AltIsDBNull(drSet.Item("DUE_DATE"), CDate("01/01/0001")), _
                        AltIsDBNull(drSet.Item("RECEIVED_DATE"), CDate("01/01/0001")), _
                        AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                        AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                        AltIsDBNull(drSet.Item("DELETED"), False), _
                        AltIsDBNull(drSet.Item("OCE_GENERATED"), False), _
                        AltIsDBNull(drSet.Item("OCE_ID"), 0))
                    colFCE.Add(oFCEInfo)
                End While
                Return colFCE
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Sub Put(ByRef oFCEInfo As MUSTER.Info.FacilityComplianceEventInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal OverrideRights As Boolean = False)
            Try
                If Not OverrideRights Then
                    If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.CAEFacilityCompliantEvent, Integer))) Then
                        returnVal = "You do not have rights to save Facility Compliance Event."
                        Exit Sub
                    Else
                        returnVal = String.Empty
                    End If
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Dim strSQL As String = "spPutCAEFCE"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)

                If oFCEInfo.ID <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = oFCEInfo.ID
                End If

                Params(1).Value = oFCEInfo.InspectionID
                Params(2).Value = oFCEInfo.OwnerID
                Params(3).Value = oFCEInfo.FacilityID
                Params(4).Value = oFCEInfo.FCEDate
                Params(5).Value = IIf(oFCEInfo.Source = String.Empty, DBNull.Value, oFCEInfo.Source.Trim)

                If Date.Compare(oFCEInfo.DueDate, CDate("01/01/0001")) = 0 Then
                    Params(6).Value = SqlDateTime.Null
                Else
                    Params(6).Value = oFCEInfo.DueDate
                End If
                If Date.Compare(oFCEInfo.ReceivedDate, CDate("01/01/0001")) = 0 Then
                    Params(7).Value = SqlDateTime.Null
                Else
                    Params(7).Value = oFCEInfo.ReceivedDate
                End If
                Params(8).Value = DBNull.Value
                Params(9).Value = DBNull.Value
                Params(10).Value = DBNull.Value
                Params(11).Value = DBNull.Value
                Params(12).Value = oFCEInfo.Deleted
                Params(13).Value = oFCEInfo.OCEGenerated
                Params(14).Value = oFCEInfo.OCEID

                If oFCEInfo.ID <= 0 Then
                    Params(15).Value = oFCEInfo.CreatedBy
                Else
                    Params(15).Value = oFCEInfo.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If oFCEInfo.ID <= 0 Then
                    oFCEInfo.ID = Params(0).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Public Function GetAllInfo(Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FacilityComplianceEventsCollection
        '    Dim drSet As SqlDataReader
        '    Try
        '        drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM vCAEGetFCE")
        '        Dim colFCE As New MUSTER.Info.FacilityComplianceEventsCollection
        '        While drSet.Read
        '            Dim oFCEInfo As New MUSTER.Info.FacilityComplianceEventInfo(drSet.Item("FCE_ID"), _
        '                AltIsDBNull(drSet.Item("INSPECTION_ID"), 0), _
        '                AltIsDBNull(drSet.Item("OWNER_ID"), 0), _
        '                AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
        '                AltIsDBNull(drSet.Item("FCE_DATE"), CDate("01/01/0001")), _
        '                AltIsDBNull(drSet.Item("SOURCE"), String.Empty), _
        '                AltIsDBNull(drSet.Item("DUE_DATE"), CDate("01/01/0001")), _
        '                AltIsDBNull(drSet.Item("RECEIVED_DATE"), CDate("01/01/0001")), _
        '                AltIsDBNull(drSet.Item("O_NAME"), String.Empty), _
        '                AltIsDBNull(drSet.Item("NAME"), String.Empty), _
        '                AltIsDBNull(drSet.Item("USER_NAME"), String.Empty), _
        '                AltIsDBNull(drSet.Item("INSPECTED ON"), CDate("01/01/0001")), _
        '                AltIsDBNull(drSet.Item("CITATIONS"), 0), _
        '                AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
        '                AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
        '                AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
        '                AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
        '                AltIsDBNull(drSet.Item("DELETED"), False), _
        '                AltIsDBNull(drSet.Item("OCE_GENERATED"), False))
        '            colFCE.Add(oFCEInfo)
        '        End While
        '        Return colFCE
        '        If Not drSet.IsClosed Then drSet.Close()
        '    Catch Ex As Exception
        '        MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try

        'End Function
        'Public Function DBGetByID(ByVal LCEID As Int32) As MUSTER.Info.FacilityComplianceEventInfo
        '    Dim drSet As SqlDataReader
        '    Dim strVal As String
        '    Dim strSQL As String
        '    Dim Params As Collection

        '    Try
        '        strSQL = "spGetCAEFCE"
        '        strVal = LCEID.ToString

        '        Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
        '        Params("@FCE_ID").Value = strVal
        '        Params("@Deleted").Value = False

        '        drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
        '        If drSet.HasRows Then
        '            drSet.Read()

        '            Return New MUSTER.Info.FacilityComplianceEventInfo(drSet.Item("FCE_ID"), _
        '                AltIsDBNull(drSet.Item("INSPECTION_ID"), 0), _
        '                AltIsDBNull(drSet.Item("OWNER_ID"), 0), _
        '                AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
        '                AltIsDBNull(drSet.Item("FCE_DATE"), CDate("01/01/0001")), _
        '                AltIsDBNull(drSet.Item("SOURCE"), String.Empty), _
        '                AltIsDBNull(drSet.Item("DUE_DATE"), CDate("01/01/0001")), _
        '                AltIsDBNull(drSet.Item("RECEIVED_DATE"), CDate("01/01/0001")), _
        '                AltIsDBNull(drSet.Item("O_NAME"), String.Empty), _
        '                AltIsDBNull(drSet.Item("NAME"), String.Empty), _
        '                AltIsDBNull(drSet.Item("USER_NAME"), String.Empty), _
        '                AltIsDBNull(drSet.Item("INSPECTED ON"), CDate("01/01/0001")), _
        '                AltIsDBNull(drSet.Item("CITATIONS"), 0), _
        '                AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
        '                AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
        '                AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
        '                AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
        '                AltIsDBNull(drSet.Item("DELETED"), False), _
        '                AltIsDBNull(drSet.Item("OCE_GENERATED"), False))
        '        Else
        '            Return New MUSTER.Info.FacilityComplianceEventInfo
        '        End If
        '    Catch ex As Exception
        '        MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    Finally
        '        If Not drSet.IsClosed Then drSet.Close()
        '    End Try

        'End Function
        'Public Sub Put(ByRef oFCEInfo As MUSTER.Info.FacilityComplianceEventInfo)
        '    Try
        '        Dim Params() As SqlParameter
        '        Dim dtTempDate As Date
        '        Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutCAEFCEAssignmnet")

        '        If oFCEInfo.ID <= 0 Then
        '            Params(0).Value = 0
        '        Else
        '            Params(0).Value = oFCEInfo.ID
        '        End If

        '        Params(1).Value = oFCEInfo.InspectionID
        '        Params(2).Value = oFCEInfo.OwnerID
        '        Params(3).Value = oFCEInfo.FacilityID

        '        If Date.Compare(oFCEInfo.FCEDate, dtTempDate) = 0 Then
        '            Params(4).Value = SqlDateTime.Null
        '        Else
        '            Params(4).Value = oFCEInfo.FCEDate
        '        End If

        '        Params(5).Value = oFCEInfo.Source

        '        If Date.Compare(oFCEInfo.DueDate, dtTempDate) = 0 Then
        '            Params(6).Value = SqlDateTime.Null
        '        Else
        '            Params(6).Value = oFCEInfo.DueDate
        '        End If
        '        If Date.Compare(oFCEInfo.ReceivedDate, dtTempDate) = 0 Then
        '            Params(7).Value = SqlDateTime.Null
        '        Else
        '            Params(7).Value = oFCEInfo.ReceivedDate
        '        End If
        '        Params(8).Value = DBNull.Value
        '        Params(9).Value = DBNull.Value
        '        Params(10).Value = DBNull.Value
        '        Params(11).Value = DBNull.Value
        '        Params(12).Value = oFCEInfo.Deleted
        '        Params(13).Value = oFCEInfo.OCEGenerated

        '        SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutCAEFCEAssignmnet", Params)
        '        If oFCEInfo.ID <= 0 Then
        '            oFCEInfo.ID = Params(0).Value
        '        End If
        '    Catch Ex As Exception
        '        MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try

        'End Sub

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
        Public Function DBGetInspections(Optional ByVal showDeleted As Boolean = False, Optional ByVal facility_id As Integer = 0, Optional ByVal managerID As Integer = Nothing) As DataSet
            Dim dsData As DataSet
            Dim strSQl As String
            Dim Params() As SqlParameter
            Try
                strSQl = "spGetCAEInspections"

                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQl)
                Params(0).Value = IIf(facility_id > 0, facility_id, DBNull.Value)
                If Not managerID = Nothing Then
                    Params(1).Value = managerID
                End If

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQl, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetAdminUsers() As DataSet
            Dim dsData As DataSet
            Dim strSQl As String
            Dim Params() As SqlParameter
            Try
                strSQl = "spGetAdminUser"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQl)
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQl, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetCompliances(Optional ByVal facID As Integer = 0, Optional ByVal showDeleted As Boolean = False, Optional ByVal managerID As Integer = Nothing) As DataSet
            Dim drSet As SqlDataReader
            Dim dsData As DataSet
            Dim strSQl As String
            Dim Params() As SqlParameter
            Try
                strSQl = "spGetCAECompliance"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQl)
                Params(0).Value = IIf(facID <= 0, DBNull.Value, facID)
                Params(1).Value = showDeleted

                If Not managerID = Nothing Then
                    Params(2).Value = managerID
                End If

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQl, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetCitations(Optional ByVal citationID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim dsData As DataSet
            Dim strSQl As String
            Dim Params As Collection
            Try
                strSQl = "spGetCAECitationPenalty"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQl)
                Params("@CITATION_ID").Value = IIf(citationID = 0, DBNull.Value, citationID)
                Params("@DELETED").Value = showDeleted

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQl, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        'Public Function GetCitationList(Optional ByVal showDeleted As Boolean = False) As DataSet
        '    Dim dsData As DataSet
        '    Dim strSQl As String
        '    Try
        '        strSQl = "SELECT StateCitation as Citation,Category,Description as CitationText,Citation_ID FROM dbo.tblCAE_CITATION_PENALTY"
        '        dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQl)
        '        Return dsData
        '    Catch ex As Exception
        '        MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function
    End Class
End Namespace
