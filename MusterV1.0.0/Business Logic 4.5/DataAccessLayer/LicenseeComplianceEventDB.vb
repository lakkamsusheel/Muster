'-------------------------------------------------------------------------------
' MUSTER.DataAccess.LicenseeCourseDB
'   Provides the means for marshalling LicenseeComplianceEvent to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MR      6/29/05    Original class definition.
'
' Function                  Description
' 
''-------------------------------------------------------------------------------

Imports System.Data.SqlClient
Imports Utils.DBUtils
Imports System.Data.SqlTypes

Namespace MUSTER.DataAccess
    <Serializable()> _
        Public Class LicenseeComplianceEventDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
        Public Function GetLCEEnforcementHistory(ByVal nLicenseeID As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.LicenseeComplianceEventCollection
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetCAELCEEnforcementHistory"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@LICENSEE_ID").Value = nLicenseeID
                Params("@Deleted").Value = False

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colLCE As New MUSTER.Info.LicenseeComplianceEventCollection
                While drSet.Read
                    Dim oLCEInfo As New MUSTER.Info.LicenseeComplianceEventInfo(AltIsDBNull(drSet.Item("LCE_ID"), 0), _
                                            AltIsDBNull(drSet.Item("Licensee_ID"), 0), _
                                            AltIsDBNull(drSet.Item("Owner_ID"), 0), _
                                            AltIsDBNull(drSet.Item("Owner"), String.Empty), _
                                            AltIsDBNull(drSet.Item("Facility"), String.Empty), _
                                            AltIsDBNull(drSet.Item("LICENSEE"), String.Empty), _
                                            AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                                            AltIsDBNull(drSet.Item("Licensee_CITATION_id"), 0), _
                                            AltIsDBNull(drSet.Item("citation_DUE_date"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("CITATIOn_RECEIVED_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("RESCINDED"), False), _
                                            AltIsDBNull(drSet.Item("LCE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LCE_PROCESS_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("NEXT_DUE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("OVERRIDE_DUE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LCE_STATUS"), 0), _
                                            AltIsDBNull(drSet.Item("STATUS"), String.Empty), _
                                            AltIsDBNull(drSet.Item("ESCALATION"), 0), _
                                            AltIsDBNull(drSet.Item("ESCALATION_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("POLICY_AMOUNT"), -1.0), _
                                            AltIsDBNull(drSet.Item("OVERRIDE_AMOUNT"), -1.0), _
                                            AltIsDBNull(drSet.Item("SETTLEMENT_AMOUNT"), -1.0), _
                                            AltIsDBNull(drSet.Item("PAID_AMOUNT"), -1.0), _
                                            AltIsDBNull(drSet.Item("DATE_RECEIVED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("WORKSHOP_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("WORKSHOP_RESULT"), 0), _
                                             AltIsDBNull(drSet.Item("WORKSHOP_RESULT_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("SHOW_CAUSE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("SHOW_CAUSE_RESULTS"), 0), _
                                            AltIsDBNull(drSet.Item("SHOW_CAUSE_RESULT_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("COMMISSION_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("COMMISSION_RESULTS"), 0), _
                                            AltIsDBNull(drSet.Item("COMMISSION_RESULT_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("PENDING_LETTER"), 0), _
                                            AltIsDBNull(drSet.Item("PENDING_LETTER_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("LETTER_GENERATED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LETTER_PRINTED"), False), _
                                            AltIsDBNull(drSet.Item("CITATION_TEXT"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DELETED"), False), _
                                            AltIsDBNull(drSet.Item("PENDING_LETTER_TEMPLATE_NUM"), 0))
                    colLCE.Add(oLCEInfo)
                End While
                Return colLCE
                If Not drSet.IsClosed Then drSet.Close()
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        Public Function GetAllInfo(Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.LicenseeComplianceEventCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetCAELCE1"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@LCE_ID").Value = DBNull.Value
                Params("@DELETED").Value = showDeleted

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colLCE As New MUSTER.Info.LicenseeComplianceEventCollection
                While drSet.Read
                    Dim oLCEInfo As New MUSTER.Info.LicenseeComplianceEventInfo(AltIsDBNull(drSet.Item("LCE_ID"), 0), _
                                            AltIsDBNull(drSet.Item("Licensee_ID"), 0), _
                                            AltIsDBNull(drSet.Item("Owner_ID"), 0), _
                                            AltIsDBNull(drSet.Item("Owner"), String.Empty), _
                                            AltIsDBNull(drSet.Item("Facility"), String.Empty), _
                                            AltIsDBNull(drSet.Item("LICENSEE"), String.Empty), _
                                            AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                                            AltIsDBNull(drSet.Item("Licensee_CITATION_id"), 0), _
                                            AltIsDBNull(drSet.Item("citation_DUE_date"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("CITATIOn_RECEIVED_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("RESCINDED"), False), _
                                            AltIsDBNull(drSet.Item("LCE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LCE_PROCESS_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("NEXT_DUE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("OVERRIDE_DUE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LCE_STATUS"), 0), _
                                            AltIsDBNull(drSet.Item("STATUS"), String.Empty), _
                                            AltIsDBNull(drSet.Item("ESCALATION"), 0), _
                                            AltIsDBNull(drSet.Item("ESCALATION_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("POLICY_AMOUNT"), -1.0), _
                                            AltIsDBNull(drSet.Item("OVERRIDE_AMOUNT"), -1.0), _
                                            AltIsDBNull(drSet.Item("SETTLEMENT_AMOUNT"), -1.0), _
                                            AltIsDBNull(drSet.Item("PAID_AMOUNT"), -1.0), _
                                            AltIsDBNull(drSet.Item("DATE_RECEIVED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("WORKSHOP_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("WORKSHOP_RESULT"), 0), _
                                             AltIsDBNull(drSet.Item("WORKSHOP_RESULT_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("SHOW_CAUSE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("SHOW_CAUSE_RESULTS"), 0), _
                                            AltIsDBNull(drSet.Item("SHOW_CAUSE_RESULT_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("COMMISSION_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("COMMISSION_RESULTS"), 0), _
                                            AltIsDBNull(drSet.Item("COMMISSION_RESULT_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("PENDING_LETTER"), 0), _
                                            AltIsDBNull(drSet.Item("PENDING_LETTER_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("LETTER_GENERATED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LETTER_PRINTED"), False), _
                                            AltIsDBNull(drSet.Item("CITATION_TEXT"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DELETED"), False), _
                                            AltIsDBNull(drSet.Item("PENDING_LETTER_TEMPLATE_NUM"), 0))
                    colLCE.Add(oLCEInfo)
                End While
                Return colLCE
                If Not drSet.IsClosed Then drSet.Close()
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        Public Function DBGetByID(ByVal LCEID As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.LicenseeComplianceEventInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                If LCEID = 0 Then
                    Return New MUSTER.Info.LicenseeComplianceEventInfo
                End If
                strSQL = "spGetCAELCE1"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@LCE_ID").Value = LCEID
                Params("@Deleted").Value = False

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()

                    Return New MUSTER.Info.LicenseeComplianceEventInfo(AltIsDBNull(drSet.Item("LCE_ID"), 0), _
                                            AltIsDBNull(drSet.Item("Licensee_ID"), 0), _
                                            AltIsDBNull(drSet.Item("Owner_ID"), 0), _
                                            AltIsDBNull(drSet.Item("Owner"), String.Empty), _
                                            AltIsDBNull(drSet.Item("Facility"), String.Empty), _
                                            AltIsDBNull(drSet.Item("LICENSEE"), String.Empty), _
                                            AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                                            AltIsDBNull(drSet.Item("Licensee_CITATION_id"), 0), _
                                            AltIsDBNull(drSet.Item("citation_DUE_date"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("CITATIOn_RECEIVED_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("RESCINDED"), False), _
                                            AltIsDBNull(drSet.Item("LCE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LCE_PROCESS_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("NEXT_DUE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("OVERRIDE_DUE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LCE_STATUS"), 0), _
                                            AltIsDBNull(drSet.Item("STATUS"), String.Empty), _
                                            AltIsDBNull(drSet.Item("ESCALATION"), 0), _
                                            AltIsDBNull(drSet.Item("ESCALATION_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("POLICY_AMOUNT"), -1.0), _
                                            AltIsDBNull(drSet.Item("OVERRIDE_AMOUNT"), -1.0), _
                                            AltIsDBNull(drSet.Item("SETTLEMENT_AMOUNT"), -1.0), _
                                            AltIsDBNull(drSet.Item("PAID_AMOUNT"), -1.0), _
                                            AltIsDBNull(drSet.Item("DATE_RECEIVED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("WORKSHOP_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("WORKSHOP_RESULT"), 0), _
                                             AltIsDBNull(drSet.Item("WORKSHOP_RESULT_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("SHOW_CAUSE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("SHOW_CAUSE_RESULTS"), 0), _
                                            AltIsDBNull(drSet.Item("SHOW_CAUSE_RESULT_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("COMMISSION_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("COMMISSION_RESULTS"), 0), _
                                            AltIsDBNull(drSet.Item("COMMISSION_RESULT_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("PENDING_LETTER"), 0), _
                                            AltIsDBNull(drSet.Item("PENDING_LETTER_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("LETTER_GENERATED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LETTER_PRINTED"), False), _
                                            AltIsDBNull(drSet.Item("CITATION_TEXT"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DELETED"), False), _
                                            AltIsDBNull(drSet.Item("PENDING_LETTER_TEMPLATE_NUM"), 0))
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not LCEID = 0 Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBGetByLicenseeID(Optional ByVal LicenseeID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.LicenseeCourseCollection
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim strVal As String = ""
            Dim Params As Collection
            Dim colLicCourse As New MUSTER.Info.LicenseeCourseCollection
            Try
                strSQL = "spGetCOMLicenseeCourses"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@LIC_COUR_ID").Value = DBNull.Value
                Params("@LICENSEE_ID").Value = IIf(LicenseeID = 0, DBNull.Value, LicenseeID.ToString)
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = False

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    While drSet.Read
                        Dim LicCourseInfo As New MUSTER.Info.LicenseeCourseInfo(drSet.Item("LIC_COUR_ID"), _
                                            AltIsDBNull(drSet.Item("LICENSEE_ID"), 0), _
                                            AltIsDBNull(drSet.Item("PROVIDER_ID"), 0), _
                                            AltIsDBNull(drSet.Item("COURSE_TYPE_ID"), 0), _
                                            AltIsDBNull(drSet.Item("COURSE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("DELETED"), False), _
                                            AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                        colLicCourse.Add(LicCourseInfo)
                    End While
                End If
                Return colLicCourse
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Sub Put(ByRef oLCEInfo As MUSTER.Info.LicenseeComplianceEventInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.CAELicenseeCompliantEvent, Integer))) Then
                    returnVal = "You do not have rights to save Licensee Compliance Event."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Dim dtTempDate As Date
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutCAELCEAssignment")

                If oLCEInfo.ID <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = oLCEInfo.ID
                End If

                Params(1).Value = oLCEInfo.LicenseeID
                Params(2).Value = oLCEInfo.FacilityID
                Params(3).Value = oLCEInfo.LicenseeCitationID
                If Date.Compare(oLCEInfo.CitationDueDate, dtTempDate) = 0 Then
                    Params(4).Value = SqlDateTime.Null
                Else
                    Params(4).Value = oLCEInfo.CitationDueDate
                End If
                If Date.Compare(oLCEInfo.CitationReceivedDate, dtTempDate) = 0 Then
                    Params(5).Value = SqlDateTime.Null
                Else
                    Params(5).Value = oLCEInfo.CitationReceivedDate
                End If
                Params(6).Value = oLCEInfo.Rescinded
                If Date.Compare(oLCEInfo.LCEDate, dtTempDate) = 0 Then
                    Params(7).Value = SqlDateTime.Null
                Else
                    Params(7).Value = oLCEInfo.LCEDate
                End If
                If Date.Compare(oLCEInfo.LCEProcessDate, dtTempDate) = 0 Then
                    Params(8).Value = SqlDateTime.Null
                Else
                    Params(8).Value = oLCEInfo.LCEProcessDate
                End If
                If Date.Compare(oLCEInfo.NextDueDate, dtTempDate) = 0 Then
                    Params(9).Value = SqlDateTime.Null
                Else
                    Params(9).Value = oLCEInfo.NextDueDate
                End If
                If Date.Compare(oLCEInfo.OverrideDueDate, dtTempDate) = 0 Then
                    Params(10).Value = SqlDateTime.Null
                Else
                    Params(10).Value = oLCEInfo.OverrideDueDate
                End If

                Params(11).Value = oLCEInfo.LCEStatus
                Params(12).Value = oLCEInfo.Escalation
                Params(13).Value = oLCEInfo.PolicyAmount
                Params(14).Value = oLCEInfo.OverrideAmount
                Params(15).Value = oLCEInfo.SettlementAmount
                Params(16).Value = oLCEInfo.PaidAmount

                If Date.Compare(oLCEInfo.DateReceived, dtTempDate) = 0 Then
                    Params(17).Value = SqlDateTime.Null
                Else
                    Params(17).Value = oLCEInfo.DateReceived
                End If

                If Date.Compare(oLCEInfo.WorkShopDate, dtTempDate) = 0 Then
                    Params(18).Value = SqlDateTime.Null
                Else
                    Params(18).Value = oLCEInfo.WorkShopDate
                End If

                Params(19).Value = oLCEInfo.WorkshopResult

                If Date.Compare(oLCEInfo.ShowCauseDate, dtTempDate) = 0 Then
                    Params(20).Value = SqlDateTime.Null
                Else
                    Params(20).Value = oLCEInfo.ShowCauseDate
                End If

                Params(21).Value = oLCEInfo.ShowCauseResults

                If Date.Compare(oLCEInfo.CommissionDate, dtTempDate) = 0 Then
                    Params(22).Value = SqlDateTime.Null
                Else
                    Params(22).Value = oLCEInfo.CommissionDate
                End If

                Params(23).Value = oLCEInfo.CommissionResults
                Params(24).Value = oLCEInfo.PendingLetter

                If Date.Compare(oLCEInfo.LetterGenerated, dtTempDate) = 0 Then
                    Params(25).Value = SqlDateTime.Null
                Else
                    Params(25).Value = oLCEInfo.LetterGenerated
                End If

                Params(26).Value = oLCEInfo.LetterPrinted
                Params(27).Value = DBNull.Value
                Params(28).Value = DBNull.Value
                Params(29).Value = DBNull.Value
                Params(30).Value = DBNull.Value
                Params(31).Value = oLCEInfo.Deleted

                If oLCEInfo.ID <= 0 Then
                    Params(32).Value = oLCEInfo.CreatedBy
                Else
                    Params(32).Value = oLCEInfo.ModifiedBy
                End If
                Params(33).Value = oLCEInfo.PendingLetterTemplateNum

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutCAELCEAssignment", Params)
                If oLCEInfo.ID <= 0 Then
                    oLCEInfo.ID = Params(0).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function GetLCE(Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim dsData As DataSet
            Dim strSQl As String
            Try
                strSQl = "spGetCAELCE"
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQl)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function getFacilityAssignedInspector(ByVal FacilityID As Integer, Optional ByVal showDeleted As Boolean = False) As Integer
            Dim InspectorID As Integer
            Dim strSQl As String
            Dim Params() As SqlParameter
            Try
                strSQl = "spGetCAEFacilityAssignedInpector"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQl)
                Params(0).Value = FacilityID
                ' tried executescalar, but not working
                Dim ds As DataSet = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQl, Params)
                If ds.Tables(0).Rows.Count = 0 Then
                    Return 0
                Else
                    Return CInt(ds.Tables(0).Rows(0).Item(0))
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetDS(ByVal strSQL As String) As DataSet
            Dim dsData As DataSet
            Try
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetCitationList(Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim dsData As DataSet
            Dim strSQl As String
            Try
                strSQl = "SELECT CITATION_ID,CITATION_TEXT,POLICY_PENALTY FROM dbo.tblCAE_LICENSEE_CITATION_PENALTY"
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQl)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetCitation(ByVal LceID As Integer, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim dsData As DataSet
            Dim strSQl As String
            Dim Params() As SqlParameter
            Try
                strSQl = "spGetCAELicenseeCitation"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQl)
                Params(0).Value = LceID
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQl, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetDropDownValues(ByVal propertyTypeID As Integer) As DataSet
            Dim dsData As DataSet
            Dim strSQl As String
            Dim Params() As SqlParameter
            Try
                strSQl = "spGetCAELCEGridDropDownValues"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQl)
                Params(0).Value = propertyTypeID
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQl, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
    End Class
End Namespace
