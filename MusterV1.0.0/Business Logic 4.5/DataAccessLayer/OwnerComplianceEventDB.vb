'-------------------------------------------------------------------------------
' MUSTER.DataAccess.OwnerComplianceEventDB
'   Provides the means for marshalling OwnerComplianceEventDB to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MR           6/29/05    Original class definition.
'  2.0    Thomas Franey   5/21/09    Added Comments
'
' Function                  Description
' 
''-------------------------------------------------------------------------------

Imports System.Data.SqlClient
Imports Utils.DBUtils
Imports System.Data.SqlTypes

Namespace MUSTER.DataAccess
    <Serializable()> _
        Public Class OwnerComplianceEventDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
        Public Function DBGetByID(Optional ByVal id As Integer = 0, Optional ByVal OwnerID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.OwnerComplianceEventsCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader

            If id = 0 And OwnerID = 0 Then
                Return New MUSTER.Info.OwnerComplianceEventsCollection
            End If

            Try
                strSQL = "spGetCAEOCE"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@OCE_ID").Value = IIf(id = 0, DBNull.Value, id)
                Params("@OWNER_ID").Value = IIf(OwnerID = 0, DBNull.Value, OwnerID)
                Params("@DELETED").Value = showDeleted

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colOCE As New MUSTER.Info.OwnerComplianceEventsCollection
                While drSet.Read
                    Dim oOCEInfo As New MUSTER.Info.OwnerComplianceEventInfo(drSet.Item("OCE_ID"), _
                                            drSet.Item("OWNER_ID"), _
                                            AltIsDBNull(drSet.Item("CITATION"), 0), _
                                            AltIsDBNull(drSet.Item("CITATION_DUE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("RESCINDED"), False), _
                                            AltIsDBNull(drSet.Item("OCE_PATH"), 0), _
                                            AltIsDBNull(drSet.Item("OCE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("OCE_PROCESS_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("NEXT_DUE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("OVERRIDE_DUE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("OCE_STATUS"), 0), _
                                            AltIsDBNull(drSet.Item("ESCALATION"), 0), _
                                            AltIsDBNull(drSet.Item("POLICY_AMOUNT"), -1.0), _
                                            AltIsDBNull(drSet.Item("OVERRIDE_AMOUNT"), -1.0), _
                                            AltIsDBNull(drSet.Item("SETTLEMENT_AMOUNT"), -1.0), _
                                            AltIsDBNull(drSet.Item("PAID_AMOUNT"), -1.0), _
                                            AltIsDBNull(drSet.Item("DATE_RECEIVED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("WORKSHOP_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("WORKSHOP_RESULT"), 0), _
                                            AltIsDBNull(drSet.Item("WORKSHOP_REQUIRED"), False), _
                                            AltIsDBNull(drSet.Item("SHOW_CAUSE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("SHOW_CAUSE_RESULTS"), 0), _
                                            AltIsDBNull(drSet.Item("COMMISSION_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("COMMISSION_RESULTS"), 0), _
                                            AltIsDBNull(drSet.Item("AGREED_ORDER"), String.Empty), _
                                            AltIsDBNull(drSet.Item("ADMINISTRATIVE_ORDER"), String.Empty), _
                                            AltIsDBNull(drSet.Item("PENDING_LETTER"), 0), _
                                            AltIsDBNull(drSet.Item("LETTER_GENERATED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LETTER_PRINTED"), False), _
                                            AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("DELETED"), False), _
                                            String.Empty, _
                                            AltIsDBNull(drSet.Item("PENDING_LETTER_TEMPLATE_NUM"), 0), _
                                            AltIsDBNull(drSet.Item("ENSITE ID"), 0), AltIsDBNull(drSet.Item("COMMENTS"), String.Empty), _
                                            AltIsDBNull(drSet.Item("ADMIN_HEARING_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("ADMIN_HEARING_RESULTS"), 0))
                    colOCE.Add(oOCEInfo)
                End While
                Return colOCE
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Sub Put(ByVal flagNFA As Short, ByRef oOCEInfo As MUSTER.Info.OwnerComplianceEventInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.CAEOwnerComplianceEvent, Integer))) Then
                    returnVal = "You do not have rights to save Owner Compliance Event."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutCAEOCE"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)

                If oOCEInfo.ID <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = oOCEInfo.ID
                End If
                Params(1).Value = oOCEInfo.OwnerID
                Params(2).Value = oOCEInfo.Citation
                If Date.Compare(oOCEInfo.CitationDueDate, CDate("01/01/0001")) = 0 Then
                    Params(3).Value = DBNull.Value
                Else
                    Params(3).Value = oOCEInfo.CitationDueDate
                End If

                Params(4).Value = oOCEInfo.Rescinded
                Params(5).Value = oOCEInfo.OCEPath

                If Date.Compare(oOCEInfo.OCEDate, CDate("01/01/0001")) = 0 Then
                    Params(6).Value = DBNull.Value
                Else
                    Params(6).Value = oOCEInfo.OCEDate
                End If
                If Date.Compare(oOCEInfo.OCEProcessDate, CDate("01/01/0001")) = 0 Then
                    Params(7).Value = DBNull.Value
                Else
                    Params(7).Value = oOCEInfo.OCEProcessDate
                End If
                If Date.Compare(oOCEInfo.NextDueDate, CDate("01/01/0001")) = 0 Then
                    Params(8).Value = DBNull.Value
                Else
                    Params(8).Value = oOCEInfo.NextDueDate
                End If
                If Date.Compare(oOCEInfo.OverrideDueDate, CDate("01/01/0001")) = 0 Then
                    Params(9).Value = DBNull.Value
                Else
                    Params(9).Value = oOCEInfo.OverrideDueDate
                End If

                Params(10).Value = oOCEInfo.OCEStatus
                Params(11).Value = oOCEInfo.Escalation
                Params(12).Value = IIf(oOCEInfo.PolicyAmount = -1.0, DBNull.Value, oOCEInfo.PolicyAmount)
                Params(13).Value = IIf(oOCEInfo.OverRideAmount = -1.0, DBNull.Value, oOCEInfo.OverRideAmount)
                Params(14).Value = IIf(oOCEInfo.SettlementAmount = -1.0, DBNull.Value, oOCEInfo.SettlementAmount)
                Params(15).Value = IIf(oOCEInfo.PaidAmount = -1.0, DBNull.Value, oOCEInfo.PaidAmount)

                If Date.Compare(oOCEInfo.DateReceived, CDate("01/01/0001")) = 0 Then
                    Params(16).Value = DBNull.Value
                Else
                    Params(16).Value = oOCEInfo.DateReceived
                End If
                If Date.Compare(oOCEInfo.WorkShopDate, CDate("01/01/0001")) = 0 Then
                    Params(17).Value = DBNull.Value
                Else
                    Params(17).Value = oOCEInfo.WorkShopDate
                End If

                Params(18).Value = oOCEInfo.WorkShopResult
                Params(19).Value = oOCEInfo.WorkshopRequired

                If Date.Compare(oOCEInfo.ShowCauseDate, CDate("01/01/0001")) = 0 Then
                    Params(20).Value = DBNull.Value
                Else
                    Params(20).Value = oOCEInfo.ShowCauseDate
                End If
                Params(21).Value = oOCEInfo.ShowCauseResult

                If Date.Compare(oOCEInfo.CommissionDate, CDate("01/01/0001")) = 0 Then
                    Params(22).Value = DBNull.Value
                Else
                    Params(22).Value = oOCEInfo.CommissionDate
                End If
                Params(23).Value = oOCEInfo.CommissionResult
                Params(24).Value = oOCEInfo.AgreedOrder
                Params(25).Value = oOCEInfo.AdministrativeOrder
                Params(26).Value = oOCEInfo.PendingLetter

                If Date.Compare(oOCEInfo.LetterGenerated, CDate("01/01/0001")) = 0 Then
                    Params(27).Value = DBNull.Value
                Else
                    Params(27).Value = oOCEInfo.LetterGenerated
                End If

                Params(28).Value = oOCEInfo.LetterPrinted
                Params(29).Value = DBNull.Value
                Params(30).Value = DBNull.Value
                Params(31).Value = DBNull.Value
                Params(32).Value = DBNull.Value
                Params(33).Value = oOCEInfo.Deleted
                Params(34).Value = oOCEInfo.EscalationString

                If oOCEInfo.ID <= 0 Then
                    Params(35).Value = oOCEInfo.CreatedBy
                Else
                    Params(35).Value = oOCEInfo.ModifiedBy
                End If
                Params(36).Value = oOCEInfo.PendingLetterTemplateNum
                Params(37).Value = oOCEInfo.EnsiteID

                If oOCEInfo.Comments <> String.Empty Then
                    Params(38).Value = oOCEInfo.Comments
                End If

                If Date.Compare(oOCEInfo.AdminHearingDate, CDate("01/01/0001")) = 0 Then
                    Params(39).Value = DBNull.Value
                Else
                    Params(39).Value = oOCEInfo.AdminHearingDate
                End If

                Params(40).Value = oOCEInfo.AdminHearingResult
                If Date.Compare(oOCEInfo.RedTagDate, CDate("01/01/0001")) = 0 Then
                    Params(41).Value = DBNull.Value
                Else
                    Params(41).Value = oOCEInfo.RedTagDate
                End If

                Params(42).Value = flagNFA

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If oOCEInfo.ID <= 0 Then
                    oOCEInfo.ID = Params(0).Value
                End If
                oOCEInfo.CreatedBy = Params(29).Value
                oOCEInfo.CreatedOn = Params(30).Value
                oOCEInfo.ModifiedBy = AltIsDBNull(Params(31).Value, String.Empty)
                oOCEInfo.ModifiedOn = AltIsDBNull(Params(32).Value, CDate("01/01/0001"))
                oOCEInfo.EscalationString = AltIsDBNull(Params(34).Value, String.Empty)
                oOCEInfo.PendingLetterTemplateNum = AltIsDBNull(Params(36).Value, 0)
                oOCEInfo.Escalation = AltIsDBNull(Params(11).Value, oOCEInfo.OCEStatus)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
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
        Public Function DBOwnersPriorViolations(ByVal ownerID As Integer, Optional ByVal excludeOceID As Integer = 0) As DataSet
            Dim strSQL As String
            Dim Params As Collection
            Dim dsData As DataSet
            Try
                strSQL = "spGetCAEOwnersPriorViolations"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@OWNER_ID").Value = ownerID
                Params("@OCE_ID").Value = excludeOceID

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
                'If drSet.HasRows Then
                '    drSet.Read()
                '    Return drSet.Item("PREVIOUSVIOLATIONS")
                'Else
                '    Return False
                'End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetOwnerSize(ByVal ownerID As Integer) As String
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetCAEOwnerSize"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@OWNER_ID").Value = ownerID

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return drSet.Item("OWNERSIZE")
                Else
                    Return String.Empty
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBGetEnforcements(Optional ByVal ownerID As Integer = 0, Optional ByVal showDeleted As Boolean = False, Optional ByVal returnOCEs As Boolean = False, Optional ByVal oceStatus As Integer = 0, Optional ByVal facility_id As Integer = 0, Optional ByVal UseAllOwnersFormat As Boolean = False, Optional ByVal managerID As Integer = Nothing) As DataSet
            Dim drSet As SqlDataReader
            Dim dsData As DataSet
            Dim strSQl As String
            Dim Params() As SqlParameter
            Try
                strSQl = "spGetCAEEnforcements"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQl)
                Params(0).Value = IIf(ownerID <= 0, DBNull.Value, ownerID)
                Params(1).Value = showDeleted
                Params(2).Value = IIf(facility_id <= 0, DBNull.Value, facility_id)
                Params(3).Value = IIf(UseAllOwnersFormat, 1, 0)
                'Params(2).Value = returnOCEs
                'Params(3).Value = IIf(oceStatus <= 0, DBNull.Value, oceStatus)

                If Not managerID = Nothing Then
                    Params(4).Value = managerID
                End If

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQl, Params)
                Return dsData
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetPriorEnforcements(Optional ByVal ownerID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim drSet As SqlDataReader
            Dim dsData As DataSet
            Dim strSQl As String
            Dim Params() As SqlParameter
            Try
                strSQl = "spGetCAEPriorEnforcements"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQl)
                Params(0).Value = IIf(ownerID <= 0, DBNull.Value, ownerID)
                Params(1).Value = showDeleted

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQl, Params)
                Return dsData
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function DBGetMyRedTagStatusChanges(ByVal OCEChangeDate As Date, ByVal user As String) As DataSet
            Dim drSet As SqlDataReader
            Dim dsData As DataSet
            Dim strSQl As String
            Dim Params() As SqlParameter
            Try
                strSQl = "spListProhibitabletanksByCNE"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQl)
                Params(0).Value = OCEChangeDate
                Params(1).Value = user

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQl, Params)
                Return dsData

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function DBGetCOIFacs(ByVal oceID As Integer, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim strSQL As String
            Dim Params As Collection
            Dim dsData As DataSet
            Try
                strSQL = "spGetCAECOI"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@OCE_ID").Value = oceID
                Params("@DELETED").Value = showDeleted

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetCAEOCEEscalation(ByVal flagNFA As Short, _
                                                ByVal oceID As Integer, _
                                                ByVal oceStatus As Integer, _
                                                ByVal ownerID As Integer, _
                                                ByVal nextDueDate As Date, _
                                                ByVal overrideDueDate As Date, _
                                                ByVal policyAmount As Decimal, _
                                                ByVal overrideAmount As Decimal, _
                                                ByVal settlementAmount As Decimal, _
                                                ByVal paidAmount As Decimal, _
                                                ByVal workshopRequired As Boolean, _
                                                ByVal workshopDate As Date, _
                                                ByVal workshopResult As Integer, _
                                                ByVal showCauseDate As Date, _
                                                ByVal showCauseResult As Integer, _
                                                ByVal commissionDate As Date, _
                                                ByVal commissionResult As Integer, _
                                                ByVal pendingLetter As Integer, _
                                                ByVal citationDueDate As Date, _
                                                ByVal ocePath As Integer, _
                                                ByVal dateReceived As Date, _
                                                ByRef escalationID As Integer, Optional ByVal admindate As Object = Nothing, Optional ByVal adminresult As Object = Nothing) As String
            Dim strSQL As String
            Dim Params() As SqlParameter
            Dim drSet As SqlDataReader
            Dim strReturnValue As String = String.Empty
            Try
                strSQL = "spGetCAEOCEEscalation"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = oceID
                Params(1).Value = Today.Date
                Params(2).Value = oceStatus
                Params(3).Value = ownerID
                If Date.Compare(nextDueDate, CDate("01/01/0001")) = 0 Then
                    Params(4).Value = DBNull.Value
                Else
                    Params(4).Value = nextDueDate
                End If
                If Date.Compare(overrideDueDate, CDate("01/01/0001")) = 0 Then
                    Params(5).Value = DBNull.Value
                Else
                    Params(5).Value = overrideDueDate
                End If
                Params(6).Value = policyAmount
                Params(7).Value = overrideAmount
                Params(8).Value = settlementAmount
                Params(9).Value = paidAmount
                Params(10).Value = workshopRequired
                If Date.Compare(workshopDate, CDate("01/01/0001")) = 0 Then
                    Params(11).Value = DBNull.Value
                Else
                    Params(11).Value = workshopDate
                End If
                Params(12).Value = workshopResult
                If Date.Compare(showCauseDate, CDate("01/01/0001")) = 0 Then
                    Params(13).Value = DBNull.Value
                Else
                    Params(13).Value = showCauseDate
                End If
                Params(14).Value = showCauseResult
                If Date.Compare(commissionDate, CDate("01/01/0001")) = 0 Then
                    Params(15).Value = DBNull.Value
                Else
                    Params(15).Value = commissionDate
                End If
                Params(16).Value = commissionResult
                Params(17).Value = pendingLetter
                If Date.Compare(citationDueDate, CDate("01/01/0001")) = 0 Then
                    Params(18).Value = DBNull.Value
                Else
                    Params(18).Value = citationDueDate
                End If
                Params(19).Value = ocePath
                If Date.Compare(dateReceived, CDate("01/01/0001")) = 0 Then
                    Params(20).Value = DBNull.Value
                Else
                    Params(20).Value = dateReceived
                End If
                Params(21).Value = escalationID
                Params(22).Value = String.Empty

                If Not admindate Is Nothing AndAlso TypeOf admindate Is Date Then
                    Params(23).Value = admindate
                End If

                If Not adminresult Is Nothing AndAlso TypeOf adminresult Is Integer Then
                    Params(24).Value = adminresult
                End If
                Params(25).Value = flagNFA
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                escalationID = AltIsDBNull(Params(21).Value, oceStatus)
                strReturnValue = AltIsDBNull(Params(22).Value, String.Empty)
                'drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                'While drSet.Read
                '    strReturnValue = AltIsDBNull(drSet.Item("ESCALATION"), String.Empty)
                'End While
                Return strReturnValue
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBExecuteCAEOCEEscalation(ByVal flagNFA As Short, _
                                                ByVal oceID As Integer, _
                                                ByRef oceStatus As Integer, _
                                                ByVal ownerID As Integer, _
                                                ByRef nextDueDate As Date, _
                                                ByRef overrideDueDate As Date, _
                                                ByVal policyAmount As Decimal, _
                                                ByVal overrideAmount As Decimal, _
                                                ByRef settlementAmount As Decimal, _
                                                ByVal paidAmount As Decimal, _
                                                ByVal workshopRequired As Boolean, _
                                                ByVal workshopDate As Date, _
                                                ByVal workshopResult As Integer, _
                                                ByRef showCauseDate As Date, _
                                                ByRef showCauseResult As Integer, _
                                                ByRef commissionDate As Date, _
                                                ByRef commissionResult As Integer, _
                                                ByRef pendingLetter As Integer, _
                                                ByVal citationDueDate As Date, _
                                                ByVal ocePath As Integer, _
                                                ByVal dateReceived As Date, _
                                                ByVal userProvidedDate As Date, _
                                                ByRef stroceStatus As String, _
                                                ByRef strpendingLetter As String, _
                                                ByRef strEscalation As String, _
                                                ByRef pendingLetterTemplateNum As Integer, _
                                                ByRef escalationID As Integer, Optional ByVal admindate As Object = Nothing, Optional ByVal adminresult As Object = Nothing, Optional ByVal userID As String = "ADMIN") As String
            Dim strSQL As String
            Dim Params() As SqlParameter
            'Dim drSet As SqlDataReader
            Dim bolNotesIsEmpty As Boolean
            Try
                bolNotesIsEmpty = False
                strSQL = "spProcessCAEOCEEscalation"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = oceID
                Params(1).Value = Today.Date
                Params(2).Value = oceStatus
                Params(3).Value = ownerID
                If Date.Compare(nextDueDate, CDate("01/01/0001")) = 0 Then
                    Params(4).Value = DBNull.Value
                Else
                    Params(4).Value = nextDueDate
                End If
                If Date.Compare(overrideDueDate, CDate("01/01/0001")) = 0 Then
                    Params(5).Value = DBNull.Value
                Else
                    Params(5).Value = overrideDueDate
                End If
                Params(6).Value = policyAmount
                Params(7).Value = overrideAmount
                Params(8).Value = settlementAmount
                Params(9).Value = paidAmount
                Params(10).Value = workshopRequired
                If Date.Compare(workshopDate, CDate("01/01/0001")) = 0 Then
                    Params(11).Value = DBNull.Value
                Else
                    Params(11).Value = workshopDate
                End If
                Params(12).Value = workshopResult
                If Date.Compare(showCauseDate, CDate("01/01/0001")) = 0 Then
                    Params(13).Value = DBNull.Value
                Else
                    Params(13).Value = showCauseDate
                End If
                Params(14).Value = showCauseResult
                If Date.Compare(commissionDate, CDate("01/01/0001")) = 0 Then
                    Params(15).Value = DBNull.Value
                Else
                    Params(15).Value = commissionDate
                End If
                Params(16).Value = commissionResult
                Params(17).Value = pendingLetter
                If Date.Compare(citationDueDate, CDate("01/01/0001")) = 0 Then
                    Params(18).Value = DBNull.Value
                Else
                    Params(18).Value = citationDueDate
                End If
                Params(19).Value = ocePath
                If Date.Compare(dateReceived, CDate("01/01/0001")) = 0 Then
                    Params(20).Value = DBNull.Value
                Else
                    Params(20).Value = dateReceived
                End If
                If Date.Compare(userProvidedDate, CDate("01/01/0001")) = 0 Then
                    Params(21).Value = DBNull.Value
                Else
                    Params(21).Value = userProvidedDate
                End If
                Params(22).Value = String.Empty
                Params(23).Value = stroceStatus
                Params(24).Value = strpendingLetter
                Params(25).Value = strEscalation
                Params(26).Value = pendingLetterTemplateNum
                Params(27).Value = escalationID

                If Not admindate Is Nothing AndAlso TypeOf admindate Is Date Then
                    Params(28).Value = admindate
                End If

                If Not adminresult Is Nothing AndAlso TypeOf adminresult Is Integer Then
                    Params(29).Value = adminresult
                End If

                Params(30).Value = userID
                Params(31).Value = flagNFA

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If Params(22).Value Is DBNull.Value Then
                    bolNotesIsEmpty = True
                ElseIf Params(22).Value = String.Empty Then
                    bolNotesIsEmpty = True
                Else
                    bolNotesIsEmpty = False
                End If
                If bolNotesIsEmpty Then
                    oceStatus = AltIsDBNull(Params(2).Value, oceStatus)
                    nextDueDate = AltIsDBNull(Params(4).Value, CDate("01/01/0001"))
                    overrideDueDate = AltIsDBNull(Params(5).Value, CDate("01/01/0001"))
                    'policyAmount = AltIsDBNull(Params("@POLICY_AMOUNT"), 0.0)
                    'overrideAmount = AltIsDBNull(Params("@OVERRIDE_AMOUNT"), 0.0)
                    settlementAmount = AltIsDBNull(Params(8).Value, 0.0)
                    'paidAmount = AltIsDBNull(Params("@PAID_AMOUNT"), 0.0)
                    'workshopRequired = AltIsDBNull(Params("@WORKSHOP_REQUIRED"), False)
                    'workshopDate = AltIsDBNull(Params("@WORKSHOP_DATE"), CDate("01/01/0001"))
                    'workshopResult = AltIsDBNull(Params("@WORKSHOP_RESULT"), 0)
                    showCauseDate = AltIsDBNull(Params(13).Value, CDate("01/01/0001"))
                    showCauseResult = AltIsDBNull(Params(14).Value, 0)
                    commissionDate = AltIsDBNull(Params(15).Value, CDate("01/01/0001"))
                    commissionResult = AltIsDBNull(Params(16).Value, 0)
                    pendingLetter = AltIsDBNull(Params(17).Value, 0)
                    'citationDueDate = AltIsDBNull(Params("@CITATION_DUE_DATE"), CDate("01/01/0001"))
                    'ocePath = AltIsDBNull(Params("@OCE_PATH"), 0)
                    'dateReceived = AltIsDBNull(Params("@DATE_RECEIVED"), CDate("01/01/0001"))
                    stroceStatus = AltIsDBNull(Params(23).Value, String.Empty)
                    strpendingLetter = AltIsDBNull(Params(24).Value, String.Empty)
                    strEscalation = AltIsDBNull(Params(25).Value, String.Empty)
                    pendingLetterTemplateNum = AltIsDBNull(Params(26).Value, 0)
                    escalationID = AltIsDBNull(Params(27).Value, oceStatus)
                    Return String.Empty
                Else
                    Return Params(22).Value
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetLetterGeneratedDate(ByVal entityID As Integer, ByVal entityType As Integer, Optional ByVal propertyIDTemplateNum As Integer = 0, Optional ByVal showDeleted As Boolean = False) As Date
            Dim retVal As Date = CDate("01/01/0001")
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetCAELetterGenDate"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@ENTITY_ID").Value = entityID
                Params("@ENTITY_TYPE").Value = entityType
                Params("@PROPERTY_ID_TEMPLATE_NUM").Value = IIf(propertyIDTemplateNum = 0, DBNull.Value, propertyIDTemplateNum)
                Params("@DELETED").Value = showDeleted

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    retVal = AltIsDBNull(drSet.Item("GENERATED_DATE"), retVal)
                End If
                Return retVal
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Sub DBPutLetterGeneratedDate(ByRef letterGenID As Integer, ByVal entityID As Integer, ByVal entityType As Integer, ByVal propertyIDTemplateNum As Integer, ByVal generatedDate As Date, ByVal deleted As Boolean, ByVal documentID As Integer, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try
                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.CAEOwnerComplianceEvent, Integer))) Then
                    returnVal = "You do not have rights to save Letter Generated Date."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutCAELetterGenDate"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)

                If letterGenID <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = letterGenID
                End If
                Params(1).Value = entityID
                Params(2).Value = entityType
                Params(3).Value = propertyIDTemplateNum
                If Date.Compare(generatedDate, CDate("01/01/0001")) = 0 Then
                    Params(4).Value = DBNull.Value
                Else
                    Params(4).Value = generatedDate
                End If
                If staffID <= 0 Then
                    Params(5).Value = DBNull.Value
                Else
                    Params(5).Value = staffID
                End If
                Params(6).Value = deleted
                Params(7).Value = documentID

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If letterGenID <= 0 Then
                    letterGenID = Params(0).Value
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function DBGetCAEPrevWorkshopDate(ByVal ownerID As Integer, Optional ByVal showdeleted As Boolean = False, Optional ByVal excludeOceID As Integer = 0) As Boolean
            Dim retVal As Boolean = False
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetCAEPrevWorkshopDate"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@OWNER_ID").Value = ownerID
                Params("@DELETED").Value = showdeleted
                Params("@OCE_ID").Value = excludeOceID

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    retVal = AltIsDBNull(drSet.Item("HAD_WORKSHOP_OPTION"), retVal)
                End If
                Return retVal
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBGetOpenOCEDocumentInfoforOwner(ByVal ownerID As Integer) As String
            Dim retVal As String = String.Empty
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetCAEOpenOCEDocumentInfo"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@OWNER_ID").Value = ownerID

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    retVal = AltIsDBNull(drSet.Item("DOCINFO"), retVal)
                End If
                Return retVal
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBOwnerHasWorkshopOCEDuringPast90Days(ByVal ownerID As Integer, Optional ByVal excludeOceID As Integer = 0) As Boolean
            Dim retVal As Boolean = False
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetCAEOwnerHasWorkshopDuringPast90Days"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@OWNER_ID").Value = ownerID
                Params("@EXCLUDE_OCE_ID").Value = excludeOceID

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    retVal = AltIsDBNull(drSet.Item("HAS_WORKSHOP"), retVal)
                End If
                Return retVal
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
    End Class
End Namespace
