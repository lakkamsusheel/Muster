
Imports System.Data.SqlClient

Imports Utils.DBUtils

Namespace MUSTER.DataAccess

    Public Class FinancialDB


#Region "Private Member Variables"
        Private _strConn As Object
        Private MusterException As Exceptions.MusterExceptions = New MUSTER.Exceptions.MusterExceptions
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

        ' Operation to return an INFO object by sending in the ID
        Public Function DBGetByID(ByVal nVal As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FinancialInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nVal = 0 Then
                    Return New MUSTER.Info.FinancialInfo
                End If
                strSQL = "spGetFinancialEvent"
                strVal = nVal

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@FIN_EVENT_ID").Value = nVal
                Params("@TEC_EVENT_ID").Value = DBNull.Value

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.FinancialInfo(drSet.Item("FIN_EVENT_ID"), _
                            AltIsDBNull(drSet.Item("FACILITY_SEQUENCE"), 0), _
                            AltIsDBNull(drSet.Item("TEC_EVENT_ID"), 0), _
                            AltIsDBNull(drSet.Item("FIN_START_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("VENDOR_ID"), 0), _
                            AltIsDBNull(drSet.Item("FIN_STATUS"), 0), _
                            AltIsDBNull(drSet.Item("FIN_CLOSED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("CREATE_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("Deleted"), 0), _
                            AltIsDBNull(drSet.Item("TEC_EVENT_IDDESC"), 0) _
                    )

                Else

                    Return New MUSTER.Info.FinancialInfo
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

        Public Function DBGetByTechID(ByVal nVal As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FinancialInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nVal = 0 Then
                    Return New MUSTER.Info.FinancialInfo
                End If
                strSQL = "spGetFinancialEvent"
                strVal = nVal

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@FIN_EVENT_ID").Value = DBNull.Value
                Params("@TEC_EVENT_ID").Value = nVal

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.FinancialInfo(drSet.Item("FIN_EVENT_ID"), _
                            AltIsDBNull(drSet.Item("FACILITY_SEQUENCE"), 0), _
                            AltIsDBNull(drSet.Item("TEC_EVENT_ID"), 0), _
                            AltIsDBNull(drSet.Item("FIN_START_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("VENDOR_ID"), 0), _
                            AltIsDBNull(drSet.Item("FIN_STATUS"), 0), _
                            AltIsDBNull(drSet.Item("FIN_CLOSED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("CREATE_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("Deleted"), 0), _
                            AltIsDBNull(drSet.Item("TEC_EVENT_IDDESC"), 0) _
                    )

                Else

                    Return New MUSTER.Info.FinancialInfo
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

        ' Operation to return an INFO object by sending in the Name
        Public Function DBGetByFacility(ByVal FacilityID As Int64, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FinancialCollection
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Dim colText As New MUSTER.Info.FinancialCollection

            Try

                strSQL = "spGetFinancialEvent_ByFacility"
                strVal = FacilityID

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@FacilityID").Value = DBNull.Value


                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Dim otmpObject As New MUSTER.Info.FinancialInfo(drSet.Item("FIN_EVENT_ID"), _
                            AltIsDBNull(drSet.Item("FACILITY_SEQUENCE"), 0), _
                            AltIsDBNull(drSet.Item("TEC_EVENT_ID"), 0), _
                            AltIsDBNull(drSet.Item("FIN_START_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("VENDOR_ID"), 0), _
                            AltIsDBNull(drSet.Item("FIN_STATUS"), 0), _
                            AltIsDBNull(drSet.Item("FIN_CLOSED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("CREATE_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("Deleted"), 0), _
                            AltIsDBNull(drSet.Item("TEC_EVENT_IDDESC"), 0) _
                    )


                    colText.Add(otmpObject)
                End If
                Return colText
            Catch ex As Exception

                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function

        ' Operation to send the INFO object to the repository
        Public Sub Put(ByRef oFinInfo As MUSTER.Info.FinancialInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal updateLastEditedBy As Boolean = True, Optional ByVal fromTechnical As Boolean = False)
            Try

                If Not fromTechnical And Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.FinancialEvent, Integer))) Then
                    returnVal = "You do not have rights to save Financial."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutFinancial")

                With oFinInfo
                    If .ID = 0 Or .ID = -1 Then
                        Params(0).Value = 0 'System.DBNull.Value
                    Else
                        Params(0).Value = .ID
                    End If
                    Params(1).Value = .Sequence
                    Params(2).Value = .TecEventID
                    Params(3).Value = .StartDate
                    Params(4).Value = .VendorID
                    Params(5).Value = .Status
                    Params(6).Value = .Deleted
                    If .ID <= 0 Then
                        Params(7).Value = .CreatedBy
                    Else
                        Params(7).Value = .ModifiedBy
                    End If
                    Params(8).Value = System.DBNull.Value
                    Params(9).Value = System.DBNull.Value
                    Params(10).Value = System.DBNull.Value
                    Params(11).Value = System.DBNull.Value
                    Params(12).Value = updateLastEditedBy
                    If Date.Compare(.ClosedDate, CDate("01/01/0001")) = 0 Then
                        Params(13).Value = DBNull.Value
                    Else
                        Params(13).Value = .ClosedDate
                    End If
                End With

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutFinancial", Params)

                'Perform check for New ID and assign, if necessary
                If Params(0).Value <> oFinInfo.ID Then
                    oFinInfo.ID = Params(0).Value
                    oFinInfo.Sequence = Params(1).Value
                End If
                oFinInfo.ModifiedBy = AltIsDBNull(Params(8).Value, String.Empty)
                oFinInfo.ModifiedOn = AltIsDBNull(Params(9).Value, CDate("01/01/0001"))
                oFinInfo.CreatedBy = AltIsDBNull(Params(10).Value, String.Empty)
                oFinInfo.CreatedOn = AltIsDBNull(Params(11).Value, CDate("01/01/0001"))
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
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetProjectEngineer(ByVal nEventSequence As Integer, ByVal nFacilityID As Integer) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params As Collection

            Try

                strSQL = "spGetFINANCIALERAC"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@TEC_SEQUENCE").value = nEventSequence
                Params("@FACILITY_ID").Value = nFacilityID
                'Params("@DELETED").value = False

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region


    End Class

End Namespace
