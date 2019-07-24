'-------------------------------------------------------------------------------
' MUSTER.DataAccess.RelationDB
'   Provides the means for marshalling Relation to/from the repository
'
'
' Release   Initials    Date        Description
'  1.0       Hua Cao     09/11/12    Original class definition.
' 
'
' Function                  Description
' 
''-------------------------------------------------------------------------------

Imports System.Data.SqlClient
Imports Utils.DBUtils
Imports System.Data.SqlTypes

Namespace MUSTER.DataAccess
    <Serializable()> _
        Public Class RelationDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
        Public Function DBGetByManagerID(Optional ByVal ManagerID As Integer = 0) As MUSTER.Info.ManagerFacRelationCollection
            Dim drSet As SqlDataReader

            Dim strSQL As String
            Dim Params As Collection
            Dim colMgrFacRelation As New MUSTER.Info.ManagerFacRelationCollection
            Try

                strSQL = "spGetCOMMgrFacRelations"


                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@MGRFACRELATION_ID").Value = DBNull.Value
                Params("@MANAGER_ID").Value = IIf(ManagerID = 0, DBNull.Value, ManagerID.ToString)
                Params("@Deleted").Value = False


                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    While drSet.Read

                        Dim MgrFacRelationInfo As New MUSTER.Info.ManagerFacRelationInfo(drSet.Item("MGRFACRELATION_ID"), _
                                                AltIsDBNull(drSet.Item("MANAGER_ID"), 0), _
                                                AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                                                AltIsDBNull(drSet.Item("RELATION_ID"), 0), _
                                                AltIsDBNull(drSet.Item("RELATION_DESC"), String.Empty), _
                                                AltIsDBNull(drSet.Item("DELETED"), False))
                        colMgrFacRelation.Add(MgrFacRelationInfo)
                    End While

                End If
                Return colMgrFacRelation
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not (drSet Is Nothing) Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBGetByID(ByVal MgrFacRelationID As Integer) As MUSTER.Info.ManagerFacRelationInfo
            Dim drSet As SqlDataReader

            Dim strSQL As String
            Dim Params As Collection

            Try
                If MgrFacRelationID <= 0 Then
                    Return New MUSTER.Info.ManagerFacRelationInfo
                End If
                strSQL = "spGetCOMMgrFacRelations"


                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)

                Params("@MgrFacRelation_ID").Value = MgrFacRelationID.ToString
                Params("@MANAGER_ID").Value = DBNull.Value
                Params("@Deleted").Value = False

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()

                    Return New MUSTER.Info.ManagerFacRelationInfo(drSet.Item("MgrFacRelation_ID"), _
                                            AltIsDBNull(drSet.Item("MANAGER_ID"), 0), _
                                            AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                                            AltIsDBNull(drSet.Item("RELATION_ID"), 0), _
                                            AltIsDBNull(drSet.Item("RELATION_DESC"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DELETED"), False))
                Else
                    Return New MUSTER.Info.ManagerFacRelationInfo
                End If

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not (drSet Is Nothing) Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function

        Public Sub Put(ByRef oRelation As MUSTER.Info.ManagerFacRelationInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Company, Integer))) Then
                    returnVal = "You do not have rights to save Manager Facility Relations."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutCOMMgrFacRelations")
                If oRelation.ID <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = oRelation.ID
                End If
                Params(1).Value = oRelation.ManagerID
                Params(2).Value = oRelation.FacilityID
                Params(3).Value = oRelation.RelationID
                Params(4).Value = oRelation.Deleted
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutCOMMgrFacRelations", Params)

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
    End Class
End Namespace
