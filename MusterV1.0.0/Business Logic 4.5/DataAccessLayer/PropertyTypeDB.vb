'-------------------------------------------------------------------------------
' MUSTER.DataAccess.PropertyTypeDB
'   Provides the means for marshalling Address to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       AN      01/03/05    Original class definition.(Added Comment Date is Invalid no header present)
'  1.1       AB      02/15/05    Replaced dynamic SQL with stored procedures in the following
'                                    Functions:  DBGetAllPropertyType, DBGetPropertyTypeByID, DBGetPropertyTypeByName
'  1.2       AB      02/16/05    Added Finally to the Try/Catch to close all datareaders
'  1.3       AB      02/18/05    Set all parameters for SP, that are not required, to NULL
'  1.4       JC      06/02/05    Added new optional MaintenanceMode parameter to getAvailableProperties.
'
'
' Function                                      Description
'   DBGetAllPropertyTypeByEntity(EntityID)          Returns all property types by the entity id
'   DBGetAllPropertyType()                          Returns all property types
'   DBGetPropertyTypeByID(nVal)                     Return Property Type by property type id
'   DBGetPropertyTypeByName(nVal)                   Return Property type by property type name
'   DBGetDS(strSql)                                 Returns a data set based on the sql string provided.
'   Put(oPropInfo)                                  Calls the "put" stored proc to create and update property types
'-------------------------------------------------------------------------------

Imports System.Data.SqlClient
Imports Utils.DBUtils
Namespace MUSTER.DataAccess
    Public Class PropertyTypeDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
        Public Function DBGetAllPropertyTypeByEntity(ByVal EntityID As Int64) As Muster.Info.PropertyTypeCollection
            Dim dr As SqlDataReader
            Try
                Dim arParams() As SqlParameter = New SqlParameter(0) {}
                arParams(0) = New SqlParameter("@EntityID", EntityID)
                SqlHelperParameterCache.CacheParameterSet(_strConn, "spGetPropertyTypesForEntity", arParams)
                dr = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, "spGetPropertyTypesForEntity", arParams)
                Dim colPropTypes As New Muster.Info.PropertyTypeCollection
                If dr.HasRows Then
                    While dr.Read
                        Dim oMusterPropTypes As New Muster.Info.PropertyTypeInfo(dr.Item("Entity_ID"), _
                             dr.Item("PROPERTY_TYPE_ID"), _
                             dr.Item("PROPERTY_TYPE_NAME"), _
                             dr.Item("CREATED_BY"), _
                             dr.Item("DATE_CREATED"))
                        colPropTypes.Add(oMusterPropTypes)

                    End While
                    Return colPropTypes
                End If
            Catch ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not dr.IsClosed Then dr.Close()
            End Try
        End Function
        Public Function DBGetAllPropertyType() As Muster.Info.PropertyTypeCollection
            Dim dr As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetSYSProperty_Type"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Property_Type_ID").Value = DBNull.Value
                Params("@Property_Type_Name").Value = DBNull.Value
                Params("@Entity_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 1

                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "select * from tblSYS_PROPERTY_MASTER ORDER BY EVENT_ID")

                dr = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colPropTypes As New MUSTER.Info.PropertyTypeCollection
                If dr.HasRows Then
                    While dr.Read
                        Dim oMusterPropTypes As New MUSTER.Info.PropertyTypeInfo(dr.Item("Entity_ID"), _
                             dr.Item("PROPERTY_TYPE_ID"), _
                             dr.Item("PROPERTY_TYPE_NAME"), _
                             dr.Item("CREATED_BY"), _
                             dr.Item("DATE_CREATED"))
                        colPropTypes.Add(oMusterPropTypes)
                    End While
                    Return colPropTypes
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not dr.IsClosed Then dr.Close()
            End Try
        End Function
        Public Function DBGetPropertyTypeByID(ByVal nVal As Int16) As Muster.Info.PropertyTypeInfo
            Dim drSet As SqlDataReader
            Dim Params As Collection
            Dim strSQL As String

            Try
                strSQL = "spGetSYSProperty_Type"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Property_Type_Name").Value = DBNull.Value
                Params("@Entity_ID").Value = DBNull.Value
                Params("@Property_Type_ID").Value = nVal
                Params("@OrderBy").Value = 1

                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_PROPERTY_TYPE WHERE PROPERTY_TYPE_ID = " & nVal.ToString)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.PropertyTypeInfo(drSet.Item("Entity_ID"), _
                          drSet.Item("PROPERTY_TYPE_ID"), _
                          drSet.Item("PROPERTY_TYPE_NAME"), _
                          drSet.Item("CREATED_BY"), _
                          drSet.Item("DATE_CREATED"))
                Else
                    Return New MUSTER.Info.PropertyTypeInfo
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetPropertyTypeByName(ByVal nVal As String) As Muster.Info.PropertyTypeInfo
            Dim drSet As SqlDataReader
            Dim Params As Collection
            Dim strSQL As String

            Try
                strSQL = "spGetSYSProperty_Type"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Property_Type_ID").Value = DBNull.Value
                Params("@Entity_ID").Value = DBNull.Value
                Params("@Property_Type_Name").Value = nVal
                Params("@OrderBy").Value = 1

                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_PROPERTY_TYPE WHERE PROPERTY_TYPE_NAME = '" & nVal & "'")

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.PropertyTypeInfo(drSet.Item("Entity_ID"), _
                          drSet.Item("PROPERTY_TYPE_ID"), _
                          drSet.Item("PROPERTY_TYPE_NAME"), _
                          drSet.Item("CREATED_BY"), _
                          drSet.Item("DATE_CREATED"))
                Else
                    Return New MUSTER.Info.PropertyTypeInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
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
        Public Sub Put(ByRef oPropInfo As MUSTER.Info.PropertyTypeInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.MUSTER, Integer))) Then
                    returnVal = "You do not have rights to save a Property."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params(2) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutSYSPROPERTYTYPE")
                Params(0).Value = oPropInfo.EntityId
                Params(1).Value = oPropInfo.ID
                Params(2).Value = oPropInfo.Name
                If oPropInfo.ID <= 0 Then
                    Params(3).Value = oPropInfo.CreatedBy
                Else
                    Params(3).Value = oPropInfo.ModifiedBy
                End If
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutSYSPROPERTYTYPE", Params)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub


        Public Function PutPropertyRelation(ByVal dtPropertyRel As DataTable, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String) As Boolean

            Dim sqlcmd As New SqlClient.SqlCommand
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.MUSTER, Integer))) Then
                    returnVal = "You do not have rights to save Property Relation."
                    Exit Function
                Else
                    returnVal = String.Empty
                End If

                sqlcmd.CommandText = "spPutPropertyRelation"
                sqlcmd.CommandType = CommandType.StoredProcedure
                sqlcmd.Parameters.Add(New SqlClient.SqlParameter("@Parent_Property_ID", SqlDbType.Int))
                sqlcmd.Parameters.Add(New SqlClient.SqlParameter("@Child_Property_ID", SqlDbType.Int))
                'sqlcmd.Parameters.Add(New SqlClient.SqlParameter("@UserID", SqlDbType.VarChar, 50))
                sqlcmd.Parameters("@Parent_Property_ID").SourceColumn = "Parent Property"
                sqlcmd.Parameters("@Child_Property_ID").SourceColumn = "Property ID"
                'sqlcmd.Parameters("@UserID").SourceColumn = "User ID"
                SqlHelper.SaveDataTable(_strConn, dtPropertyRel, sqlcmd)
                dtPropertyRel.AcceptChanges()
                Return True
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Public Function DeletePropertyRelation(ByVal Parent_Property_ID As Int64, ByVal Child_Property_ID As Int64) As Boolean
            Try
                Dim arParams() As SqlClient.SqlParameter = New SqlClient.SqlParameter(1) {}
                arParams(0) = New SqlClient.SqlParameter("@Parent_Property_ID", SqlDbType.Int)
                arParams(1) = New SqlClient.SqlParameter("@Child_Property_ID", SqlDbType.Int)
                arParams(0).Value = Parent_Property_ID
                arParams(1).Value = Child_Property_ID
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spDeletePropertyRelation", arParams)
                Return True
            Catch ex As Exception
                Throw ex
            End Try

        End Function


        Public Function getAvailableProperties(ByVal Parent_Property_ID As Int64, Optional ByVal Child_Parent_ID As Int64 = 0, Optional ByVal MaintenanceMode As Boolean = False) As DataTable
            Dim sqlcmd As New SqlClient.SqlCommand
            Try
                Dim arParams() As SqlClient.SqlParameter = New SqlClient.SqlParameter(2) {}
                arParams(0) = New SqlClient.SqlParameter("@Parent_Property_ID", SqlDbType.Int)
                arParams(1) = New SqlClient.SqlParameter("@Child_Property_ID", SqlDbType.Int)
                arParams(2) = New SqlClient.SqlParameter("@MaintenanceMode", SqlDbType.Bit)
                arParams(0).Value = Parent_Property_ID
                arParams(1).Value = Child_Parent_ID
                arParams(2).Value = IIf(MaintenanceMode, 0, 1)
                Dim ds As DataSet = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, "spGetAvailableProperties", arParams)
                If ds.Tables.Count > 0 Then
                    Return ds.Tables(0)
                End If
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Public Function getPropertyTypesbyEntity(ByVal EntityID As Integer) As DataTable

            Dim dtPropertyTypeList As DataTable
            Dim ds As DataSet
            Dim arParams() As SqlClient.SqlParameter = New SqlClient.SqlParameter(0) {}
            arParams(0) = New SqlClient.SqlParameter("@EntityID", EntityID)
            ds = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, "spGetPropertyTypesForEntity", arParams)
            Return ds.Tables(0)

        End Function
    End Class
End Namespace
