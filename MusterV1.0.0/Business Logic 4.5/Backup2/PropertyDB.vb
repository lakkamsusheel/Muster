'-------------------------------------------------------------------------------
' MUSTER.DataAccess.PropertymasterDB
'   Provides the means for marshalling Address to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       AN      01/03/05    Original class definition.(Added Comment Date is Invalid no header present)
'  1.1       AB      02/15/05    Replaced dynamic SQL with stored procedures in the following
'                                  Functions:  DBGetAllProperties, DBGetPropertyByID, DBGetPropertyByName
'  1.2       AB      02/15/05    Added Finally to the Try/Catch to close all datareaders
'  1.3       AB      02/16/05    Removed any IsNull calls for fields the DB requires
'  1.4       AB      02/18/05    Set all parameters for SP, that are not required, to NULL
'  1.5       JC      02/21/05    Modified DBGetAllChildPropertiesByParentID to take Parent Property ID and Type
'  1.6       AB      02/28/05      Modified Get functions based upon changes made to 
'                                     make several nullable fields non-nullable
'
'
' Function                                      Description
'   DBGetAllProperties()                            Return all properties 
'   DBgetAvailableProperties _
'    (Parent_Property_ID ,Child_Parent_ID )         Return all avaliable properties based on parent and child id.
'   DBGetAllPropByTypeID(nVal)                      Return all properties by property type
'   DBGetAllChildPropertiesByParentID(nVal)         Return all properties by parent id
'   DBGetPropertyByID(nVal)                         Return Property by property id
'   DBGetPropertyByName(nVal)                       Return Property by property name
'   DBGetDS(strSql)                                 Returns a data set based on the sql string provided.
'   Put(oPropInfo)                                  Calls the "put" stored proc to create and update properties
'-------------------------------------------------------------------------------

Imports System.Data.SqlClient
Imports Utils.DBUtils
Namespace MUSTER.DataAccess
    Public Class PropertymasterDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
        Public Function DBGetAllProperties() As Muster.Info.PropertyCollection
            Dim colMusterProperties As New MUSTER.Info.PropertyCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetSYSProperty"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Property_Name").Value = DBNull.Value
                Params("@Property_Active").Value = DBNull.Value
                Params("@Property_ID").Value = DBNull.Value

                Params("@OrderBy").Value = 1

                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_PROPERTY_MASTER ORDER BY PROPERTY_ID")
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                While drSet.Read
                    Dim oMusterPropertyInfo As New MUSTER.Info.PropertyInfo( _
                                drSet.Item("PROPERTY_ID"), _
                                0, _
                                drSet.Item("PROPERTY_TYPE_ID"), _
                                drSet.Item("PROPERTY_NAME"), _
                                AltIsDBNull(drSet.Item("PROPERTY_DESCRIPTION"), String.Empty), _
                                AltIsDBNull(drSet.Item("PROPERTY_POSITION"), 1), _
                                AltIsDBNull(drSet.Item("BUSINESS_TAG"), "0"), _
                                IIf(drSet.Item("PROPERTY_ACTIVE").ToString.ToUpper = "YES", True, False), _
                                drSet.Item("CREATED_BY"), _
                                drSet.Item("CREATE_DATE"), _
                                AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("1/1/1900")))
                    colMusterProperties.Add(oMusterpropertyInfo)
                End While
                Return colmusterProperties
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBgetAvailableProperties(ByVal Parent_Property_ID As Int64, Optional ByVal Child_Parent_ID As Int64 = 0) As Muster.Info.PropertyCollection
            Dim colmusterProperties As New MUSTER.Info.PropertyCollection
            Dim drSet As SqlDataReader

            Try
                Dim arParams() As SqlParameter = New SqlParameter(1) {}
                arParams(0) = New SqlParameter("@Parent_Property_ID", Parent_Property_ID)
                arParams(1) = New SqlParameter("@Child_Property_ID", Child_Parent_ID)
                SqlHelperParameterCache.CacheParameterSet(_strConn, "spGetAvailableProperties", arParams)
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, "spGetAvailableProperties", arParams)
                If drSet.HasRows Then
                    While drSet.Read
                        Dim omusterpropertyInfo As New MUSTER.Info.PropertyInfo( _
                                    drSet.Item("PROPERTY_ID"), _
                                    (AltIsDBNull(drSet.Item("PARENT_ID"), 0)), _
                                    drSet.Item("PROPERTY_TYPE_ID"), _
                                    drSet.Item("PROPERTY_NAME"), _
                                    AltIsDBNull(drSet.Item("PROPERTY_DESCRIPTION"), String.Empty), _
                                    AltIsDBNull(drSet.Item("PROPERTY_POSITION"), 1), _
                                    AltIsDBNull(drSet.Item("BUSINESS_TAG"), "0"), _
                                    IIf(drSet.Item("PROPERTY_ACTIVE").ToString.ToUpper = "YES", True, False), _
                                    drSet.Item("CREATED_BY"), _
                                    drSet.Item("CREATE_DATE"), _
                                    (AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty)), _
                                    (AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("1/1/1900"))))
                        colmusterProperties.Add(omusterpropertyInfo)
                    End While
                End If
                Return colmusterProperties
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function

        Public Function DBGetAllPropByTypeID_DS(ByVal nval As Int64) As DataSet
            Dim colParams As Collection
            Try
                Dim ds As DataSet
                Dim arParams() As SqlClient.SqlParameter = New SqlClient.SqlParameter(0) {}
                arParams(0) = New SqlClient.SqlParameter("@PROPERTY_TYPE_ID", nval)
                ds = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, "sp_GET_PROPERTIES_AND_CHILDREN", arParams)
                Return ds
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function DBGetAllPropByTypeID(ByVal nval As Int64) As MUSTER.Info.PropertyCollection
            Dim colmusterProperties As New MUSTER.Info.PropertyCollection
            Dim drSet As SqlDataReader
            Dim colParams As Collection
            Try
                colParams = SqlHelperParameterCache.GetSpParameterCol(_strConn, "sp_GET_PROPERTIES_AND_CHILDREN", False)
                colParams("@PROPERTY_TYPE_ID").VALUE = nval
                colParams.Remove("@PROPERTY_TYPE_NAME")
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, "sp_GET_PROPERTIES_AND_CHILDREN", colParams)
                'drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_PROPERTY_MASTER where ")
                If drSet.HasRows Then
                    While drSet.Read
                        Dim omusterpropertyInfo As New MUSTER.Info.PropertyInfo( _
                            drSet.Item("PROPERTY_ID"), _
                            AltIsDBNull(drSet.Item("PARENT_ID"), 0), _
                            drSet.Item("PROPERTY_TYPE_ID"), _
                            drSet.Item("PROPERTY_NAME"), _
                            AltIsDBNull(drSet.Item("PROPERTY_DESCRIPTION"), String.Empty), _
                            AltIsDBNull(drSet.Item("PROPERTY_POSITION"), 1), _
                            AltIsDBNull(drSet.Item("BUSINESS_TAG"), "0"), _
                            IIf(drSet.Item("PROPERTY_ACTIVE").ToString.ToUpper = "YES", True, False), _
                            drSet.Item("CREATED_BY"), _
                            drSet.Item("CREATE_DATE"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("1/1/1900")))
                        colmusterProperties.Add(omusterpropertyInfo)
                    End While
                End If
                Return colmusterProperties
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try

        End Function
        Public Function DBGetAllPropByType(ByVal strType As String) As MUSTER.Info.PropertyCollection
            Dim colmusterProperties As New MUSTER.Info.PropertyCollection
            Dim drSet As SqlDataReader
            Dim colParams As Collection
            Try
                colParams = SqlHelperParameterCache.GetSpParameterCol(_strConn, "sp_GET_PROPERTIES_AND_CHILDREN", False)
                colParams("@PROPERTY_TYPE_NAME").Value = strType
                colParams.Remove("@PROPERTY_TYPE_ID")
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, "sp_GET_PROPERTIES_AND_CHILDREN", colParams)
                'drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_PROPERTY_MASTER where ")
                If drSet.HasRows Then
                    While drSet.Read
                        Dim omusterpropertyInfo As New MUSTER.Info.PropertyInfo( _
                            drSet.Item("PROPERTY_ID"), _
                            AltIsDBNull(drSet.Item("PARENT_ID"), 0), _
                            drSet.Item("PROPERTY_TYPE_ID"), _
                            drSet.Item("PROPERTY_NAME"), _
                            AltIsDBNull(drSet.Item("PROPERTY_DESCRIPTION"), String.Empty), _
                            AltIsDBNull(drSet.Item("PROPERTY_POSITION"), 1), _
                            AltIsDBNull(drSet.Item("BUSINESS_TAG"), "0"), _
                            IIf(drSet.Item("PROPERTY_ACTIVE").ToString.ToUpper = "YES", True, False), _
                            drSet.Item("CREATED_BY"), _
                            drSet.Item("CREATE_DATE"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("1/1/1900")))
                        colmusterProperties.Add(omusterpropertyInfo)
                    End While
                End If
                Return colmusterProperties
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try

        End Function
        Public Function DBGetAllChildPropertiesByParentID(ByVal nval As Int64, ByVal nID As Int64) As MUSTER.Info.PropertyCollection
            Dim colmusterProperties As New MUSTER.Info.PropertyCollection
            Dim drSet As SqlDataReader
            Dim colParams As Collection
            Try
                colParams = SqlHelperParameterCache.GetSpParameterCol(_strConn, "sp_GET_PROPERTIES_AND_CHILDREN", False)
                colParams("@PROPERTY_TYPE_ID").VALUE = nID
                colParams.Remove("@PROPERTY_TYPE_NAME")
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, "sp_GET_PROPERTIES_AND_CHILDREN", colParams)
                'drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_PROPERTY_MASTER where ")
                While drSet.Read
                    If drSet.Item("PARENT_ID") = nval Then
                        Dim omusterpropertyInfo As New MUSTER.Info.PropertyInfo( _
                            drSet.Item("PROPERTY_ID"), _
                            AltIsDBNull(drSet.Item("PARENT_ID"), 0), _
                            drSet.Item("PROPERTY_TYPE_ID"), _
                            drSet.Item("PROPERTY_NAME"), _
                            AltIsDBNull(drSet.Item("PROPERTY_DESCRIPTION"), String.Empty), _
                            AltIsDBNull(drSet.Item("PROPERTY_POSITION"), 1), _
                            AltIsDBNull(drSet.Item("BUSINESS_TAG"), "0"), _
                            IIf(drSet.Item("PROPERTY_ACTIVE").ToString.ToUpper = "YES", True, False), _
                            drSet.Item("CREATED_BY"), _
                            drSet.Item("CREATE_DATE"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("1/1/1900")))
                        colmusterProperties.Add(omusterpropertyInfo)
                    End If
                End While
                Return colmusterProperties
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetPropertyByID(ByVal nVal As Int64) As MUSTER.Info.PropertyInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetSYSProperty"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Property_Name").Value = DBNull.Value
                Params("@Property_Active").Value = DBNull.Value

                Params("@OrderBy").Value = 1
                Params("@Property_ID").Value = nVal

                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_PROPERTY_MASTER WHERE PROPERTY_ID = '" & nVal & "'")

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    While drSet.Read
                        Return New MUSTER.Info.PropertyInfo( _
                            drSet.Item("PROPERTY_ID"), _
                            0, _
                            drSet.Item("PROPERTY_TYPE_ID"), _
                            drSet.Item("PROPERTY_NAME"), _
                            AltIsDBNull(drSet.Item("PROPERTY_DESCRIPTION"), String.Empty), _
                            AltIsDBNull(drSet.Item("PROPERTY_POSITION"), 1), _
                            AltIsDBNull(drSet.Item("BUSINESS_TAG"), "0"), _
                            IIf(drSet.Item("PROPERTY_ACTIVE").ToString.ToUpper = "YES", True, False), _
                            drSet.Item("CREATED_BY"), _
                            drSet.Item("CREATE_DATE"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("1/1/1900")))
                    End While
                Else
                    Return New MUSTER.Info.PropertyInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetPropertyByName(ByVal nVal As String) As MUSTER.Info.PropertyInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetSYSProperty"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Property_ID").Value = DBNull.Value
                Params("@Property_Active").Value = DBNull.Value

                Params("@OrderBy").Value = 1
                Params("@Property_Name").Value = nVal
                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_PROPERTY_MASTER WHERE PROPERTY_NAME = '" & nVal & "'")

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    While drSet.Read
                        Return New MUSTER.Info.PropertyInfo( _
                                drSet.Item("PROPERTY_ID"), _
                                0, _
                                drSet.Item("PROPERTY_TYPE_ID"), _
                                drSet.Item("PROPERTY_NAME"), _
                                AltIsDBNull(drSet.Item("PROPERTY_DESCRIPTION"), String.Empty), _
                                AltIsDBNull(drSet.Item("PROPERTY_POSITION"), 1), _
                                AltIsDBNull(drSet.Item("BUSINESS_TAG"), "0"), _
                                IIf(drSet.Item("PROPERTY_ACTIVE").ToString.ToUpper = "YES", True, False), _
                                drSet.Item("CREATED_BY"), _
                                drSet.Item("CREATE_DATE"), _
                                AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("1/1/1900")))
                    End While
                Else
                    Return New MUSTER.Info.PropertyInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetPropertyNameByID(ByVal nVal As Int16) As String
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            'strSQL = "SELECT * FROM tblSYS_PROPERTY_MASTER WHERE PROPERTY_ID ='" + strVal + "'"

            Try
                strSQL = "spGetSYSProperty"
                strVal = nVal

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Property_Name").Value = DBNull.Value
                Params("@Property_Active").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Property_ID").Value = strVal

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    While drSet.Read
                        Return drSet.Item("PROPERTY_NAME")
                    End While
                Else
                    Return "N/A"
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
        Public Sub Put(ByRef oPropInfo As MUSTER.Info.PropertyInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                Dim Params(8) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutSYSPROPERTYMASTER")
                Params(0).Value = CInt(oPropInfo.ID)
                Params(1).Value = CInt(oPropInfo.PropType_ID)
                Params(2).Value = AltIsDBNull(oPropInfo.Name, String.Empty)
                Params(3).Value = AltIsDBNull(oPropInfo.PropDesc, String.Empty)
                Params(4).Value = AltIsDBNull(oPropInfo.PropPos, 0)
                Params(5).Value = AltIsDBNull(oPropInfo.BUSINESSTAG, 0)
                Params(6).Value = Left(oPropInfo.PropIsActive.ToString, 3)
                'Params(6).Value = oPropInfo.CreatedOn
                If oPropInfo.ID <= 0 Then
                    Params(7).Value = oPropInfo.CreatedBy
                Else
                    Params(7).Value = oPropInfo.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutSYSPROPERTYMASTER", Params)

                If Params(0).Value <> oPropInfo.ID Then
                    oPropInfo.ID = Params(0).Value
                End If

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        Public Function PutProperties(ByRef dtProperties As DataTable, ByVal Property_Type_ID As Integer)
            Try
                SqlHelper.SaveDataTable(_strConn, dtProperties, getPutCommand(Property_Type_ID), getPutCommand(Property_Type_ID))
                dtProperties.AcceptChanges()
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Public Function getPutCommand(ByVal Property_Type_ID As Integer) As SqlClient.SqlCommand
            Dim oCmd As New SqlClient.SqlCommand
            Try

                oCmd.CommandType = CommandType.StoredProcedure
                oCmd.CommandText = "spPutSYSPROPERTYMASTER"

                oCmd.Parameters.Add(New SqlClient.SqlParameter("@PROPERTY_ID", SqlDbType.Int))
                oCmd.Parameters.Add(New SqlClient.SqlParameter("@PROPERTY_TYPE_ID", SqlDbType.Int))
                oCmd.Parameters.Add(New SqlClient.SqlParameter("@PROPERTY_NAME", SqlDbType.VarChar, 150))
                oCmd.Parameters.Add(New SqlClient.SqlParameter("@PROPERTY_DESCRIPTION", SqlDbType.VarChar, 200))
                oCmd.Parameters.Add(New SqlClient.SqlParameter("@PROPERTY_POSITION", SqlDbType.Int))
                oCmd.Parameters.Add(New SqlClient.SqlParameter("@BUSINESS_TAG", SqlDbType.Int))
                oCmd.Parameters.Add(New SqlClient.SqlParameter("@PROPERTY_ACTIVE", SqlDbType.Char, 3))
                oCmd.Parameters.Add(New SqlClient.SqlParameter("@CREATE_DATE", SqlDbType.DateTime))
                'oCmd.Parameters.Add(New SqlClient.SqlParameter("@NEW_PROPERTY_ID", SqlDbType.Int))
                oCmd.Parameters("@PROPERTY_ID").Direction = ParameterDirection.InputOutput

                oCmd.Parameters("@PROPERTY_ID").SourceColumn = "Property ID"
                oCmd.Parameters("@PROPERTY_TYPE_ID").Value = Property_Type_ID
                oCmd.Parameters("@PROPERTY_NAME").SourceColumn = "Property Name"
                oCmd.Parameters("@PROPERTY_DESCRIPTION").Value = ""
                oCmd.Parameters("@PROPERTY_POSITION").SourceColumn = "Property Position"
                oCmd.Parameters("@BUSINESS_TAG").Value = 1
                oCmd.Parameters("@PROPERTY_ACTIVE").SourceColumn = "Property Active"
                oCmd.Parameters("@CREATE_DATE").Value = Now().ToShortDateString


                getPutCommand = oCmd
            Catch ex As Exception
                Throw ex
            End Try
        End Function
    End Class
End Namespace
