'-------------------------------------------------------------------------------
' MUSTER.DataAccess.UserDB
'   Provides the means for marshalling User state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        PN      12/14/04    Original class definition.
'  1.1        AN      12/30/04    Added Try catch and Exception Handling/Logging
'  1.2        EN      02/10/05    Modifie 01/01/1901  to 01/01/0001 
'  1.3        AB      02/14/05    Replaced dynamic SQL with stored procedures in the following
'                                 Functions:  DBGetAllCommentsInfo, DBGetByModuleName, DBGetByID
'  1.4        AB      02/15/05    Added Finally to the Try/Catch to close all datareaders
'  1.5        AB      02/16/05    Removed any IsNull calls for fields the DB requires
'  1.6        AB      02/18/05    Set all parameters for SP, that are not required, to NULL
'  1.7        AB      02/23/05    Modified Get and Put functions based upon changes made to 
'                                   make several nullable fields non-nullable
'
'
' Function                  Description
' GetAllInfo()        Returns an UserCollection containing all User objects in the repository.
' DBGetByName(NAME)   Returns an UserInfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an UserInfo object indicated by arg ID.
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils
Namespace MUSTER.DataAccess
    Public Class CommentsDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
        Public Function DBGetAllCommentsInfo() As Muster.Info.CommentsCollection
            Dim colComments As New MUSTER.Info.CommentsCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetComments"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)

                Params("@MODULE_ID").Value = DBNull.Value
                Params("@ENTITY_TYPE").Value = DBNull.Value
                Params("@ENTITY_ID").Value = DBNull.Value
                Params("@ENTITY_ADDITIONAL_INFO").Value = DBNull.Value
                Params("@USER_ID").Value = DBNull.Value
                Params("@COMMENT_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@DELETED").Value = False

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                While drSet.Read
                    Dim oCommentsInfo As New MUSTER.Info.CommentsInfo(drSet.Item("COMMENT_ID"), _
                                                              drSet.Item("ENTITY ID"), _
                                                              AltIsDBNull(drSet.Item("ENTITY_ADDITIONAL_INFO"), String.Empty), _
                                                              drSet.Item("ENTITY_TYPE"), _
                                                              drSet.Item("COMMENT"), _
                                                              drSet.Item("VIEWABLE BY"), _
                                                              drSet.Item("DELETED"), _
                                                              drSet.Item("USER ID"), _
                                                              drSet.Item("COMMENT_DATE"), _
                                                              drSet.Item("MODULE"), _
                                                              drSet.Item("CREATEDBY"), _
                                                              drSet.Item("CREATED ON"), _
                                                              AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                              AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                    colComments.Add(oCommentsInfo)
                End While

                Return colComments

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        'Public Function DBGetByEntityType(ByVal nEntityType As Integer, ByVal nEntityID As Integer, Optional ByVal entityAddnInfo As String = "") As MUSTER.Info.CommentsCollection
        '    Dim colComments As New MUSTER.Info.CommentsCollection
        '    Dim drSet As SqlDataReader
        '    Dim Params As Collection
        '    Dim strSQL As String
        '    Try
        '        strSQL = "spGetComments"

        '        Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
        '        Params("@MODULE_ID").Value = DBNull.Value
        '        Params("@ENTITY_TYPE").Value = IIFIsIntegerNull(nEntityType, System.DBNull.Value)
        '        Params("@ENTITY_ID").Value = IIFIsIntegerNull(nEntityID, System.DBNull.Value)
        '        Params("@ENTITY_ADDITIONAL_INFO").Value = IIf(entityAddnInfo = String.Empty, DBNull.Value, entityAddnInfo)
        '        Params("@USER_ID").Value = DBNull.Value
        '        Params("@COMMENT_ID").Value = DBNull.Value
        '        Params("@OrderBy").Value = 1
        '        Params("@DELETED").Value = False

        '        drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
        '        If drSet.HasRows Then
        '            While drSet.Read
        '                Dim oCommentsInfo As New MUSTER.Info.CommentsInfo(drSet.Item("COMMENT_ID"), _
        '                                                      drSet.Item("ENTITY ID"), _
        '                                                      AltIsDBNull(drSet.Item("ENTITY_ADDITIONAL_INFO"), String.Empty), _
        '                                                      drSet.Item("ENTITY TYPE"), _
        '                                                      drSet.Item("COMMENT"), _
        '                                                      drSet.Item("SCOPE"), _
        '                                                      drSet.Item("DELETED"), _
        '                                                      drSet.Item("USER ID"), _
        '                                                      drSet.Item("COMMENT_DATE"), _
        '                                                      drSet.Item("MODULE"), _
        '                                                      drSet.Item("CREATED_BY"), _
        '                                                      drSet.Item("CREATED ON"), _
        '                                                      AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
        '                                                      AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
        '                colComments.Add(oCommentsInfo)
        '            End While
        '        End If

        '        Return colComments

        '    Catch Ex As Exception
        '        MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    Finally
        '        If Not drSet.IsClosed Then drSet.Close()
        '    End Try
        'End Function
        'Public Function DBGetByModuleName(ByVal strModuleName As String, ByVal nEntityType As Integer, ByVal nEntityID As Integer, Optional ByVal entityAddnInfo As String = "") As MUSTER.Info.CommentsCollection
        '    Dim colComments As New MUSTER.Info.CommentsCollection
        '    Dim drSet As SqlDataReader
        '    Dim Params As Collection
        '    Dim strSQL As String
        '    Try
        '        strSQL = "spGetComments"

        '        Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
        '        Params("@MODULE_ID").Value = strModuleName
        '        Params("@ENTITY_TYPE").Value = IIFIsIntegerNull(nEntityType, System.DBNull.Value)
        '        Params("@ENTITY_ID").Value = IIFIsIntegerNull(nEntityID, System.DBNull.Value)
        '        Params("@ENTITY_ADDITIONAL_INFO").Value = IIf(entityAddnInfo = String.Empty, DBNull.Value, entityAddnInfo)
        '        Params("@USER_ID").Value = DBNull.Value
        '        Params("@COMMENT_ID").Value = DBNull.Value
        '        Params("@OrderBy").Value = 1
        '        Params("@DELETED").Value = False

        '        drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
        '        If drSet.HasRows Then
        '            While drSet.Read
        '                Dim oCommentsInfo As New MUSTER.Info.CommentsInfo(drSet.Item("COMMENT_ID"), _
        '                                                      drSet.Item("ENTITY ID"), _
        '                                                      AltIsDBNull(drSet.Item("ENTITY_ADDITIONAL_INFO"), String.Empty), _
        '                                                      drSet.Item("ENTITY TYPE"), _
        '                                                      drSet.Item("COMMENT"), _
        '                                                      drSet.Item("SCOPE"), _
        '                                                      drSet.Item("DELETED"), _
        '                                                      drSet.Item("USER ID"), _
        '                                                      drSet.Item("COMMENT_DATE"), _
        '                                                      drSet.Item("MODULE"), _
        '                                                      drSet.Item("CREATED_BY"), _
        '                                                      drSet.Item("CREATED ON"), _
        '                                                      AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
        '                                                      AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
        '                colComments.Add(oCommentsInfo)
        '            End While
        '        End If

        '        Return colComments

        '    Catch Ex As Exception
        '        MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    Finally
        '        If Not drSet.IsClosed Then drSet.Close()
        '    End Try
        'End Function
        Public Function DBGetByID(ByVal nCommentID As Integer, Optional ByVal strUserID As String = "") As MUSTER.Info.CommentsInfo
            Dim drSet As SqlDataReader
            Dim Params As Collection
            Dim strSQL As String
            Try
                strSQL = "spGetComments"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@MODULE_ID").Value = DBNull.Value
                Params("@ENTITY_TYPE").Value = DBNull.Value
                Params("@ENTITY_ID").Value = DBNull.Value
                Params("@ENTITY_ADDITIONAL_INFO").Value = DBNull.Value
                Params("@USER_ID").Value = IIf(strUserID = "", DBNull.Value, strUserID)
                Params("@COMMENT_ID").Value = nCommentID
                Params("@OrderBy").Value = 1
                Params("@DELETED").Value = False

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.CommentsInfo(drSet.Item("COMMENT_ID"), _
                                                              drSet.Item("ENTITY ID"), _
                                                              AltIsDBNull(drSet.Item("ENTITY_ADDITIONAL_INFO"), String.Empty), _
                                                              drSet.Item("ENTITY_TYPE"), _
                                                              drSet.Item("COMMENT"), _
                                                              drSet.Item("VIEWABLE BY"), _
                                                              drSet.Item("DELETED"), _
                                                              drSet.Item("USER ID"), _
                                                              drSet.Item("COMMENT_DATE"), _
                                                              drSet.Item("MODULE"), _
                                                              drSet.Item("CREATEDBY"), _
                                                              drSet.Item("CREATED ON"), _
                                                              AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                              AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                Else
                    Return New MUSTER.Info.CommentsInfo
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function put(ByVal oCommentsInfo As MUSTER.Info.CommentsInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String) As Integer

            Dim Params(8) As SqlParameter
            Dim strSQL As String
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Comment, Integer))) Then
                    returnVal = "You do not have rights to save a Comment."
                    Exit Function
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutSYSCOMMENT"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIFIsIntegerNull(oCommentsInfo.ID, System.DBNull.Value)
                Params(1).Value = IIFIsIntegerNull(oCommentsInfo.EntityID, System.DBNull.Value)
                Params(2).Value = IIf(oCommentsInfo.EntityAdditionalInfo = String.Empty, DBNull.Value, oCommentsInfo.EntityAdditionalInfo)
                Params(3).Value = IIFIsIntegerNull(oCommentsInfo.EntityType, System.DBNull.Value)
                Params(4).Value = oCommentsInfo.Comments
                Params(5).Value = oCommentsInfo.CommentsScope
                Params(6).Value = oCommentsInfo.Deleted
                Params(7).Value = oCommentsInfo.UserID
                Params(8).Value = oCommentsInfo.CommentDate
                Params(9).Value = oCommentsInfo.ModuleName
                Params(0).Direction = ParameterDirection.InputOutput

                If oCommentsInfo.ID <= 0 Then
                    Params(10).Value = oCommentsInfo.CreatedBy
                Else
                    Params(10).Value = oCommentsInfo.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If oCommentsInfo.ID <= 0 Then
                    oCommentsInfo.ID = Params(0).Value
                End If
                oCommentsInfo.Archive()
                Return Params(0).Value
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
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
        Public Function DBGetComments(Optional ByVal strModuleName As String = "", Optional ByVal entityType As Integer = 0, Optional ByVal entityID As Integer = 0, Optional ByVal entityAddnInfo As String = "", Optional ByVal userID As String = "", Optional ByVal commentID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim dsData As DataSet
            Dim Params As Collection
            Dim strSQL As String
            Try
                strSQL = "spGetComments"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@MODULE_ID").Value = IIf(strModuleName = "", DBNull.Value, strModuleName)
                Params("@ENTITY_TYPE").Value = IIf(entityType = 0, DBNull.Value, entityType)
                Params("@ENTITY_ID").Value = IIf(entityID = 0, DBNull.Value, entityID)
                Params("@ENTITY_ADDITIONAL_INFO").Value = IIf(entityAddnInfo = "", DBNull.Value, entityAddnInfo)
                Params("@USER_ID").Value = IIf(userID = "", DBNull.Value, userID)
                Params("@COMMENT_ID").Value = IIf(commentID = 0, DBNull.Value, commentID)
                Params("@OrderBy").Value = 1
                Params("@DELETED").Value = showDeleted

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
    End Class
End Namespace

