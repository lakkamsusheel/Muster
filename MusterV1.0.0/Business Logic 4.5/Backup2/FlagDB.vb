'-------------------------------------------------------------------------------
' MUSTER.DataAccess.FlagDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MNR     12/14/04    Original class definition.
'  1.1        AN      12/30/04    Added Try catch and Exception Handling/Logging
'  1.2        MR      01/31/05    Modified PUT function to add additional Parameters.
'  1.3        JVC2    01/31/05    Changed Parm list for PUT
'  1.4        JVC2    01/31/2005  Added code to accomodate new column SOURCE_USER_ID
'  1.5        JVC2    02/03/2005  Added where clause (DELETED = 0) in GetAllInfo()
'  1.6        EN      02/10/2005   Modified 01/01/1901  to 01/01/0001 
'                                   Also modified GetAllInfo and DBGetByID to match NEW signature.
'  1.7        AB      02/15/05    Replaced dynamic SQL with stored procedures in the following
'                                   Functions:  GetAllInfo, DBGetByID
'  1.8        AB      02/16/05    Added Finally to the Try/Catch to close all datareaders
'  1.9        AB      02/16/05    Removed any IsNull calls for fields the DB requires
'  1.10       AB      02/18/05    Set all parameters for SP, that are not required, to NULL
'  1.11       AB      02/23/05    Modified Get functions based upon changes made to 
'                                   make several nullable fields non-nullable
'
' Function                  Description
' GetAllInfo()      Returns an EntityCollection containing all Entity objects in the repository
' DBGetByID(ID)     Returns an EntityInfo object indicated by int arg ID
' DBGetDS(SQL)      Returns a resultant Dataset by running query specified by the string arg SQL
' Put(Entity)       Saves the Entity passed as an argument, to the DB
'-------------------------------------------------------------------------------

Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class FlagDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function GetAllInfo() As MUSTER.Info.FlagsCollection
            Try
                Return DBGetFlags()
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function DBGetByID(ByVal nVal As Integer) As MUSTER.Info.FlagInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            Try
                Dim flagsCol As MUSTER.Info.FlagsCollection
                flagsCol = DBGetFlags(, , , , nVal)
                If flagsCol.Count > 0 Then
                    Return flagsCol.Item(flagsCol.GetKeys(0))
                Else
                    Return New MUSTER.Info.FlagInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetFlags(Optional ByVal entityID As Integer = 0, Optional ByVal entityType As Integer = 0, Optional ByVal showDeleted As Boolean = False, Optional ByVal [Module] As String = "", Optional ByVal flagID As Integer = 0, Optional ByVal calID As Integer = 0, Optional ByVal userID As String = "", Optional ByVal flagDesc As String = "") As MUSTER.Info.FlagsCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetFlags"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@ENTITY_ID").Value = IIf(entityID = 0, DBNull.Value, entityID)
                Params("@ENTITY_TYPE").Value = IIf(entityType = 0, DBNull.Value, entityType)
                Params("@DELETED").Value = showDeleted
                Params("@MODULE").Value = IIf([Module] = "", DBNull.Value, [Module])
                Params("@FLAG_ID").Value = IIf(flagID = 0, DBNull.Value, flagID)
                Params("@CALENDAR_INFO_ID").Value = IIf(calID = 0, DBNull.Value, calID)
                Params("@SOURCE_USER_ID").Value = IIf(userID = "", DBNull.Value, userID)
                Params("@FLAG_DESCRIPTION").Value = IIf(flagDesc = "", DBNull.Value, flagDesc)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colEntities As New MUSTER.Info.FlagsCollection
                While drSet.Read
                    Dim oFlagInfo As New MUSTER.Info.FlagInfo(drSet.Item("FLAG_ID"), _
                                                                drSet.Item("ENTITY ID"), _
                                                                drSet.Item("ENTITY_TYPE"), _
                                                                drSet.Item("DESCRIPTION"), _
                                                                drSet.Item("DELETED"), _
                                                                AltIsDBNull(drSet.Item("DUE DATE"), CDate("01/01/0001")), _
                                                                drSet.Item("MODULE"), _
                                                                AltIsDBNull(drSet.Item("CALENDAR_INFO_ID"), 0), _
                                                                drSet.Item("CREATED_BY"), _
                                                                drSet.Item("CREATED ON"), _
                                                                AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("TURNS RED ON"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("USER ID"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("FLAG_COLOR"), String.Empty))
                    colEntities.Add(oFlagInfo)
                End While
                Return colEntities
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
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
        Public Function DBGetFlagsDS(Optional ByVal entityID As Integer = 0, Optional ByVal entityType As Integer = 0, Optional ByVal showDeleted As Boolean = False, Optional ByVal [Module] As String = "", Optional ByVal flagID As Integer = 0, Optional ByVal calID As Integer = 0, Optional ByVal userID As String = "", Optional ByVal flagDesc As String = "") As DataSet
            Dim dsData As DataSet
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Try
                strSQL = "spGetFlags"
                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(entityID = 0, DBNull.Value, entityID)
                Params(1).Value = IIf(entityType = 0, DBNull.Value, entityType)
                Params(2).Value = showDeleted
                Params(3).Value = IIf([Module] = "", DBNull.Value, [Module])
                Params(4).Value = IIf(flagID = 0, DBNull.Value, flagID)
                Params(5).Value = IIf(calID = 0, DBNull.Value, calID)
                Params(6).Value = IIf(userID = "", DBNull.Value, userID)
                Params(7).Value = IIf(flagDesc = "", DBNull.Value, flagDesc)
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub Put(ByRef oFlagInf As MUSTER.Info.FlagInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                If Not moduleID = 0 Then
                    If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Flag, Integer))) Then
                        returnVal = "You do not have rights to save a Flag."
                        Exit Sub
                    Else
                        returnVal = String.Empty
                    End If
                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutSYSFLAGS")

                If oFlagInf.ID <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = oFlagInf.ID
                End If
                Params(1).Value = oFlagInf.EntityID
                Params(2).Value = oFlagInf.EntityType
                Params(3).Value = oFlagInf.FlagDescription
                Params(4).Value = oFlagInf.Deleted

                If Date.Compare(oFlagInf.DueDate, CDate("01/01/0001")) = 0 Then
                    Params(5).Value = DBNull.Value
                Else
                    Params(5).Value = oFlagInf.DueDate
                End If
                Params(6).Value = oFlagInf.ModuleID
                Params(7).Value = oFlagInf.CalendarInfoID
                Params(8).Value = oFlagInf.SourceUserID
                Params(9).Value = oFlagInf.FlagColor
                If Date.Compare(oFlagInf.TurnsRedOn, CDate("01/01/0001")) = 0 Then
                    Params(10).Value = DBNull.Value
                Else
                    Params(10).Value = oFlagInf.TurnsRedOn
                End If

                If oFlagInf.ID <= 0 Then
                    Params(11).Value = oFlagInf.CreatedBy
                Else
                    Params(11).Value = oFlagInf.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutSYSFLAGS", Params)
                If Params(0).Value <> oFlagInf.ID Then
                    oFlagInf.ID = Params(0).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function GetBarometerColors(Optional ByVal entityID As Integer = 0, Optional ByVal entityType As Integer = 0, Optional ByVal eventID As Integer = 0, Optional ByVal eventType As Integer = 0) As DataSet
            Dim dsData As DataSet
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Try
                strSQL = "spGetBarometerColors"
                Dim Params(3) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(entityID = 0, DBNull.Value, entityID)
                Params(1).Value = IIf(entityType = 0, DBNull.Value, entityType)
                Params(2).Value = IIf(eventID = 0, DBNull.Value, eventID)
                Params(3).Value = IIf(eventType = 0, DBNull.Value, eventType)
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
    End Class
End Namespace
