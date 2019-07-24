
Imports Utils.DBUtils

Imports System.Data.SqlClient

Namespace MUSTER.DataAccess
    ' -------------------------------------------------------------------------------
    '   MUSTER.DataAccess.OwnerDB
    '       Provides the means for marshalling Entity state to/from the repository
    ' 
    '   Copyright (C) 2004 CIBER, Inc.
    '   All rights reserved.
    ' 
    '   Release   Initials    Date        Description
    '     1.0        JVC2      02/08/2005  Original framework from Rational XDE.
    ' 
    '   Function                  Description
    '   GetAllInfo()      Returns an EntityCollection containing all Entity objects in the repository
    '   DBGetByID(ID)     Returns an EntityInfo object indicated by int arg ID
    '   DBGetDS(SQL)      Returns a resultant Dataset by running query specified by the string arg SQL
    '   Put(Entity)       Saves the Entity passed as an argument, to the DB
    ' -------------------------------------------------------------------------------
    '
    Public Class RegistrationActivityDB
#Region "Private Member Variables"
        Private _strConn
        Private MusterException As MUSTER.Exceptions.MusterExceptions = New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
#Region "Exposed Operations"
        ' Retrieves a registration activity from the repository with a matching ID
        Public Function DBGetByID(ByVal nVal As Int64) As MUSTER.Info.RegistrationActivityInfo
            Dim drSet As SqlDataReader
            Try
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblREG_REGDTL WHERE REG_ACTION_ID = " & nVal.ToString)
                If drSet.HasRows Then
                    drSet.Read()
                    '********************************************************
                    '
                    ' Code to take items from datastream and build new object
                    '
                    '********************************************************
                    Return New MUSTER.Info.RegistrationActivityInfo( _
                                    (AltIsDBNull(drSet.Item("REG_ACTION_ID"), 0)), _
                                    (AltIsDBNull(drSet.Item("REG_ID"), 0)), _
                                    (AltIsDBNull(drSet.Item("ENTITY_TYPE"), 0)), _
                                    (AltIsDBNull(drSet.Item("ENTITY_ID"), 0)), _
                                    (AltIsDBNull(drSet.Item("USERID"), String.Empty)), _
                                    (AltIsDBNull(drSet.Item("ACTIVITY"), 0)), _
                                    (AltIsDBNull(drSet.Item("PROCESSED"), False)), _
                                    (AltIsDBNull(drSet.Item("DATE_INITIATED"), "1/1/1900")), _
                                    (AltIsDBNull(drSet.Item("CALENDAR_INFO_ID"), 0)))
                    If Not drSet.IsClosed Then drSet.Close()
                Else
                    If Not drSet.IsClosed Then drSet.Close()
                    'Return New MUSTER.Info.RegistrationActivityInfo
                End If
            Catch ex As Exception
                If Not drSet.IsClosed Then drSet.Close()
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function
        Public Function DBGetByEntity(ByVal EntityID As Integer, ByVal EntityType As Integer, ByVal Activity As String) As MUSTER.Info.RegistrationActivityInfo
            Dim drSet As SqlDataReader
            Try
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblREG_REGDTL WHERE ENTITY_ID = " & EntityID.ToString & " AND ENTITY_TYPE = " & EntityType.ToString & " AND ACTIVITY = " & Activity.ToString & " AND PROCESSED = 0 AND DELETED = 0")
                If drSet.HasRows Then
                    drSet.Read()
                    '********************************************************
                    '
                    ' Code to take items from datastream and build new object
                    '
                    '********************************************************
                    Return New MUSTER.Info.RegistrationActivityInfo( _
                                    (AltIsDBNull(drSet.Item("REG_ACTION_ID"), 0)), _
                                    (AltIsDBNull(drSet.Item("REG_ID"), 0)), _
                                    (AltIsDBNull(drSet.Item("ENTITY_TYPE"), 0)), _
                                    (AltIsDBNull(drSet.Item("ENTITY_ID"), 0)), _
                                    (AltIsDBNull(drSet.Item("USERID"), String.Empty)), _
                                    (AltIsDBNull(drSet.Item("ACTIVITY"), 0)), _
                                    (AltIsDBNull(drSet.Item("PROCESSED"), False)), _
                                    (AltIsDBNull(drSet.Item("DATE_INITIATED"), "1/1/1900")), _
                                    (AltIsDBNull(drSet.Item("CALENDAR_INFO_ID"), 0)))
                    If Not drSet.IsClosed Then drSet.Close()
                Else
                    If Not drSet.IsClosed Then drSet.Close()
                    Return New MUSTER.Info.RegistrationActivityInfo
                End If
            Catch ex As Exception
                If Not drSet.IsClosed Then drSet.Close()
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function
        ' Retrieves all registration activities for the supplied Registration ID
        Public Function DBGetByRegistration(ByVal regID As Integer) As MUSTER.Info.RegistrationActivityCollection

            Dim localRegActivityCol As New MUSTER.Info.RegistrationActivityCollection
            Dim localRegistrationActivityInfo As MUSTER.Info.RegistrationActivityInfo
            Dim drSet As SqlDataReader
            Try
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblREG_REGDTL WHERE REG_ID = " & regID.ToString & " AND PROCESSED = 0 AND DELETED = 0")
                While drSet.Read
                    '********************************************************
                    '
                    ' Code to take items from datastream and build new object
                    '
                    '********************************************************
                    localRegistrationActivityInfo = New MUSTER.Info.RegistrationActivityInfo( _
                                    (AltIsDBNull(drSet.Item("REG_ACTION_ID"), 0)), _
                                    (AltIsDBNull(drSet.Item("REG_ID"), 0)), _
                                    (AltIsDBNull(drSet.Item("ENTITY_TYPE"), 0)), _
                                    (AltIsDBNull(drSet.Item("ENTITY_ID"), 0)), _
                                    (AltIsDBNull(drSet.Item("USERID"), String.Empty)), _
                                    (AltIsDBNull(drSet.Item("ACTIVITY"), String.Empty)), _
                                    (AltIsDBNull(drSet.Item("PROCESSED"), False)), _
                                    (AltIsDBNull(drSet.Item("DATE_INITIATED"), "1/1/1900")), _
                                    (AltIsDBNull(drSet.Item("CALENDAR_INFO_ID"), 0)))
                    localRegActivityCol.Add(localRegistrationActivityInfo)
                End While
                Return localRegActivityCol
            Catch ex As Exception
                'If Not drSet.IsClosed Then drSet.Close()
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
        ' Retrieves all active registrations
        Public Function GetAllInfo() As MUSTER.Info.RegistrationActivityCollection
        End Function
        Public Sub Put(ByRef RegistrationActivityInfo As MUSTER.Info.RegistrationActivityInfo)
            Try
                Dim Params(9) As SqlParameter
                '********************************************************
                '
                ' Change second argument of GetSpParameterSet to name of SPROC
                '
                '********************************************************
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutRegDetail")
                '********************************************************
                '
                ' Add Params assignments as necessary
                '
                '********************************************************


                If RegistrationActivityInfo.RegActionIndex > 0 Then
                    Params(0).Value = RegistrationActivityInfo.RegActionIndex
                Else
                    Params(0).Value = RegistrationActivityInfo.RegActionIndex
                End If

                Params(1).Value = RegistrationActivityInfo.RegistrationID
                Params(2).Value = RegistrationActivityInfo.EntityType
                Params(3).Value = RegistrationActivityInfo.EntityId
                Params(4).Value = RegistrationActivityInfo.ActivityDesc
                Params(5).Value = RegistrationActivityInfo.Processed
                Params(6).Value = RegistrationActivityInfo.Deleted
                Params(7).Value = RegistrationActivityInfo.UserID
                Params(8).Value = RegistrationActivityInfo.DateAdded
                Params(9).Value = RegistrationActivityInfo.CalendarID
                '********************************************************
                '
                ' Change second argument of GetSpParameterSet to name of SPROC
                '
                '********************************************************
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutRegDetail", Params)
                If Params(0).Value <> RegistrationActivityInfo.RegActionIndex Then
                    RegistrationActivityInfo.RegActionIndex = Params(0).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
#End Region
    End Class
End Namespace
