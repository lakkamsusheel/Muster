'-------------------------------------------------------------------------------
' MUSTER.DataAccess.RegistrationDB
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
' NOTE: This file to be used as Registration to build other objects.
'       Replace keyword "Registration" with respective Object name.
'       In addition, change SQL strings and data stream Items as necessary
'       Don't forget to update the history information above!!!
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class RegistrationDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function GetAllInfo() As Muster.Info.RegistrationsCollection
            Try
                '********************************************************
                '
                ' Alter SQL String as necessary
                '
                '********************************************************
                Dim drSet As SqlDataReader = SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * from tblREG_REGHDR Where COMPLETED <> 1")
                Dim colEntities As New Muster.Info.RegistrationsCollection
                While drSet.Read
                    '********************************************************
                    '
                    ' Code to take items from datastream and build new object
                    '
                    '********************************************************
                    Dim oRegistrationInfo As New MUSTER.Info.RegistrationInfo( _
                                    (AltIsDBNull(drSet.Item("REG_ID"), 0)), _
                                    (AltIsDBNull(drSet.Item("OWNER_ID"), 0)), _
                                    (AltIsDBNull(drSet.Item("DATE_STARTED"), "1/1/1900")), _
                                    (AltIsDBNull(drSet.Item("DATE_COMPLETED"), "1/1/1900")), _
                                    (AltIsDBNull(drSet.Item("COMPLETED"), 0)), _
                                    (AltIsDBNull(drSet.Item("DELETED"), 0)), "", Now(), "", Now())

                    '********************************************************
                    '
                    ' Other private member variables for current state here
                    '
                    '********************************************************
                    colEntities.Add(oRegistrationInfo)
                End While
                If Not drSet.IsClosed Then drSet.Close()
                Return colEntities
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        Public Function DBGetByOwnerID(ByVal ownerID As Int64) As MUSTER.Info.RegistrationInfo
            If ownerID <= 0 Then
                Return New MUSTER.Info.RegistrationInfo
            End If
            Dim drSet As SqlDataReader
            Try
                '********************************************************
                '
                ' Alter SQL String as necessary
                '
                '********************************************************
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * from tblREG_REGHDR Where COMPLETED <> 1 AND OWNER_ID = " + ownerID.ToString)
                If drSet.HasRows Then
                    drSet.Read()
                    '********************************************************
                    '
                    ' Code to take items from datastream and build new object
                    '
                    '********************************************************
                    Return New MUSTER.Info.RegistrationInfo( _
                                    (AltIsDBNull(drSet.Item("REG_ID"), 0)), _
                                    (AltIsDBNull(drSet.Item("OWNER_ID"), 0)), _
                                    (AltIsDBNull(drSet.Item("DATE_STARTED"), "1/1/1900")), _
                                    (AltIsDBNull(drSet.Item("DATE_COMPLETED"), "1/1/1900")), _
                                    (AltIsDBNull(drSet.Item("COMPLETED"), 0)), _
                                    (AltIsDBNull(drSet.Item("DELETED"), 0)), "", Now(), "", Now())

                    If Not drSet.IsClosed Then drSet.Close()
                Else
                    Return New MUSTER.Info.RegistrationInfo
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBGetByID(ByVal nVal As Int64) As MUSTER.Info.RegistrationInfo
            Dim drSet As SqlDataReader
            Try
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblREG_REGHDR WHERE REG_ID = " & nVal.ToString)
                If drSet.HasRows Then
                    drSet.Read()
                    '********************************************************
                    '
                    ' Code to take items from datastream and build new object
                    '
                    '********************************************************
                    Return New MUSTER.Info.RegistrationInfo( _
                                    (AltIsDBNull(drSet.Item("REG_ID"), 0)), _
                                    (AltIsDBNull(drSet.Item("OWNER_ID"), 0)), _
                                    (AltIsDBNull(drSet.Item("DATE_STARTED"), "1/1/1900")), _
                                    (AltIsDBNull(drSet.Item("DATE_COMPLETED"), "1/1/1900")), _
                                    (AltIsDBNull(drSet.Item("COMPLETED"), 0)), _
                                    (AltIsDBNull(drSet.Item("DELETED"), 0)), "", Now(), "", Now())

                    If Not drSet.IsClosed Then drSet.Close()
                Else
                    If Not drSet.IsClosed Then drSet.Close()
                    Return New MUSTER.Info.RegistrationInfo
                End If
            Catch ex As Exception
                If Not drSet.IsClosed Then drSet.Close()
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
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
        Public Sub Put(ByRef oRegistrationInfo As MUSTER.Info.RegistrationInfo)
            Try
                Dim Params(6) As SqlParameter
                '********************************************************
                '
                ' Change second argument of GetSpParameterSet to name of SPROC
                '
                '********************************************************
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutRegHeader")
                '********************************************************
                '
                ' Add Params assignments as necessary
                '
                '********************************************************

                If oRegistrationInfo.ID > 0 Then
                    Params(0).Value = oRegistrationInfo.ID
                Else
                    Params(0).Value = 0
                End If
                Params(1).Value = oRegistrationInfo.OWNER_ID
                Params(2).Value = oRegistrationInfo.DATE_STARTED
                Params(3).Value = IIf(oRegistrationInfo.DATE_COMPLETED = CDate("#12:00:00 AM#"), DBNull.Value, oRegistrationInfo.DATE_COMPLETED)
                Params(4).Value = oRegistrationInfo.COMPLETED
                Params(5).Value = oRegistrationInfo.Deleted
                '********************************************************
                '
                ' Change second argument of GetSpParameterSet to name of SPROC
                '
                '********************************************************
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutRegHeader", Params)
                If Params(0).Value <> oRegistrationInfo.ID Then
                    oRegistrationInfo.ID = Params(0).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
    End Class
End Namespace
