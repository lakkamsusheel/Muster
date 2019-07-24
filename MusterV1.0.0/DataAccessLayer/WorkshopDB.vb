'-------------------------------------------------------------------------------
' MUSTER.DataAccess.WorkshopDB
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
' NOTE: This file to be used as Workshop to build other objects.
'       Replace keyword "Workshop" with respective Object name.
'       In addition, change SQL strings and data stream Items as necessary
'       Don't forget to update the history information above!!!
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class WorkshopDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function GetAllInfo() As Muster.Info.WorkshopsCollection
            Try
                '********************************************************
                '
                ' Alter SQL String as necessary
                '
                '********************************************************
                Dim drSet As SqlDataReader = SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SQL TO SELECT Workshop DATA HERE")
                Dim colEntities As New Muster.Info.WorkshopsCollection
                While drSet.Read
                    '********************************************************
                    '
                    ' Code to take items from datastream and build new object
                    '
                    '********************************************************
                    Dim oWorkshopInfo As New Muster.Info.WorkshopInfo
                    '********************************************************
                    '
                    ' Other private member variables for current state here
                    '
                    '********************************************************
                    colEntities.Add(oWorkshopInfo)
                End While
                If Not drSet.IsClosed Then drSet.Close()
                Return colEntities
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        Public Function DBGetByName(ByVal strVal As String) As Muster.Info.WorkshopInfo
            Dim drSet As SqlDataReader
            Try
                '********************************************************
                '
                ' Alter SQL String as necessary
                '
                '********************************************************
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SQL TO SELECT Workshop DATA HERE = '" & strVal & "'")
                If drSet.HasRows Then
                    drSet.Read()
                    '********************************************************
                    '
                    ' Code to take items from datastream and build new object
                    '
                    '********************************************************
                    Return New Muster.Info.WorkshopInfo
                    If Not drSet.IsClosed Then drSet.Close()
                Else
                    '********************************************************
                    '
                    ' Dim an empty object and set the name (identifying attribute)
                    '
                    '********************************************************
                    Dim oWorkshopInfo As New Muster.Info.WorkshopInfo
                    '
                    ' Had to comment out so class would compile.  Normally the next
                    '  line is where the identifying attribute is assigned the value
                    '  passed from the client.
                    '
                    'oWorkshopInfo.Name = strVal
                    If Not drSet.IsClosed Then drSet.Close()
                    Return oWorkshopInfo
                End If
            Catch ex As Exception
                If Not drSet.IsClosed Then drSet.Close()
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function
        Public Function DBGetByID(ByVal nVal As Int64) As Muster.Info.WorkshopInfo
            Dim drSet As SqlDataReader
            Try
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_REPORT_MASTER WHERE REPORT_ID = " & nVal.ToString)
                If drSet.HasRows Then
                    drSet.Read()
                    '********************************************************
                    '
                    ' Code to take items from datastream and build new object
                    '
                    '********************************************************
                    Return New Muster.Info.WorkshopInfo
                    If Not drSet.IsClosed Then drSet.Close()
                Else
                    If Not drSet.IsClosed Then drSet.Close()
                    Return New Muster.Info.WorkshopInfo
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
        Public Sub Put(ByRef oWorkshopInfo As Muster.Info.WorkshopInfo)
            Try
                Dim Params(5) As SqlParameter
                '********************************************************
                '
                ' Change second argument of GetSpParameterSet to name of SPROC
                '
                '********************************************************
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "SPROC NAME HERE")
                '********************************************************
                '
                ' Add Params assignments as necessary
                '
                '********************************************************
                Params(0).Value = oWorkshopInfo.ID
                '********************************************************
                '
                ' Change second argument of GetSpParameterSet to name of SPROC
                '
                '********************************************************
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "SPROC NAME HERE", Params)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
    End Class
End Namespace
