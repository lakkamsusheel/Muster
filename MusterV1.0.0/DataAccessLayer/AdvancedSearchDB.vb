'-------------------------------------------------------------------------------
' MUSTER.DataAccess.AdvancedSearchDB
'   Provides the means for marshalling search results to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       AN      01/03/05    Original class definition.(Added Comment Date is Invalid no header present)
'  1.0       EN      02/01/05    Added GetResults Method... 
'
'
' Function            Description
'   DBGetDS(strSql)         Returns a data set based on the query provided.     
'-------------------------------------------------------------------------------

Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class AdvancedSearchDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
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
        Public Function GetResults(ByVal SearchType As String, ByVal strCriteria As String, ByVal strTankstatus As String, ByVal nLustStatus As Integer) As DataSet
            Try
                Dim dsSet As DataSet
                Dim Params(3) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spSELAdvance_Search")
                Params(0).Value = SearchType
                Params(1).Value = strCriteria
                Params(2).Value = strTankstatus
                Params(3).Value = nLustStatus
                dsSet = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, "spSELAdvance_Search", Params)
                Return dsSet
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetEntity(ByVal strSQL As String) As DataSet
            Dim dsData As DataSet
            Try
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
    End Class
End Namespace
