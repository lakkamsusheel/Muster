'-------------------------------------------------------------------------------
' MUSTER.DataAccess.ZipCodeDB
'   Provides the means for marshalling Address to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       EN      12/21/04    Original class definition.
'  1.2       AN      12/30/04    Added Try catch and Exception Handling/Logging
'  1.3       EN      12/10/05     Modified 01/01/1901  to 01/01/0001 
'  1.4       AB      02/15/05    Replaced dynamic SQL with stored procedures in the following
'                                    Functions:  GetAllInfo, DBGetByKey
'  1.5       AB      02/17/05    Added Finally to the Try/Catch to close all datareaders
'  1.6       AB      02/18/05    Set all parameters for SP, that are not required, to NULL
'
' Function            Description
' GetAllInfo()        Returns an ZipCodeCollection containing all letter objects in the repository.
' DBCheckZip          Returns True if Zip is existing based on state,city,zip and county....
' DBGetByKey(strArray)Returns an ZipCodeInfo object indicated by Array 
' DBGetDS(SQL)        Returns a resultant Dataset by running query specified by the string arg SQL
' AddZip(oZipCode)    Saves the ZipCode passed as an argument, to the DB
''-------------------------------------------------------------------------------
'
' TODO - Add to app 1/3/2005 - JVC 2
'


Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class ZipCodeDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
#Region "Exposed Operations"
        Public Function GetAllInfo() As Muster.Info.ZipCodeCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            'strSQL = "SELECT * FROM tblSYS_ZIPCODES"

            Try
                strSQL = "spGetZipCode"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Zip").Value = DBNull.Value
                Params("@State").Value = DBNull.Value
                Params("@City").Value = DBNull.Value
                Params("@County").Value = DBNull.Value
                Params("@OrderBy").Value = 1


                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colZipCode As New MUSTER.Info.ZipCodeCollection
                Dim strID As String
                While drSet.Read
                    Dim oZipCodeInfo As New MUSTER.Info.ZipCodeInfo(AltIsDBNull(drSet.Item("Zip"), String.Empty), _
                                          AltIsDBNull(drSet.Item("state"), String.Empty), _
                                          AltIsDBNull(drSet.Item("City"), String.Empty), _
                                          AltIsDBNull(drSet.Item("County"), String.Empty), _
                                          AltIsDBNull(drSet.Item("Fips"), String.Empty), _
                                          AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                          AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                          AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                    colZipCode.Add(oZipCodeInfo)
                End While

                Return colZipCode
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBCheckZip(ByRef ozipCode As MUSTER.Info.ZipCodeInfo) As Boolean
            Dim Params(3) As SqlParameter
            Try
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spCheckZip")
                Params(0).Value = ozipCode.City
                Params(1).Value = ozipCode.state
                Params(2).Value = ozipCode.Zip
                Params(3).Value = ozipCode.County
                Dim orderCount As Integer = CInt(SqlHelper.ExecuteScalar(_strConn, CommandType.StoredProcedure, "spCheckZip", Params))
                If orderCount > 0 Then
                    Return True
                Else
                    Return False
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function AddZip(ByRef oZipCode As MUSTER.Info.ZipCodeInfo) As Boolean
            ' Only Add, NO update...
            Dim Params(4) As SqlParameter
            Try
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutZipCode")
                Params(0).Value = oZipCode.City
                Params(1).Value = oZipCode.state
                Params(2).Value = oZipCode.Zip
                Params(3).Value = oZipCode.County
                Params(4).Value = oZipCode.Fips
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutZipCode", Params)
                Return True
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
        Public Function DBGetByKey(ByVal strArray() As String) As MUSTER.Info.ZipCodeInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim oZipCodeInfo As MUSTER.Info.ZipCodeInfo
            Dim Params As Collection

            'strSQL = "select * from tblSYS_ZIPCODES where ZIP='" & strArray(0) & "' and STATE='" & strArray(1) & "' and CITY ='" & strArray(2) & "' and County ='" & strArray(3) & "'"
            Try
                strSQL = "spGetZipCode"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@OrderBy").Value = 1
                Params("@Zip").Value = strArray(0)
                Params("@State").Value = strArray(1)
                Params("@City").Value = strArray(2)
                Params("@County").Value = strArray(3)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.ZipCodeInfo(AltIsDBNull(drSet.Item("Zip"), String.Empty), _
                                      AltIsDBNull(drSet.Item("state"), String.Empty), _
                                      AltIsDBNull(drSet.Item("City"), String.Empty), _
                                      AltIsDBNull(drSet.Item("County"), String.Empty), _
                                      AltIsDBNull(drSet.Item("Fips"), String.Empty), _
                                      AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                      AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                      AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                      AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))

                Else
                    Return New MUSTER.Info.ZipCodeInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
#End Region
    End Class
End Namespace



