'-------------------------------------------------------------------------------
' MUSTER.DataAccess.CourseDB
'   Provides the means for marshalling Course state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MR     5/21/04    Original class definition.
'                                  
'
' Function                  Description
' GetAllInfo()      Returns an CourseCollection containing all Course objects in the repository
' DBGetByID(ID)     Returns an CourseInfo object indicated by int arg ID
' DBGetDS(SQL)      Returns a resultant Dataset by running query specified by the string arg SQL
' Put(Entity)       Saves the Course passed as an argument, to the DB
'-------------------------------------------------------------------------------

Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class CourseDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function GetAllInfo() As MUSTER.Info.CourseCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                'AltIsDBNull(drSet.Item("COURSE_DATES"), String.Empty), _
                'AltIsDBNull(drSet.Item("LOCATION"), String.Empty), _
                'AltIsDBNull(drSet.Item("COURSE_TYPE_ID"), 0), _
                'AltIsDBNull(drSet.Item("CREDIT_HOURS"), String.Empty), _

                strSQL = "spGetCOMProviderCourses"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Course_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = False

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colCourses As New MUSTER.Info.CourseCollection
                While drSet.Read

                    Dim oCourseInfo As New MUSTER.Info.CourseInfo(drSet.Item("COURSE_ID"), _
                                            AltIsDBNull(drSet.Item("ACTIVE"), False), _
                                            AltIsDBNull(drSet.Item("PROVIDER_ID"), 0), _
                                            AltIsDBNull(drSet.Item("COURSE_TITLE"), String.Empty), _
                                            AltIsDBNull(drSet.Item("PROVIDER_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DELETED"), False), _
                                            AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))

                    colCourses.Add(oCourseInfo)
                End While

                Return colCourses
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByID(ByVal nVal As Integer) As MUSTER.Info.CourseInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            'AltIsDBNull(drSet.Item("COURSE_DATES"), String.Empty), _
            'AltIsDBNull(drSet.Item("LOCATION"), String.Empty), _
            'AltIsDBNull(drSet.Item("COURSE_TYPE_ID"), 0), _
            'AltIsDBNull(drSet.Item("CREDIT_HOURS"), String.Empty), _

            Try
                strSQL = "spGetCOMProviderCourses"
                strVal = nVal

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Course_ID").Value = strVal
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = False

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.CourseInfo(drSet.Item("COURSE_ID"), _
                                            AltIsDBNull(drSet.Item("ACTIVE"), False), _
                                            AltIsDBNull(drSet.Item("PROVIDER_ID"), 0), _
                                            AltIsDBNull(drSet.Item("COURSE_TITLE"), String.Empty), _
                                            AltIsDBNull(drSet.Item("PROVIDER_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DELETED"), False), _
                                            AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                Else
                    Return New MUSTER.Info.CourseInfo
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
        Public Sub Put(ByRef oCourseInf As MUSTER.Info.CourseInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Licensee, Integer))) Then
                    returnVal = "You do not have rights to save License Course."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutCOMProvidersCourses")

                If oCourseInf.ID <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = oCourseInf.ID
                End If
                Params(1).Value = oCourseInf.Active
                Params(2).Value = oCourseInf.ProviderID
                Params(3).Value = oCourseInf.CourseTitle
                'Params(4).Value = oCourseInf.CourseDates
                'Params(5).Value = oCourseInf.Location
                'Params(6).Value = oCourseInf.CourseTypeID
                Params(4).Value = oCourseInf.ProviderName
                'Params(8).Value = oCourseInf.CreditHours
                Params(5).Value = oCourseInf.Deleted
                Params(6).Value = DBNull.Value
                Params(7).Value = DBNull.Value
                Params(8).Value = DBNull.Value
                Params(9).Value = DBNull.Value

                If oCourseInf.ID <= 0 Then
                    Params(10).Value = oCourseInf.CreatedBy
                Else
                    Params(10).Value = oCourseInf.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutCOMProvidersCourses", Params)
                If oCourseInf.ID <= 0 Then
                    oCourseInf.ID = Params(0).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function DBGetCourseDates(Optional ByVal CourseID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try
                strSQL = "spGetComProviderCourseDates"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(CourseID = 0, DBNull.Value, CourseID)
                Params(1).Value = showDeleted
                'Params(2).Value = 1

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBPutCourseDates(ByVal nCourseDatesID As Integer, _
                                    ByVal nCourseID As Integer, _
                                     ByVal nCourseDatesNum As Integer, _
                                    ByVal nCourseDates As String, _
                                    ByVal strLocation As String, _
                                    ByVal nCourseType As Integer, _
                                    ByVal strHours As String, _
                                    ByVal bolDeleted As Boolean, _
                                    ByVal UserID As String) As Boolean
            DBPutCourseDates = False
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try
                strSQL = "spPutCOMProviderCourseDates"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                If nCourseDatesID <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = nCourseDatesID
                End If
                Params(1).Value = 1
                Params(2).Value = nCourseID
                Params(3).Value = nCourseDatesNum
                Params(4).Value = nCourseDates
                Params(5).Value = strLocation
                Params(6).Value = nCourseType
                Params(7).Value = strHours
                Params(8).Value = bolDeleted
                Params(9).Value = DBNull.Value
                Params(10).Value = DBNull.Value
                Params(11).Value = DBNull.Value
                Params(12).Value = DBNull.Value
                Params(13).Value = UserID

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If Params(0).Value <> nCourseDatesID Then
                    nCourseDatesID = Params(0).Value
                End If
                Return True
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
    End Class
End Namespace

