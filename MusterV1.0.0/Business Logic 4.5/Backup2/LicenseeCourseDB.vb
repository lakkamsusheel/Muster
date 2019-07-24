'-------------------------------------------------------------------------------
' MUSTER.DataAccess.LicenseeCourseDB
'   Provides the means for marshalling LicenseeCourse to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       KR      05/07/05    Original class definition.
'  1.1       MR      6/3/05      Modified all the methods with New Parameters.
'
' Function                  Description
' 
''-------------------------------------------------------------------------------

Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    <Serializable()> _
        Public Class LicenseeCourseDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
        Public Function DBGetByID(ByVal LicenseeCourseID As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.LicenseeCourseInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            Try
                If LicenseeCourseID <= 0 Then
                    Return New MUSTER.Info.LicenseeCourseInfo
                End If
                strSQL = "spGetCOMLicenseeCourses"
                strVal = LicenseeCourseID.ToString

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@LIC_COUR_ID").Value = strVal
                Params("@LICENSEE_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = False

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()

                    Return New MUSTER.Info.LicenseeCourseInfo(drSet.Item("LIC_COUR_ID"), _
                                            AltIsDBNull(drSet.Item("LICENSEE_ID"), 0), _
                                            AltIsDBNull(drSet.Item("PROVIDER_ID"), 0), _
                                            AltIsDBNull(drSet.Item("COURSE_TYPE_ID"), 0), _
                                            AltIsDBNull(drSet.Item("COURSE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("DELETED"), False), _
                                            AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                Else
                    Return New MUSTER.Info.LicenseeCourseInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not (drSet Is Nothing) Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBGetByLicenseeID(Optional ByVal LicenseeID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.LicenseeCourseCollection
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim strVal As String = ""
            Dim Params As Collection
            Dim colLicCourse As New MUSTER.Info.LicenseeCourseCollection
            Try
                strSQL = "spGetCOMLicenseeCourses"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@LIC_COUR_ID").Value = DBNull.Value
                Params("@LICENSEE_ID").Value = IIf(LicenseeID = 0, DBNull.Value, LicenseeID.ToString)
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = False

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    While drSet.Read
                        Dim LicCourseInfo As New MUSTER.Info.LicenseeCourseInfo(drSet.Item("LIC_COUR_ID"), _
                                            AltIsDBNull(drSet.Item("LICENSEE_ID"), 0), _
                                            AltIsDBNull(drSet.Item("PROVIDER_ID"), 0), _
                                            AltIsDBNull(drSet.Item("COURSE_TYPE_ID"), 0), _
                                            AltIsDBNull(drSet.Item("COURSE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("DELETED"), False), _
                                            AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                        colLicCourse.Add(LicCourseInfo)
                    End While
                End If
                Return colLicCourse
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Sub Put(ByRef oLicenseeCourse As MUSTER.Info.LicenseeCourseInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Company, Integer))) Then
                    returnVal = "You do not have rights to save Licensee Course."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutCOMLicenseeCourses")

                If oLicenseeCourse.ID <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = oLicenseeCourse.ID
                End If

                Params(1).Value = oLicenseeCourse.LicenseeID
                Params(2).Value = oLicenseeCourse.ProviderID
                Params(3).Value = oLicenseeCourse.CourseTypeID
                Params(4).Value = oLicenseeCourse.CourseDate
                Params(5).Value = oLicenseeCourse.Deleted
                Params(6).Value = DBNull.Value
                Params(7).Value = DBNull.Value
                Params(8).Value = DBNull.Value
                Params(9).Value = DBNull.Value
                If oLicenseeCourse.ID <= 0 Then
                    Params(10).Value = oLicenseeCourse.CreatedBy
                Else
                    Params(10).Value = oLicenseeCourse.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutCOMLicenseeCourses", Params)
                If oLicenseeCourse.ID <= 0 Then
                    oLicenseeCourse.ID = Params(0).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function DBGetDS(ByVal strSQL As String) As DataSet
            Dim dsData As DataSet
            Try
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
    End Class
End Namespace
