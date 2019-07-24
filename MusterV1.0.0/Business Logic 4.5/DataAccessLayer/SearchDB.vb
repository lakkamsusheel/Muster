'-------------------------------------------------------------------------------
' MUSTER.DataAccess.SearchDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MNR     12/09/04    Original class definition.
'  1.1        AN      12/30/04    Added Try catch and Exception Handling/Logging
'  1.2        EN      01/19/05    Added new DBGetDS and commented the old one. 
'
' Function          Description
' DBGetDS(Keyword, Module, Filter)
'                   Returns a resultant Dataset by running query for string arg Keyword
'                   associated with the string args Module, and Filter
'
'-------------------------------------------------------------------------------
'
' TODO - Check spQuick_Search for proper resultset return...
'

Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class SearchDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
        'Public Function DBGetDS(ByVal strKeyword As String, ByVal strModule As String, ByVal strFilter As String) As DataSet
        '    Dim dsData As New DataSet
        '    Dim strSQL As String
        '    Try
        '        Select Case strModule
        '            Case "Registration"
        '                Select Case strFilter
        '                    Case "Owner ID"
        '                        strSQL = "SELECT tblREG_OWNER.OWNER_ID AS OwnerID, "
        '                        strSQL += "tblREG_PERSON_MASTER.FIRST_NAME AS OwnerName, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.ADDRESS_LINE_ONE AS Address, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.CITY AS City, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.STATE AS State, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.ZIP AS Zip "
        '                        strSQL += "FROM tblREG_OWNER INNER JOIN tblREG_ADDRESS_MASTER ON "
        '                        strSQL += "tblREG_OWNER.ADDRESS_ID = tblREG_ADDRESS_MASTER.ADDRESS_ID INNER JOIN tblREG_PERSON_MASTER ON "
        '                        strSQL += "tblREG_OWNER.PERSON_ID = tblREG_PERSON_MASTER.PERSON_ID "
        '                        strSQL += "WHERE (tblREG_OWNER.OWNER_ID = " + strKeyword + ") "
        '                        strSQL += "ORDER BY tblREG_OWNER.OWNER_ID"
        '                    Case "Owner Name"
        '                        strSQL = "SELECT tblREG_OWNER.OWNER_ID AS OwnerID, "
        '                        strSQL += "tblREG_PERSON_MASTER.FIRST_NAME AS OwnerName, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.ADDRESS_LINE_ONE AS Address, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.CITY AS City, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.STATE AS State, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.ZIP AS Zip "
        '                        strSQL += "FROM tblREG_OWNER INNER JOIN tblREG_ADDRESS_MASTER ON "
        '                        strSQL += "tblREG_OWNER.ADDRESS_ID = tblREG_ADDRESS_MASTER.ADDRESS_ID INNER JOIN tblREG_PERSON_MASTER ON "
        '                        strSQL += "tblREG_OWNER.PERSON_ID = tblREG_PERSON_MASTER.PERSON_ID "
        '                        strSQL += "WHERE (tblREG_PERSON_MASTER.FIRST_NAME LIKE '%" + strKeyword + "%') "
        '                        strSQL += "ORDER BY tblREG_PERSON_MASTER.FIRST_NAME"
        '                    Case "Owner Address"
        '                        strSQL = "SELECT tblREG_OWNER.OWNER_ID AS OwnerID, "
        '                        strSQL += "tblREG_PERSON_MASTER.FIRST_NAME AS OwnerName, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.ADDRESS_LINE_ONE AS Address, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.CITY AS City, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.STATE AS State, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.ZIP AS Zip "
        '                        strSQL += "FROM tblREG_OWNER INNER JOIN tblREG_ADDRESS_MASTER ON "
        '                        strSQL += "tblREG_OWNER.ADDRESS_ID = tblREG_ADDRESS_MASTER.ADDRESS_ID INNER JOIN tblREG_PERSON_MASTER ON "
        '                        strSQL += "tblREG_OWNER.PERSON_ID = tblREG_PERSON_MASTER.PERSON_ID "
        '                        strSQL += "WHERE (tblREG_ADDRESS_MASTER.ADDRESS_LINE_ONE LIKE '%" + strKeyword + "%') "
        '                        strSQL += "ORDER BY tblREG_ADDRESS_MASTER.ADDRESS_LINE_ONE"
        '                    Case "Facility ID"
        '                        strSQL = "(SELECT tblREG_FACILITY.NAME AS FacilityName, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.ADDRESS_LINE_ONE AS Address, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.CITY AS City, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.STATE AS State, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.ZIP AS Zip, "
        '                        strSQL += "tblREG_FACILITY.OWNER_ID AS FacilityID, "
        '                        strSQL += "tblREG_ORGANIZATION_MASTER.NAME AS OrgName "
        '                        strSQL += "FROM tblREG_FACILITY INNER JOIN tblREG_ADDRESS_MASTER ON "
        '                        strSQL += "tblREG_FACILITY.ADDRESS_ID = tblREG_ADDRESS_MASTER.ADDRESS_ID INNER JOIN tblREG_OWNER ON "
        '                        strSQL += "tblREG_FACILITY.OWNER_ID = tblREG_OWNER.OWNER_ID INNER JOIN tblREG_ORGANIZATION_MASTER ON "
        '                        strSQL += "tblREG_OWNER.ORGANIZATION_ID = tblREG_ORGANIZATION_MASTER.ORGANIZATION_ID "
        '                        strSQL += "WHERE (tblREG_FACILITY.OWNER_ID = '" + strKeyword + "'))"
        '                        strSQL += " UNION "
        '                        strSQL += "(SELECT tblREG_PERSON_MASTER.FIRST_NAME, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.ADDRESS_LINE_ONE, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.CITY, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.STATE, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.ZIP, "
        '                        strSQL += "NULL, "
        '                        strSQL += "tblREG_FACILITY.NAME "
        '                        strSQL += "FROM tblREG_FACILITY INNER JOIN tblREG_ADDRESS_MASTER ON "
        '                        strSQL += "tblREG_FACILITY.ADDRESS_ID = tblREG_ADDRESS_MASTER.ADDRESS_ID INNER JOIN tblREG_OWNER ON "
        '                        strSQL += "tblREG_FACILITY.OWNER_ID = tblREG_OWNER.OWNER_ID INNER JOIN tblREG_PERSON_MASTER ON "
        '                        strSQL += "tblREG_OWNER.PERSON_ID = tblREG_PERSON_MASTER.PERSON_ID "
        '                        strSQL += "WHERE (tblREG_FACILITY.DELETED = '0' AND tblREG_FACILITY.FACILITY_ID = '" + strKeyword + "')"
        '                        strSQL += "ORDER BY tblREG_FACILITY.OWNER_ID)"
        '                    Case "Facility Name"
        '                        strSQL = "(SELECT tblREG_FACILITY.NAME AS FacilityName, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.ADDRESS_LINE_ONE AS Address, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.CITY AS City, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.STATE AS State, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.ZIP AS Zip, "
        '                        strSQL += "tblREG_FACILITY.OWNER_ID AS FacilityID, "
        '                        strSQL += "tblREG_ORGANIZATION_MASTER.NAME AS OrgName "
        '                        strSQL += "FROM tblREG_FACILITY INNER JOIN tblREG_ADDRESS_MASTER ON "
        '                        strSQL += "tblREG_FACILITY.ADDRESS_ID = tblREG_ADDRESS_MASTER.ADDRESS_ID INNER JOIN tblREG_OWNER ON "
        '                        strSQL += "tblREG_FACILITY.OWNER_ID = tblREG_OWNER.OWNER_ID INNER JOIN tblREG_ORGANIZATION_MASTER ON "
        '                        strSQL += "tblREG_OWNER.ORGANIZATION_ID = tblREG_ORGANIZATION_MASTER.ORGANIZATION_ID "
        '                        strSQL += "WHERE (tblREG_FACILITY.NAME LIKE '%" + strKeyword + "%'))"
        '                        strSQL += " UNION "
        '                        strSQL += "(SELECT tblREG_PERSON_MASTER.FIRST_NAME, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.ADDRESS_LINE_ONE, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.CITY, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.STATE, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.ZIP, "
        '                        strSQL += "NULL, "
        '                        strSQL += "tblREG_FACILITY.NAME "
        '                        strSQL += "FROM tblREG_FACILITY INNER JOIN tblREG_ADDRESS_MASTER ON "
        '                        strSQL += "tblREG_FACILITY.ADDRESS_ID = tblREG_ADDRESS_MASTER.ADDRESS_ID INNER JOIN tblREG_OWNER ON "
        '                        strSQL += "tblREG_FACILITY.OWNER_ID = tblREG_OWNER.OWNER_ID INNER JOIN tblREG_PERSON_MASTER ON "
        '                        strSQL += "tblREG_OWNER.PERSON_ID = tblREG_PERSON_MASTER.PERSON_ID "
        '                        strSQL += "WHERE (tblREG_FACILITY.NAME LIKE '%" + strKeyword + "%'))"
        '                        strSQL += "ORDER BY tblREG_FACILITY.OWNER_ID"
        '                    Case "Facility Address"
        '                        strSQL = "(SELECT tblREG_FACILITY.NAME AS FacilityName, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.ADDRESS_LINE_ONE AS Address, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.CITY AS City, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.STATE AS State, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.ZIP AS Zip, "
        '                        strSQL += "tblREG_FACILITY.OWNER_ID AS FacilityID, "
        '                        strSQL += "tblREG_ORGANIZATION_MASTER.NAME AS OrgName "
        '                        strSQL += "FROM tblREG_FACILITY INNER JOIN tblREG_ADDRESS_MASTER ON "
        '                        strSQL += "tblREG_FACILITY.ADDRESS_ID = tblREG_ADDRESS_MASTER.ADDRESS_ID INNER JOIN tblREG_OWNER ON "
        '                        strSQL += "tblREG_FACILITY.OWNER_ID = tblREG_OWNER.OWNER_ID INNER JOIN tblREG_ORGANIZATION_MASTER ON "
        '                        strSQL += "tblREG_OWNER.ORGANIZATION_ID = tblREG_ORGANIZATION_MASTER.ORGANIZATION_ID "
        '                        strSQL += "WHERE (tblREG_ADDRESS_MASTER.ADDRESS_LINE_ONE LIKE '%" + strKeyword + "%'))"
        '                        strSQL += " UNION "
        '                        strSQL += "(SELECT tblREG_PERSON_MASTER.FIRST_NAME, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.ADDRESS_LINE_ONE, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.CITY, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.STATE, "
        '                        strSQL += "tblREG_ADDRESS_MASTER.ZIP, "
        '                        strSQL += "NULL, "
        '                        strSQL += "tblREG_FACILITY.NAME "
        '                        strSQL += "FROM tblREG_FACILITY INNER JOIN tblREG_ADDRESS_MASTER ON "
        '                        strSQL += "tblREG_FACILITY.ADDRESS_ID = tblREG_ADDRESS_MASTER.ADDRESS_ID INNER JOIN tblREG_OWNER ON "
        '                        strSQL += "tblREG_FACILITY.OWNER_ID = tblREG_OWNER.OWNER_ID INNER JOIN tblREG_PERSON_MASTER ON "
        '                        strSQL += "tblREG_OWNER.PERSON_ID = tblREG_PERSON_MASTER.PERSON_ID "
        '                        strSQL += "WHERE (tblREG_ADDRESS_MASTER.ADDRESS_LINE_ONE LIKE '%" + strKeyword + "%')) "
        '                        strSQL += "ORDER BY tblREG_FACILITY.OWNER_ID"
        '                End Select

        '            Case "Technical"
        '                MsgBox("Not Yet Implemented")
        '                Return dsData
        '            Case "C & E"
        '                MsgBox("Not Yet Implemented")
        '                Return dsData
        '            Case "Company"
        '                MsgBox("Not Yet Implemented")
        '                Return dsData
        '            Case "Inspector"
        '                MsgBox("Not Yet Implemented")
        '                Return dsData
        '        End Select

        '        dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
        '        Return dsData
        '    Catch Ex As Exception
        '        MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function
        Public Function DBGetDS(ByVal strKeyword As String, ByVal strFilter As String, ByVal strOperator As String) As DataSet
            Try
                Dim Params(2) As SqlParameter
                Dim dsSet As DataSet
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spQuick_Search")
                Params(0).Value = strFilter
                Params(1).Value = strKeyword
                Params(2).Value = strOperator
                dsSet = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, "spQuick_Search", Params)
                Return dsSet
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetsearchFilter(ByVal strSQL As String) As DataSet
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
