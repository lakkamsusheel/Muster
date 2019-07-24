'-------------------------------------------------------------------------------
' MUSTER.DataAccess.CompartmentDB
'   Provides the means for marshalling Address to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       EN      12/13/04    Original class definition.
'  1.1       AN      12/30/04    Added Try catch and Exception Handling/Logging
'  1.1       EN      01/27/05    Added DBgetByTankID.
'  1.2       EN      02/10/05    Modified 01/01/1901  to 01/01/0001 
'  1.3       AB      02/14/05    Replaced dynamic SQL with stored procedures in the following
'                                 Functions:  DBGetByTankID, DBGetAllInfo, DBGetByKey
'  1.4       AB      02/15/05    Added Finally to the Try/Catch to close all datareaders
'  1.5       AB      02/16/05    Removed any IsNull calls for fields the DB requires
'  1.6       AB      02/23/05    Modified Get and Put functions based upon changes made to 
'                                   make several nullable fields non-nullable
'  1.7       MNR     03/14/05    Modified the sp name from spPutREGCOMPARTMENTS_Test to spPutREGCOMPARTMENTS in Sub Put
'
' Function                  Description
' GetAllInfo()        Returns an CompartmentCollection containing all Persona objects in the repository.
' DBGetByName(NAME)   Returns an CompartmentInfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an CompartmentInfo object indicated by arg ID.
' DBGetDS(SQL)      Returns a resultant Dataset by running query specified by the string arg SQL
' DBGetByKey 
' Put(oCompartment)       Saves the Persona passed as an argument, to the DB
''-------------------------------------------------------------------------------

Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class CompartmentDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
        Public Function DBGetByTankID(ByVal Tankid As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.CompartmentCollection
            Dim strVal As String
            Dim strSQL As String
            Dim drSet As SqlDataReader
            Dim Params As Collection
            Dim colCompartment As New MUSTER.Info.CompartmentCollection
            Try
                If Tankid = 0 Then
                    Return colCompartment
                End If
                strVal = Tankid
                strSQL = "spGetTankCompartments"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Tank_ID").Value = Tankid
                Params("@Compartment_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = IIf(showDeleted, True, False)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    While drSet.Read
                        Dim oCompartment As New MUSTER.Info.CompartmentInfo(drSet.Item("TANK_ID"), _
                                  drSet.Item("COMPARTMENT_NUMBER"), _
                                  AltIsDBNull(drSet.Item("CAPACITY"), 0), _
                                  AltIsDBNull(drSet.Item("CERCLA#"), 0), _
                                  AltIsDBNull(drSet.Item("SUBSTANCE"), 0), _
                                  AltIsDBNull(drSet.Item("FUEL_TYPE_ID"), 0), _
                                  drSet.Item("DELETED"), _
                                  drSet.Item("CREATED_BY"), _
                                  drSet.Item("DATE_CREATED"), _
                                  AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                  AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))

                        colCompartment.Add(oCompartment)
                    End While
                End If

                Return colCompartment
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBGetAllInfo(Optional ByVal Tankid As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.CompartmentCollection
            Dim intDeleted As Integer
            Dim strSQL As String
            Dim drSet As SqlDataReader
            Dim Params As Collection

            Try
                'strSQL = "SELECT * FROM tblREG_COMPARTMENTS WHERE 1=1"
                'strSQL += IIf(Not showDeleted, " AND DELETED <> 1 ", "")
                'strSQL += IIf(Tankid <> 0, " AND TANK_ID =" & Tankid, "")
                'strSQL += " ORDER BY TANK_ID, COMPARTMENT_NUMBER"

                strSQL = "spGetTankCompartments"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Tank_ID").Value = DBNull.Value
                Params("@Compartment_ID").Value = DBNull.Value

                If Tankid <> 0 Then
                    Params("@Tank_ID").Value = Tankid
                End If

                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colCompartment As New MUSTER.Info.CompartmentCollection
                While drSet.Read
                    Dim oCompartment As New MUSTER.Info.CompartmentInfo(drSet.Item("TANK_ID"), _
                                  drSet.Item("COMPARTMENT_NUMBER"), _
                                  AltIsDBNull(drSet.Item("CAPACITY"), 0), _
                                  AltIsDBNull(drSet.Item("CERCLA#"), 0), _
                                  AltIsDBNull(drSet.Item("SUBSTANCE"), 0), _
                                  AltIsDBNull(drSet.Item("FUEL_TYPE_ID"), 0), _
                                  drSet.Item("DELETED"), _
                                  drSet.Item("CREATED_BY"), _
                                  drSet.Item("DATE_CREATED"), _
                                  AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                  AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))

                    colCompartment.Add(oCompartment)
                End While
                Return colCompartment
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
        Public Function DBGetByKey(ByVal strArray() As String, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.CompartmentInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim oCompartmentInfo As MUSTER.Info.CompartmentInfo
            Dim Params As Collection

            'strSQL = "SELECT * FROM tblREG_COMPARTMENTS WHERE TANK_ID = '" & strArray(0) & "' AND COMPARTMENT_NUMBER = '" & strArray(1) & "' "
            'strSQL += IIf(Not showDeleted, " AND DELETED <> 1", "")
            Try
                If strArray(1) = "0" Then
                    Return New MUSTER.Info.CompartmentInfo
                End If
                strSQL = "spGetTankCompartments"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Tank_ID").Value = strArray(0)
                Params("@Compartment_ID").Value = strArray(1)
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.CompartmentInfo(drSet.Item("TANK_ID"), _
                                  drSet.Item("COMPARTMENT_NUMBER"), _
                                  AltIsDBNull(drSet.Item("CAPACITY"), 0), _
                                  AltIsDBNull(drSet.Item("CERCLA#"), 0), _
                                  AltIsDBNull(drSet.Item("SUBSTANCE"), 0), _
                                  AltIsDBNull(drSet.Item("FUEL_TYPE_ID"), 0), _
                                  drSet.Item("DELETED"), _
                                  drSet.Item("CREATED_BY"), _
                                  drSet.Item("DATE_CREATED"), _
                                  AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                  AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                Else
                    Return New MUSTER.Info.CompartmentInfo
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
        Public Sub Put(ByRef oCompartment As MUSTER.Info.CompartmentInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal strUser As String)
            Dim Params(8) As SqlParameter
            Try
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutREGCOMPARTMENTS")
                If oCompartment.TankId <= 0 Then
                    Throw New Exception("TankID (" + oCompartment.TankId.ToString + ") cannot be < 0 for Compartment: " + oCompartment.COMPARTMENTNumber.ToString)
                    Exit Sub
                Else
                    Params(0).Value = oCompartment.TankId
                End If
                If oCompartment.COMPARTMENTNumber < 0 Then
                    Params(1).Value = 0
                    Params(1).Direction = ParameterDirection.Output

                Else
                    Params(1).Value = oCompartment.COMPARTMENTNumber
                    Params(1).Direction = ParameterDirection.InputOutput

                End If
                If oCompartment.Capacity = 0 Then
                    Params(2).Value = System.DBNull.Value
                Else
                    Params(2).Value = oCompartment.Capacity
                End If
                If oCompartment.CCERCLA = 0 Then
                    Params(3).Value = System.DBNull.Value
                Else
                    Params(3).Value = oCompartment.CCERCLA
                End If
                If oCompartment.Substance = 0 Then
                    Params(4).Value = System.DBNull.Value
                Else
                    Params(4).Value = oCompartment.Substance
                End If
                If oCompartment.FuelTypeId = 0 Then
                    Params(5).Value = System.DBNull.Value
                Else
                    Params(5).Value = oCompartment.FuelTypeId
                End If
                Params(6).Value = oCompartment.Deleted
                Params(7).Value = strUser
                'If oCompartment.COMPARTMENTNumber <= 0 Then
                '    Params(7).Value = oCompartment.CreatedBy
                'Else
                '    Params(7).Value = oCompartment.ModifiedBy
                'End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutREGCOMPARTMENTS", Params)
                If Params(1).Value <> oCompartment.COMPARTMENTNumber Then
                    oCompartment.COMPARTMENTNumber = Params(1).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function getManifold(ByVal nTank_ID As Integer) As DataSet
            Dim Params() As SqlParameter
            Dim strSql As String
            Dim dsData As DataSet
            Try
                strSql = "spGetManifoldInfo"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSql)
                Params(0).Value = nTank_ID.ToString
                'Params(0).Value = nfacility_ID.ToString
                'Params(1).Value = nTank_ID.ToString
                'Params(2).Value = nCompartment_Number.ToString
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSql, Params)
                Return dsData
            Catch ex As Exception

            End Try
        End Function
    End Class
End Namespace
