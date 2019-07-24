'-------------------------------------------------------------------------------
' MUSTER.DataAccess.EntityDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC2      11/19/04    Original class definition.
'  1.1        AN        12/30/04    Added Try catch and Exception Handling/Logging
'  1.2        JVC2      02/07/05    Fixed DBGetByID() to use reader instead of dataset
'  1.3        EN        02/10/05    Modified 01/01/1901  to 01/01/0001 
'  1.4        AB        02/14/05    Replaced dynamic SQL with stored procedures in the following
'                                   Functions:  DBGetByID, GetAllInfo, DBGetByName
'  1.5        AB        02/15/05    Added Finally to the Try/Catch to close all datareaders
'  1.6        AB        02/16/05    Removed any IsNull calls for fields the DB requires
'  1.7        AB        02/18/05    Set all parameters for SP, that are not required, to NULL
'  1.8        AB        02/23/05    Modified Get and Put functions based upon changes made to 
'                                   make several nullable fields non-nullable
'
' Function                  Description
' GetAllEntityInfo()        Returns an EntityCollection containing all Entity objects in the repository.
' checkEntityByName(NAME)   Returns an EntityInfo object indicated by arg NAME.
' checkEntityByID(ID)       Returns an EntityInfo object indicated by arg ID.
' New(ds)           Instantiates a populated EntityInfo object taking member state
'                       from the first row in the first table in the dataset provided
' Reset()           Sets the object state to the original state when loaded from or
'                       last saved to the repository.
' Save()            Saves the object state to the repository.
'
'Attribute          Description
' ID                The unique identifier associated with the Entity in the repository.
' Name              The name of the Entity.
' IsDirty           Indicates if the Entity state has been altered since it was
'                       last loaded from or saved to the repository.
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class EntityDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
#Region "Miscellaneous Functions"
        Public Function GetEntityTypeDescByID(ByVal EntityTypeID As Integer)
            Try
                Dim EntityName As String = String.Empty

                Select Case EntityTypeID
                    Case 4
                        EntityName = SqlHelper.EntityTypes.Contact.ToString
                    Case 5
                        EntityName = SqlHelper.EntityTypes.Contractor.ToString
                    Case 6
                        EntityName = SqlHelper.EntityTypes.Facility.ToString
                    Case 7
                        EntityName = SqlHelper.EntityTypes.LUST_Event.ToString
                    Case 8
                        EntityName = SqlHelper.EntityTypes.NONE.ToString
                    Case 9
                        EntityName = SqlHelper.EntityTypes.Owner.ToString
                    Case 10
                        EntityName = SqlHelper.EntityTypes.Pipe.ToString
                    Case 11
                        EntityName = SqlHelper.EntityTypes.Report.ToString
                    Case 12
                        EntityName = SqlHelper.EntityTypes.Tank.ToString
                    Case 13
                        EntityName = SqlHelper.EntityTypes.Violation.ToString
                    Case 14
                        EntityName = SqlHelper.EntityTypes.Organization.ToString
                    Case 15
                        EntityName = SqlHelper.EntityTypes.User.ToString
                    Case 16
                        EntityName = SqlHelper.EntityTypes.Persona.ToString
                    Case 17
                        EntityName = SqlHelper.EntityTypes.Flag.ToString
                    Case 18
                        EntityName = SqlHelper.EntityTypes.Calendar.ToString
                    Case 19
                        EntityName = SqlHelper.EntityTypes.Letter.ToString
                    Case 20
                        EntityName = SqlHelper.EntityTypes.Profile.ToString
                    Case 21
                        EntityName = SqlHelper.EntityTypes.Address.ToString
                    Case 22
                        EntityName = SqlHelper.EntityTypes.ClosureEvent.ToString
                    Case 23
                        EntityName = SqlHelper.EntityTypes.LustActivity.ToString
                    Case 24
                        EntityName = SqlHelper.EntityTypes.LustDocument.ToString
                    Case 25
                        EntityName = SqlHelper.EntityTypes.Fees.ToString
                    Case 26
                        EntityName = SqlHelper.EntityTypes.Company.ToString
                    Case 27
                        EntityName = SqlHelper.EntityTypes.Licensee.ToString
                    Case 28
                        EntityName = SqlHelper.EntityTypes.Provider.ToString
                    Case 29
                        EntityName = SqlHelper.EntityTypes.Financial.ToString
                    Case 30
                        EntityName = SqlHelper.EntityTypes.Inspection.ToString
                    Case 31
                        EntityName = SqlHelper.EntityTypes.CAE.ToString
                    Case 32
                        EntityName = SqlHelper.EntityTypes.FinancialEvent.ToString
                    Case 33
                        EntityName = SqlHelper.EntityTypes.FinancialCommitment.ToString
                    Case 34
                        EntityName = SqlHelper.EntityTypes.FinancialInvoice.ToString
                    Case 35
                        EntityName = SqlHelper.EntityTypes.FinancialReimbursement.ToString
                    Case 36
                        EntityName = SqlHelper.EntityTypes.CAEOwnerComplianceEvent.ToString
                    Case 37
                        EntityName = SqlHelper.EntityTypes.CAELicenseeCompliantEvent.ToString
                    Case 38
                        EntityName = SqlHelper.EntityTypes.TechnicalActivity.ToString
                    Case 39
                        EntityName = SqlHelper.EntityTypes.TechnicalDocument.ToString
                    Case 40
                        EntityName = SqlHelper.EntityTypes.LustRemediation.ToString
                    Case 41
                        EntityName = SqlHelper.EntityTypes.CAEFacilityCompliantEvent.ToString
                    Case 42
                        EntityName = SqlHelper.EntityTypes.Comment.ToString
                End Select



                Return EntityName
            Catch ex As Exception

            End Try

        End Function

#End Region

#Region "Exposed Operations"
        Public Function GetAllInfo() As MUSTER.Info.EntityCollection
            Dim strSQL As String
            Dim drSet As SqlDataReader
            Try

                strSQL = "spGetEntity"

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL)
                Dim colEntities As New MUSTER.Info.EntityCollection
                While drSet.Read
                    Dim oEntityInfo As New MUSTER.Info.EntityInfo(drSet.Item("ENTITY_ID"), _
                                                          drSet.Item("ENTITY_NAME"), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("CREATED_DATE"), _
                                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                    colEntities.Add(oEntityInfo)
                End While
                Return colEntities
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByName(ByVal strVal As String) As MUSTER.Info.EntityInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetEntity"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Entity_ID").value = DBNull.Value
                Params("@Entity_Name").Value = strVal
                Params("@OrderBy").Value = 1


                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.EntityInfo(drSet.Item("ENTITY_ID"), _
                                                      drSet.Item("ENTITY_NAME"), _
                                                      drSet.Item("CREATED_BY"), _
                                                      drSet.Item("CREATED_DATE"), _
                                                      AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                      AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                Else
                    Return New MUSTER.Info.EntityInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByID(ByVal nVal As Int64) As MUSTER.Info.EntityInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetEntity"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Entity_ID").Value = nVal
                Params("@OrderBy").Value = 1
                Params.Remove("@Entity_Name")

                'drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_ENTITY WHERE ENTITY_ID = " & nVal.ToString)
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.EntityInfo(drSet.Item("ENTITY_ID"), _
                                                      drSet.Item("ENTITY_NAME"), _
                                                      drSet.Item("CREATED_BY"), _
                                                      drSet.Item("CREATED_DATE"), _
                                                      AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                      AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                Else
                    Return New MUSTER.Info.EntityInfo
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

        Public Sub Put(ByRef obj As MUSTER.Info.EntityInfo)
            Try
                Dim Params(3) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPUTSYSENTITY")
                Params("@ENTITY_ID").Value = obj.ID
                Params("@ENTITY_NAME").Value = obj.Name
                Params("@CREATED_DATE").Value = obj.CreatedOn
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPUTSYSENTITY", Params)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
    End Class
End Namespace
