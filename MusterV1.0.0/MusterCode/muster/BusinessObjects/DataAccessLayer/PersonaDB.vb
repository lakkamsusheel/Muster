'-------------------------------------------------------------------------------
' MUSTER.DataAccess.PersonaDB
'   Provides the means for marshalling Address to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       EN    12/13/04    Original class definition.
'  1.1       EN      12/28/04    Changed the SqlQuery. Changed the ID from strID = "P" & "|" & drSet.Item("ID") to  strID = "P" & "|" & drSet.Item("Person_ID")
'  1.2       AN      12/30/04    Added Try catch and Exception Handling/Logging
'  1.3       EN                  Modified 01/01/1901  to 01/01/0001 
'  1.4      JVCII    02/11/05    Modified PUT to account for ID which is completely empty (no ID value present)
'  1.5       AB      02/15/05    Replaced dynamic SQL with stored procedures in the following
'                                    Functions:  GetAllInfo, DBGetByKey
'  1.6       AB      02/15/05    Added Finally to the Try/Catch to close all datareaders
'  1.7       AB      02/16/05    Removed any IsNull calls for fields the DB requires
'  1.8       AB      02/18/05    Set all parameters for SP, that are not required, to NULL
'  1.9       AB      02/28/05    Modified Get functions based upon changes made to 
'                                     make several nullable fields non-nullable
'
' Function                  Description
' GetAllInfo()        Returns an PersonaCollection containing all Persona objects in the repository.
' DBGetByName(NAME)   Returns an Personainfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an Personainfo object indicated by arg ID.
' DBGetDS(SQL)      Returns a resultant Dataset by running query specified by the string arg SQL
' Put(opersona)       Saves the Persona passed as an argument, to the DB
''-------------------------------------------------------------------------------

Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class PersonaDB
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
        Public Function GetAllInfo(Optional ByVal showDeleted As Boolean = False) As Muster.Info.PersonaCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            'strSQL = "SELECT * FROM vPersonandOrgInfo WHERE 1=1"
            'strSQL += IIf(Not showDeleted, " AND DELETED <> 1 ", "")
            'strSQL += " ORDER BY ID"

            Try
                strSQL = "spGetPersona"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Person_ID").Value = DBNull.Value
                Params("@Organization_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colPersona As New MUSTER.Info.PersonaCollection
                Dim strID As String
                While drSet.Read
                    If AltIsDBNull(drSet.Item("Person_ID"), 0) <> 0 Then
                        strID = "P" & "|" & drSet.Item("Person_ID")
                    ElseIf AltIsDBNull(drSet.Item("Organization_ID"), 0) <> 0 Then
                        strID = "O" & "|" & drSet.Item("Organization_ID")
                    End If
                    Dim oPersonainfo As New MUSTER.Info.PersonaInfo(strID, _
                                      drSet.Item("Person_id"), _
                                      drSet.Item("Organization_ID"), _
                                      drSet.Item("Organization_Entity_code"), _
                                      drSet.Item("CompanyName"), _
                                      drSet.Item("Title"), _
                                      drSet.Item("Prefix"), _
                                      drSet.Item("First_name"), _
                                      drSet.Item("Middle_name"), _
                                      drSet.Item("Last_name"), _
                                      drSet.Item("Suffix"), _
                                      drSet.Item("DELETED"), _
                                      drSet.Item("CREATED_BY"), _
                                      drSet.Item("DATE_CREATED"), _
                                      AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                      AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))


                    colPersona.Add(oPersonainfo)
                End While

                Return colPersona
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByKey(ByVal strArray() As String, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.PersonaInfo

            Dim drSet As SqlDataReader
            Dim strID As String
            Dim strSQL As String
            Dim Params As Collection

            'strSQL = "SELECT * FROM vPersonandOrgInfo"
            'If strArray(0).ToUpper = "O" Then
            '    strSQL += " WHERE Person_ID = 0 and Organization_ID = " + strArray(1)
            'ElseIf strArray(0).ToUpper = "P" Then
            '    strSQL += " WHERE Organization_ID = 0 and Person_ID = " + strArray(1)
            'End If
            'If Not showDeleted Then
            '    strSQL += " AND DELETED <> 1"
            'End If

            Try
                strSQL = "spGetPersona"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)

                If strArray(0).ToUpper = "O" Then
                    Params("@Organization_ID").Value = CInt(strArray(1))
                    Params("@Person_ID").Value = 0
                Else
                    Params("@Organization_ID").Value = 0
                    Params("@Person_ID").Value = CInt(strArray(1))
                End If

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    If AltIsDBNull(drSet.Item("Person_ID"), 0) <> 0 Then
                        strID = "P" & "|" & drSet.Item("Person_ID")
                    ElseIf AltIsDBNull(drSet.Item("Organization_ID"), 0) <> 0 Then
                        strID = "O" & "|" & drSet.Item("Organization_ID")
                    End If
                    Return New MUSTER.Info.PersonaInfo(strID, _
                                      drSet.Item("Person_id"), _
                                      drSet.Item("Organization_ID"), _
                                      drSet.Item("Organization_Entity_code"), _
                                      drSet.Item("CompanyName"), _
                                      drSet.Item("Title"), _
                                      drSet.Item("Prefix"), _
                                      drSet.Item("First_name"), _
                                      drSet.Item("Middle_name"), _
                                      drSet.Item("Last_name"), _
                                      drSet.Item("Suffix"), _
                                      drSet.Item("DELETED"), _
                                      drSet.Item("CREATED_BY"), _
                                      drSet.Item("DATE_CREATED"), _
                                      AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                      AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))

                Else
                    Return New MUSTER.Info.PersonaInfo
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
        Public Function Put(ByRef oPersona As MUSTER.Info.PersonaInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String) As String
            Dim BnewFlag As Boolean
            Dim strArray() As String

            BnewFlag = False
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Owner, Integer))) Then
                    returnVal = "You do not have rights to save a Owner."
                    Exit Function
                Else
                    returnVal = String.Empty
                End If

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Persona, Integer))) Then
                    returnVal = "You do not have rights to save a Persona."
                    Exit Function
                Else
                    returnVal = String.Empty
                End If

                strArray = Split(oPersona.ID, "|")
                If oPersona.PersonId <> 0 Or (oPersona.FirstName <> "" And oPersona.LastName <> "") Then
                    Dim Params(8) As SqlParameter
                    Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutREGPERSONMASTER")
                    If oPersona.ID = String.Empty Then
                        Params(0).Value = 0
                        BnewFlag = True
                    Else
                        If strArray(1) <= 0 Then
                            Params(0).Value = 0
                            BnewFlag = True
                        ElseIf oPersona.PersonId = 0 Then
                            BnewFlag = True
                            Params(0).Value = 0
                        Else
                            Params(0).Value = oPersona.PersonId
                        End If
                    End If
                    Params(0).Direction = ParameterDirection.InputOutput
                    params(1).Value = oPersona.FirstName
                    params(2).Value = oPersona.Title
                    params(3).Value = oPersona.MiddleName
                    params(4).Value = oPersona.Prefix
                    params(5).Value = oPersona.LastName
                    params(6).Value = oPersona.Suffix

                    If oPersona.PersonId <= 0 Then
                        params(7).Value = oPersona.CreatedBy
                    Else
                        params(7).Value = oPersona.ModifiedBy
                    End If

                    SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutREGPERSONMASTER", Params)
                    If BnewFlag = True And Params(0).Value <> 0 Then
                        Return "P" & "|" & CStr(Params(0).Value)
                    End If
                Else

                    If oPersona.OrgID <> 0 Or (oPersona.Company <> "") Then
                        Dim Params(4) As SqlParameter
                        Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutREGORGANIZATIONMASTER")
                        If oPersona.ID = String.Empty Then
                            Params(0).Value = 0
                            BnewFlag = True
                        Else
                            If strArray(1) <= 0 Then
                                Params(0).Value = 0
                                BnewFlag = True
                            ElseIf oPersona.OrgID = 0 Then
                                Params(0).Value = 0
                                BnewFlag = True
                            Else
                                Params(0).Value = oPersona.OrgID
                            End If
                        End If
                        Params(0).Direction = ParameterDirection.InputOutput
                        Params(1).Value = oPersona.Company
                        Params(2).Value = oPersona.Org_Entity_Code

                        If oPersona.OrgID <= 0 Then
                            params(3).Value = oPersona.CreatedBy
                        Else
                            params(3).Value = oPersona.ModifiedBy
                        End If

                        SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutREGORGANIZATIONMASTER", Params)
                        If BnewFlag = True And Params(0).Value <> 0 Then
                            Return "O" & "|" & CStr(Trim(Params(0).Value))
                        End If
                    End If
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
    End Class
End Namespace
