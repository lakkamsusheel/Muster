'-------------------------------------------------------------------------------
' MUSTER.DataAccess.TecDocDB
'   Provides the means for marshalling Technical Document state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC     05/24/05    Original class definition.
'
' Function                  Description
' DBGetByID(ID)         Returns an LustRemediation Object indicated by Lust Remediation ID
' DBGetByEventID(ID)    Returns an LustRemediation Collection indicated by Lust Event ID
' DBGetDS(SQL)          Returns a resultant Dataset by running query specified by the string arg SQL
' Put(oTecDoc)        Saves the LustRemediation passed as an argument, to the DB
'-------------------------------------------------------------------------------
'
'

Imports System.Data.SqlClient

Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class TecDocDB
#Region "Private Member Variables"
        Private _strConn As Object
        Private MusterException As Exceptions.MusterExceptions = New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Exposed Methods"
        Public Sub New(Optional ByVal strDBConn As String = "", Optional ByRef MusterXCEP As MUSTER.Exceptions.MusterExceptions = Nothing)
            Try
                If MusterXCEP Is Nothing Then
                    MusterException = New MUSTER.Exceptions.MusterExceptions
                Else
                    MusterException = MusterXCEP
                End If
                If strDBConn = String.Empty Then
                    Dim oCnn As New ConnectionSettings
                    _strConn = oCnn.cnString
                    oCnn = Nothing
                Else
                    _strConn = strDBConn
                End If
            Catch ex As Exception
                Throw ex
            End Try
        End Sub
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

        Public Function DBGetByID(ByVal nVal As Int64) As MUSTER.Info.TecDocInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nVal = 0 Then
                    Return New MUSTER.Info.TecDocInfo
                End If
                strSQL = "spGetTecDocuments"
                strVal = nVal

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@DOCUMENT_ID").Value = nVal
                Params("@DocName").Value = ""

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.TecDocInfo(drSet.Item("DocID"), _
                            AltIsDBNull(drSet.Item("Doc_Type"), 0), _
                            AltIsDBNull(drSet.Item("DocName"), AltIsDBNull(drSet.Item("Property_Name"), "")), _
                            AltIsDBNull(drSet.Item("DocFileName"), ""), _
                            AltIsDBNull(drSet.Item("DocTrigger"), 0), _
                            AltIsDBNull(drSet.Item("NTFE_Flag"), 0), _
                            AltIsDBNull(drSet.Item("STFS_Flag"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_1"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_2"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_3"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_4"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_5"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_6"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_7"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_8"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_9"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_10"), 0), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("Active"), 1), _
                            AltIsDBNull(drSet.Item("Deleted"), 0), AltIsDBNull(drSet.Item("FinActivityTypeID"), 0) _
                    )

                Else

                    Return New MUSTER.Info.TecDocInfo
                End If
            Catch ex As Exception

                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBGetByName(ByVal strVal As String) As MUSTER.Info.TecDocInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                If strVal = "" Then
                    Return New MUSTER.Info.TecDocInfo
                End If
                strSQL = "spGetTecDocuments"
                strVal = strVal

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@DocName").Value = strVal
                Params("@DOCUMENT_ID").Value = 0

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.TecDocInfo(drSet.Item("DocID"), _
                            AltIsDBNull(drSet.Item("Doc_Type"), 0), _
                            AltIsDBNull(drSet.Item("DocName"), AltIsDBNull(drSet.Item("Property_Name"), "")), _
                            AltIsDBNull(drSet.Item("DocFileName"), ""), _
                            AltIsDBNull(drSet.Item("DocTrigger"), 0), _
                            AltIsDBNull(drSet.Item("NTFE_Flag"), 0), _
                            AltIsDBNull(drSet.Item("STFS_Flag"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_1"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_2"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_3"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_4"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_5"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_6"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_7"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_8"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_9"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_10"), 0), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("Active"), 1), _
                            AltIsDBNull(drSet.Item("Deleted"), 0), AltIsDBNull(drSet.Item("FinActivityTypeID"), 0) _
                    )

                Else

                    Return New MUSTER.Info.TecDocInfo
                End If
            Catch ex As Exception

                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function

        Public Function GetByActivity(ByVal nVal As Int64) As MUSTER.Info.TecDocCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetTecDocuments_ByActivity"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Activity_ID").Value = nVal

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colEntities As New MUSTER.Info.TecDocCollection
                While drSet.Read
                    Dim oTecDocInfo As New MUSTER.Info.TecDocInfo(drSet.Item("DocID"), _
                            AltIsDBNull(drSet.Item("Doc_Type"), 0), _
                            AltIsDBNull(drSet.Item("DocName"), AltIsDBNull(drSet.Item("Property_Name"), "")), _
                            AltIsDBNull(drSet.Item("DocFileName"), ""), _
                            AltIsDBNull(drSet.Item("DocTrigger"), 0), _
                            AltIsDBNull(drSet.Item("NTFE_Flag"), 0), _
                            AltIsDBNull(drSet.Item("STFS_Flag"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_1"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_2"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_3"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_4"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_5"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_6"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_7"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_8"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_9"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_10"), 0), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("Active"), 1), _
                            AltIsDBNull(drSet.Item("Deleted"), 0), AltIsDBNull(drSet.Item("FinActivityTypeID"), 0) _
                    )

                    colEntities.Add(oTecDocInfo)
                End While

                Return colEntities
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function

        Public Function GetAllInfo() As MUSTER.Info.TecDocCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetTecDocuments"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@DOCUMENT_ID").Value = 0
                Params("@DocName").Value = ""

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colEntities As New MUSTER.Info.TecDocCollection
                While drSet.Read
                    Dim oTecDocInfo As New MUSTER.Info.TecDocInfo(drSet.Item("DocID"), _
                            AltIsDBNull(drSet.Item("Doc_Type"), 0), _
                            AltIsDBNull(drSet.Item("DocName"), AltIsDBNull(drSet.Item("Property_Name"), "")), _
                            AltIsDBNull(drSet.Item("DocFileName"), ""), _
                            AltIsDBNull(drSet.Item("DocTrigger"), 0), _
                            AltIsDBNull(drSet.Item("NTFE_Flag"), 0), _
                            AltIsDBNull(drSet.Item("STFS_Flag"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_1"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_2"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_3"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_4"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_5"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_6"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_7"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_8"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_9"), 0), _
                            AltIsDBNull(drSet.Item("Auto_Doc_10"), 0), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("Active"), 1), _
                            AltIsDBNull(drSet.Item("Deleted"), 0), AltIsDBNull(drSet.Item("FinActivityTypeID"), 0) _
                    )

                    colEntities.Add(oTecDocInfo)
                End While

                Return colEntities
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Sub Put(ByRef oTecDocInfo As MUSTER.Info.TecDocInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.TechnicalDocument, Integer))) Then
                    returnVal = "You do not have rights to save a Technical Document."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutTecDocument")

                With oTecDocInfo
                    If .ID = 0 Or .ID = -1 Then
                        Params(0).Value = System.DBNull.Value
                    Else
                        Params(0).Value = .ID
                    End If
                    Params(1).Value = .Active
                    Params(2).Value = .Deleted
                    Params(3).Value = .Auto_Doc_1
                    Params(4).Value = .Auto_Doc_2
                    Params(5).Value = .Auto_Doc_3
                    Params(6).Value = .Auto_Doc_4
                    Params(7).Value = .Auto_Doc_5
                    Params(8).Value = .Auto_Doc_6
                    Params(9).Value = .Auto_Doc_7
                    Params(10).Value = .Auto_Doc_8
                    Params(11).Value = .Auto_Doc_9
                    Params(12).Value = .Auto_Doc_10
                    Params(13).Value = .NTFE_Flag
                    Params(14).Value = .STFS_Flag
                    Params(15).Value = IIf(IsNothing(.DocType), 0, .DocType)
                    Params(16).Value = IIf(IsNothing(.Name), "", .Name)
                    Params(17).Value = IIf(IsNothing(.Physical_File_Name), "", .Physical_File_Name)
                    Params(18).Value = IIf(IsNothing(.Trigger_Field), 0, .Trigger_Field)

                    If .ID <= 0 Then
                        Params(19).Value = .CreatedBy
                    Else
                        Params(19).Value = .ModifiedBy
                    End If

                    Params(20).Value = IIf(.FinActivityType = 0, DBNull.Value, .FinActivityType)

                End With

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutTecDocument", Params)

                'Perform check for New ID and assign, if necessary
                If Params(0).Value <> oTecDocInfo.ID Then
                    oTecDocInfo.ID = Params(0).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function DBGetAutoCreatedDocumentParent(ByVal nEventActivityID As Integer, ByVal nAutoDocId As Integer) As Integer
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Dim nDocID As Integer
            Try

                strSQL = "spTEC_GETAUTOCREATED_DOCUMENT_PARENT"


                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@EVENT_ACTIVITY_ID").Value = nEventActivityID
                Params("@AUTO_DOC_ID").Value = nAutoDocId

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    nDocID = AltIsDBNull(drSet.Item("document_id"), 0)
                End If
                Return nDocID
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
    End Class
End Namespace
