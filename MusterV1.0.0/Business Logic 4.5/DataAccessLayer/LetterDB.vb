'-------------------------------------------------------------------------------
' MUSTER.DataAccess.LetterDB
'   Provides the means for marshalling Address to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       EN   12/13/04    Original class definition.
'  1.1       AN   12/30/04    Added Try catch and Exception Handling/Logging
'  1.2       EN   02/10/05    Modified 01/01/1901  to 01/01/0001
'  1.3       AB   02/15/05    Replaced dynamic SQL with stored procedures in the following
'                                   Functions:  GetAllInfo, DBGetByID
'  1.4       AB   02/15/05    Added Finally to the Try/Catch to close all datareaders
'  1.5       AB   02/16/05    Removed any IsNull calls for fields the DB requires
'  1.6       AB   02/18/05    Set all parameters for SP, that are not required, to NULL
'  1.7       MR   03/27/05    Modified put and Get functions to pass Modified On and Created On values.
'
' Function                  Description
' GetAllInfo()        Returns an LetterCollection containing all letter objects in the repository.
' DBGetByName(NAME)   Returns an letterinfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an letterinfo object indicated by arg ID.
' DBGetDS(SQL)        Returns a resultant Dataset by running query specified by the string arg SQL
' Put(oletter)        Saves the letter passed as an argument, to the DB
''-------------------------------------------------------------------------------

Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class LetterDB
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
        Public Function GetAllInfo(Optional ByVal strUserID As String = "", Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.LetterCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            'strSQL = "SELECT * FROM tblSYS_DOCUMENT_MANAGER where 1=1"
            'strSQL += IIf(Not showDeleted, " AND DELETED <> 1", "")

            Try
                strSQL = "spGetDocumentManager"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Document_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)
                Params("@UserID").Value = IIf(strUserID = String.Empty, DBNull.Value, strUserID)


                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colletter As New MUSTER.Info.LetterCollection
                Dim strID As String
                While drSet.Read
                    Dim oLetterinfo As New MUSTER.Info.LetterInfo(drSet.Item("DOCUMENT_ID"), _
                                          drSet.Item("DOCUMENT_NAME"), _
                                          drSet.Item("TYPE_OF_DOCUMENT"), _
                                          drSet.Item("DOCUMENT_LOCATION"), _
                                          drSet.Item("ENTITY_TYPE"), _
                                          drSet.Item("ENTITY_ID"), _
                                          drSet.Item("DOCUMENT_DESCRIPTION"), _
                                          drSet.Item("WORKFLOW"), _
                                          AltIsDBNull(drSet.Item("DATE_PRINTED"), CDate("01/01/0001")), _
                                          drSet.Item("DELETED"), _
                                          drSet.Item("CREATED_BY"), _
                                          drSet.Item("DATE_EDITED"), _
                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                          AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                          AltIsDBNull(drSet.Item("OWNING_USER"), String.Empty), _
                                          AltIsDBNull(drSet.Item("MODULE_ID"), 0), _
                                          AltIsDBNull(drSet.Item("EVENT_ID"), 0), _
                                          AltIsDBNull(drSet.Item("EVENT_SEQUENCE"), 0), _
                                          AltIsDBNull(drSet.Item("EVENT_TYPE"), 0))
                    colletter.Add(oLetterinfo)
                End While
                Return colletter
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByID(ByVal nVal As Int64) As MUSTER.Info.LetterInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetDocumentManager"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@OrderBy").Value = 1
                Params("@Document_ID").Value = nVal
                Params("@Deleted").Value = False

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.LetterInfo(drSet.Item("DOCUMENT_ID"), _
                                          drSet.Item("DOCUMENT_NAME"), _
                                          drSet.Item("TYPE_OF_DOCUMENT"), _
                                          drSet.Item("DOCUMENT_LOCATION"), _
                                          drSet.Item("ENTITY_TYPE"), _
                                          drSet.Item("ENTITY_ID"), _
                                          drSet.Item("DOCUMENT_DESCRIPTION"), _
                                          drSet.Item("WORKFLOW"), _
                                          AltIsDBNull(drSet.Item("DATE_PRINTED"), CDate("01/01/0001")), _
                                          drSet.Item("DELETED"), _
                                          AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                          AltIsDBNull(drSet.Item("DATE_EDITED"), CDate("01/01/0001")), _
                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                          AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                          AltIsDBNull(drSet.Item("OWNING_USER"), String.Empty), _
                                          AltIsDBNull(drSet.Item("MODULE_ID"), 0), _
                                          AltIsDBNull(drSet.Item("EVENT_ID"), 0), _
                                          AltIsDBNull(drSet.Item("EVENT_SEQUENCE"), 0), _
                                          AltIsDBNull(drSet.Item("EVENT_TYPE"), 0))
                Else
                    Return New MUSTER.Info.LetterInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByDocName(ByVal strDocName As String, ByVal strOwningUser As String, Optional ByVal bolDeleted As Boolean = False) As MUSTER.Info.LetterInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetDocumentManagerByDocName"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@DOCUMENT_NAME").Value = strDocName
                Params("@OWNING_USER").Value = strOwningUser
                Params("@DELETED").Value = bolDeleted

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.LetterInfo(drSet.Item("DOCUMENT_ID"), _
                                          drSet.Item("DOCUMENT_NAME"), _
                                          drSet.Item("TYPE_OF_DOCUMENT"), _
                                          drSet.Item("DOCUMENT_LOCATION"), _
                                          drSet.Item("ENTITY_TYPE"), _
                                          drSet.Item("ENTITY_ID"), _
                                          drSet.Item("DOCUMENT_DESCRIPTION"), _
                                          drSet.Item("WORKFLOW"), _
                                          AltIsDBNull(drSet.Item("DATE_PRINTED"), CDate("01/01/0001")), _
                                          drSet.Item("DELETED"), _
                                          AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                          AltIsDBNull(drSet.Item("DATE_EDITED"), CDate("01/01/0001")), _
                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                          AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                          AltIsDBNull(drSet.Item("OWNING_USER"), String.Empty), _
                                          AltIsDBNull(drSet.Item("MODULE_ID"), 0), _
                                          AltIsDBNull(drSet.Item("EVENT_ID"), 0), _
                                          AltIsDBNull(drSet.Item("EVENT_SEQUENCE"), 0), _
                                          AltIsDBNull(drSet.Item("EVENT_TYPE"), 0))
                Else
                    Return New MUSTER.Info.LetterInfo
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
        Public Function Put(ByRef oletter As MUSTER.Info.LetterInfo, Optional ByVal Type As String = "SYSTEM") As Integer
            Try
                Dim Params() As SqlParameter
                Dim bolNewFlag As Boolean
                Dim strSQL As String = String.Empty
                If Type = "SYSTEM" Then
                    strSQL = "spPutSYSDOCUMENTMANAGER"
                Else
                    strSQL = "spPutSYSMANUALDOCUMENT"     'MANUAL
                End If
                bolNewFlag = False
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = oletter.ID
                If oletter.ID <= 0 Then
                    Params(0).Value = 0
                    bolNewFlag = True
                Else
                    Params(0).Value = oletter.ID
                End If
                Params(0).Direction = ParameterDirection.InputOutput
                Params(1).Value = oletter.Name
                Params(2).Value = oletter.TypeofDocument
                Params(3).Value = IsNull(oletter.DocumentLocation, 0)
                Params(4).Value = IsNull(oletter.EntityType, 0)
                Params(5).Value = IsNull(oletter.EntityId, 0)
                Params(6).Value = IsNull(oletter.DocumentDescription, String.Empty)
                Params(7).Value = oletter.WorkFlow
                Params(8).Value = IIf(oletter.DatePrinted = CDate("01/01/0001"), System.DBNull.Value, oletter.DatePrinted)
                Params(9).Value = oletter.Deleted
                Params(10).Value = System.DBNull.Value
                If Type = "SYSTEM" Then
                    Params(11).Value = DBNull.Value
                    Params(12).Value = DBNull.Value
                    Params(13).Value = DBNull.Value
                Else
                    If Date.Compare(oletter.ModifiedOn, CDate("01/01/0001")) = 0 Then
                        Params(11).Value = DBNull.Value
                    Else
                        Params(11).Value = oletter.ModifiedOn
                    End If
                    Params(12).Value = System.DBNull.Value
                    If Date.Compare(oletter.CreatedOn, CDate("01/01/0001")) = 0 Then
                        Params(13).Value = Now
                    Else
                        Params(13).Value = oletter.CreatedOn
                    End If
                End If
                Params(14).Value = oletter.OwningUser
                Params(15).Value = oletter.ModuleID
                Params(16).Value = oletter.EventID
                Params(17).Value = oletter.EventSequence
                Params(18).Value = oletter.EventType

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If bolNewFlag And Params(0).Value <> 0 Then
                    oletter.ID = Params(0).Value
                End If
                oletter.ModifiedBy = AltIsDBNull(Params(10).Value, String.Empty)
                oletter.ModifiedOn = AltIsDBNull(Params(11).Value, CDate("01/01/0001"))
                oletter.CreatedBy = AltIsDBNull(Params(12).Value, String.Empty)
                oletter.CreatedOn = AltIsDBNull(Params(13).Value, CDate("01/01/0001"))
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetDocumentsList(ByVal strUserID As String, Optional ByVal PrintedFlag As Boolean = False) As DataSet
            Dim strSQL As String
            Dim Params As Collection
            Dim dsData As DataSet

            Try
                strSQL = "spGetDocumentsList"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@PrintedFlag").Value = PrintedFlag
                Params("@UserID").Value = IIf(strUserID = String.Empty, DBNull.Value, strUserID)
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function UpdatePrintedStatus(ByVal DocumentID As Integer, ByVal DocumentLocation As String, ByVal DatePrinted As DateTime)
            Dim Params() As SqlParameter
            Try

                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutDocumentManager_Printed")
                Params(0).Value = DocumentID
                Params(1).Value = DocumentLocation
                Params(2).Value = IIf(DatePrinted = CDate("01/01/0001"), System.DBNull.Value, DatePrinted)

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutDocumentManager_Printed", Params)

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        Public Function GetCalendarYear(ByVal strUserID As String, Optional ByVal PrintedFlag As Integer = 0)
            Dim strSQL As String
            Dim Params As Collection
            Dim dsData As DataSet

            Try
                strSQL = "spGetDocument_CalendarYear"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@PrintedFlag").Value = PrintedFlag
                Params("@UserID").Value = IIf(strUserID = String.Empty, DBNull.Value, strUserID)
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function DeleteManualDocuments(ByVal UserID As String)
            Dim Params() As SqlParameter
            Try

                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spDeleteManualDocuments")
                Params(0).Value = IIf(UserID = String.Empty, System.DBNull.Value, UserID)
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spDeleteManualDocuments", Params)

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetManualAndSystemDocuments(ByVal strUserID As String) As DataSet
            Dim strSQL As String
            Dim Params As Collection
            Dim dsData As DataSet

            Try
                strSQL = "spGetManualAndSystemDocuments"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@UserID").Value = IIf(strUserID = String.Empty, DBNull.Value, strUserID)
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetManualDocsWithDesc(ByVal strUserID As String) As DataSet
            Dim strSQL As String
            Dim Params As Collection
            Dim dsData As DataSet

            Try
                strSQL = "spGetManualDocsWithDesc"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@UserID").Value = IIf(strUserID = String.Empty, DBNull.Value, strUserID)
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub DBSaveDocDescription(ByVal ndocID As Integer, ByVal strdocDesc As String, ByVal isManualDoc As Boolean, ByVal staffID As Integer)
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try
                strSQL = "spPutDocDesc"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = ndocID
                Params(1).Value = strdocDesc
                Params(2).Value = isManualDoc
                Params(3).Value = staffID

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region

    End Class
End Namespace
