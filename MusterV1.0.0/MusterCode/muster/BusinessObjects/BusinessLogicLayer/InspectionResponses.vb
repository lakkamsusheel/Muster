'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.InspectionResponse
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MNR        06/11/05    Original class definition
'
' Function          Description
' GetEntity(NAME)   Returns the Entity requested by the string arg NAME
' GetEntity(ID)     Returns the Entity requested by the int arg ID
' GetAll()          Returns an ReportsCollection with all Entity objects
' Add(ID)           Adds the Entity identified by arg ID to the 
'                           internal ReportsCollection
' Add(Name)         Adds the Entity identified by arg NAME to the internal 
'                           ReportsCollection
' Add(Entity)       Adds the Entity passed as the argument to the internal 
'                           ReportsCollection
' Remove(ID)        Removes the Entity identified by arg ID from the internal 
'                           ReportsCollection
' Remove(NAME)      Removes the Entity identified by arg NAME from the 
'                           internal ReportsCollection
' EntityTable()     Returns a datatable containing all columns for the Entity 
'                           objects in the internal ReportsCollection.
'
' NOTE: This file to be used as InspectionResponse to build other objects.
'       Replace keyword "InspectionResponse" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pInspectionResponse
#Region "Public Events"
        Public Event evtInspectionResponseErr(ByVal MsgStr As String)
        Public Event evtInspectionResponseChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private oInspection As MUSTER.Info.InspectionInfo
        Private WithEvents oInspectionResponseInfo As MUSTER.Info.InspectionResponsesInfo
        Private oInspectionResponseDB As MUSTER.DataAccess.InspectionResponsesDB
        Private MusterException As MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
#End Region
#Region "Constructors"
        Public Sub New(Optional ByVal strDBConn As String = "", Optional ByRef MusterXCEP As MUSTER.Exceptions.MusterExceptions = Nothing, Optional ByRef inspection As MUSTER.Info.InspectionInfo = Nothing)
            If MusterXCEP Is Nothing Then
                MusterException = New MUSTER.Exceptions.MusterExceptions
            Else
                MusterException = MusterXCEP
            End If
            If inspection Is Nothing Then
                oInspection = New MUSTER.Info.InspectionInfo
            Else
                oInspection = inspection
            End If
            oInspectionResponseInfo = New MUSTER.Info.InspectionResponsesInfo
            oInspectionResponseDB = New MUSTER.DataAccess.InspectionResponsesDB
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int64
            Get
                Return oInspectionResponseInfo.ID
            End Get
            Set(ByVal Value As Int64)
                oInspectionResponseInfo.ID = Value
            End Set
        End Property
        Public Property InspectionID() As Int64
            Get
                Return oInspectionResponseInfo.InspectionID
            End Get
            Set(ByVal Value As Int64)
                oInspectionResponseInfo.InspectionID = Value
            End Set
        End Property
        Public Property QuestionID() As Int64
            Get
                Return oInspectionResponseInfo.QuestionID
            End Get
            Set(ByVal Value As Int64)
                oInspectionResponseInfo.QuestionID = Value
            End Set
        End Property
        Public Property SOC() As Boolean
            Get
                Return oInspectionResponseInfo.SOC
            End Get
            Set(ByVal Value As Boolean)
                oInspectionResponseInfo.SOC = Value
            End Set
        End Property
        Public Property Response() As Int64
            Get
                Return oInspectionResponseInfo.Response
            End Get
            Set(ByVal Value As Int64)
                oInspectionResponseInfo.Response = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oInspectionResponseInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oInspectionResponseInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oInspectionResponseInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oInspectionResponseInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public ReadOnly Property colIsDirty() As Boolean
            Get
                Dim xInspectionResponseinfo As MUSTER.Info.InspectionResponsesInfo
                For Each xInspectionResponseinfo In oInspection.ResponsesCollection.Values
                    If xInspectionResponseinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
        End Property
        Public Property CreatedBy() As String
            Get
                Return oInspectionResponseInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oInspectionResponseInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oInspectionResponseInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oInspectionResponseInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oInspectionResponseInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oInspectionResponseInfo.ModifiedOn
            End Get
        End Property
        Public Property InspectionInfo() As MUSTER.Info.InspectionInfo
            Get
                Return oInspection
            End Get
            Set(ByVal Value As MUSTER.Info.InspectionInfo)
                oInspection = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Sub Load(ByRef inspection As MUSTER.Info.InspectionInfo, ByRef ds As DataSet)
            Dim dr As DataRow
            oInspection = inspection
            Try
                If ds.Tables("Responses").Rows.Count > 0 Then
                    For Each dr In ds.Tables("Responses").Rows
                        oInspectionResponseInfo = New MUSTER.Info.InspectionResponsesInfo(dr)
                        oInspection.ResponsesCollection.Add(oInspectionResponseInfo)
                    Next
                End If
                ds.Tables.Remove("Responses")
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function Retrieve(ByRef inspection As MUSTER.Info.InspectionInfo, ByVal id As Int64) As MUSTER.Info.InspectionResponsesInfo
            Try
                oInspection = inspection
                oInspectionResponseInfo = oInspection.ResponsesCollection.Item(id)
                If oInspectionResponseInfo Is Nothing Then
                    Add(id)
                End If
                Return oInspectionResponseInfo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        'Public Function RetrieveByQID(ByRef inspection As MUSTER.Info.InspectionInfo, ByVal qid As Int64) As MUSTER.Info.InspectionResponsesInfo
        '    Try
        '        oInspection = inspection
        '        ' declared as new instance cause, if not found in collection
        '        ' need to return a new instance
        '        oInspectionResponseInfo = New MUSTER.Info.InspectionResponsesInfo
        '        For Each resp As MUSTER.Info.InspectionResponsesInfo In oInspection.ResponsesCollection.Values
        '            If resp.QuestionID = qid Then
        '                oInspectionResponseInfo = resp
        '                Exit For
        '            End If
        '        Next
        '        Return oInspectionResponseInfo
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False) As Boolean
            Dim oldID As Integer
            Try
                If Not bolValidated And Not oInspectionResponseInfo.Deleted And Not bolDelete Then
                    If Not Me.ValidateData() Then
                        Return False
                    End If
                End If
                If Not (oInspectionResponseInfo.ID < 0 And oInspectionResponseInfo.Deleted) Then
                    oldID = oInspectionResponseInfo.ID
                    oInspectionResponseDB.Put(oInspectionResponseInfo, moduleID, staffID, returnVal)
                    If Not bolValidated Then
                        If oldID < 0 Then
                            oInspection.ResponsesCollection.ChangeKey(oldID, oInspectionResponseInfo.ID)
                        End If
                    End If
                    oInspectionResponseInfo.Archive()
                    oInspectionResponseInfo.IsDirty = False
                End If
                If Not bolValidated And bolDelete Then
                    If oInspectionResponseInfo.Deleted Then
                        ' check if other inspections are present else load new instance
                        Dim strNext As String = Me.GetNext()
                        Dim strPrev As String = Me.GetPrevious()
                        If strNext = oInspectionResponseInfo.ID Then
                            If strPrev = oInspectionResponseInfo.ID Then
                                RaiseEvent evtInspectionResponseErr("Inspection " + oInspectionResponseInfo.ID.ToString + " deleted")
                                oInspection.ResponsesCollection.Remove(oInspectionResponseInfo)
                                If bolDelete Then
                                    oInspectionResponseInfo = New MUSTER.Info.InspectionResponsesInfo
                                Else
                                    oInspectionResponseInfo = Me.Retrieve(oInspection, 0)
                                End If
                            Else
                                RaiseEvent evtInspectionResponseErr("Inspection " + oInspectionResponseInfo.ID.ToString + " deleted")
                                oInspection.ResponsesCollection.Remove(oInspectionResponseInfo)
                                oInspectionResponseInfo = Me.Retrieve(oInspection, strPrev)
                            End If
                        Else
                            RaiseEvent evtInspectionResponseErr("Inspection " + oInspectionResponseInfo.ID.ToString + " deleted")
                            oInspection.ResponsesCollection.Remove(oInspectionResponseInfo)
                            oInspectionResponseInfo = Me.Retrieve(oInspection, strNext)
                        End If
                    End If
                End If
                RaiseEvent evtInspectionResponseErr(oInspectionResponseInfo.IsDirty)
                Return True
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
                Return False
            End Try
        End Function
        Public Function ValidateData() As Boolean
            Return True
        End Function
#End Region
#Region "Collection Operations"
        Public Sub Add(ByVal id As Int64)
            Try
                oInspectionResponseInfo = oInspectionResponseDB.DBGetByID(id)
                If oInspectionResponseInfo.ID = 0 Then
                    oInspectionResponseInfo.ID = nID
                    nID -= 1
                End If
                oInspection.ResponsesCollection.Add(oInspectionResponseInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oInspectionResponse As MUSTER.Info.InspectionResponsesInfo)
            Try
                oInspectionResponseInfo = oInspectionResponse
                If oInspectionResponseInfo.ID = 0 Then
                    oInspectionResponseInfo.ID = nID
                    nID -= 1
                End If
                oInspection.ResponsesCollection.Add(oInspectionResponseInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal id As Int64)
            Try
                If oInspection.ResponsesCollection.Contains(id) Then
                    oInspection.ResponsesCollection.Remove(id)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oInspectionResponse As MUSTER.Info.InspectionResponsesInfo)
            Try
                If oInspection.ResponsesCollection.Contains(oInspectionResponse) Then
                    oInspection.ResponsesCollection.Remove(oInspectionResponse)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim IDs As New Collection
            Dim delIDs As New Collection
            Dim index As Integer
            Dim respInfo As MUSTER.Info.InspectionResponsesInfo
            Try
                For Each respInfo In oInspection.ResponsesCollection.Values
                    If respInfo.IsDirty Or respInfo.ID < 0 Then
                        oInspectionResponseInfo = respInfo
                        If oInspectionResponseInfo.Deleted Then
                            If oInspectionResponseInfo.ID < 0 Then
                                delIDs.Add(oInspectionResponseInfo.ID)
                            Else
                                Me.Save(moduleID, staffID, returnVal, True, True)
                            End If
                        Else
                            If Me.ValidateData Then
                                If oInspectionResponseInfo.ID < 0 Then
                                    IDs.Add(oInspectionResponseInfo.ID)
                                End If
                                Me.Save(moduleID, staffID, returnVal, True)
                            Else : Exit For
                            End If
                        End If
                    End If
                Next
                If Not (delIDs Is Nothing) Then
                    For index = 1 To delIDs.Count
                        respInfo = oInspection.ResponsesCollection.Item(CType(delIDs.Item(index), String))
                        oInspection.ResponsesCollection.Remove(respInfo)
                    Next
                End If
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        respInfo = oInspection.ResponsesCollection.Item(colKey)
                        oInspection.ResponsesCollection.ChangeKey(colKey, respInfo.ID)
                    Next
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = oInspection.ResponsesCollection.GetKeys()
            Dim y As String
            strArr.Sort(strArr)
            colIndex = Array.BinarySearch(strArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= strArr.GetUpperBound(0) Then
                Return oInspection.ResponsesCollection.Item(strArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return oInspection.ResponsesCollection.Item(strArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oInspectionResponseInfo = New MUSTER.Info.InspectionResponsesInfo
        End Sub
        Public Sub Reset()
            oInspectionResponseInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oInspectionResponseInfoLocal As New MUSTER.Info.InspectionResponsesInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("InspectionResponse ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oInspectionResponseInfoLocal In oInspection.ResponsesCollection.Values
                    dr = tbEntityTable.NewRow()
                    dr("Replace ID") = oInspectionResponseInfoLocal.ID
                    dr("Deleted") = oInspectionResponseInfoLocal.Deleted
                    dr("Created By") = oInspectionResponseInfoLocal.CreatedBy
                    dr("Date Created") = oInspectionResponseInfoLocal.CreatedOn
                    dr("Last Edited By") = oInspectionResponseInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oInspectionResponseInfoLocal.ModifiedOn
                    tbEntityTable.Rows.Add(dr)
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub InspectionResponseInfoChanged(ByVal bolValue As Boolean) Handles oInspectionResponseInfo.evtInspectionResponsesInfoChanged
            RaiseEvent evtInspectionResponseChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
