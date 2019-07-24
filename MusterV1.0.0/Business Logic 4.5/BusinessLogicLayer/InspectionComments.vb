'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.InspectionComments
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MNR        06/15/05    Original class definition
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
' NOTE: This file to be used as InspectionComments to build other objects.
'       Replace keyword "InspectionComments" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pInspectionComments
#Region "Public Events"
        Public Event evtInspectionCommentsErr(ByVal MsgStr As String)
        Public Event evtInspectionCommentsChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private oInspection As MUSTER.Info.InspectionInfo
        Private WithEvents oInspectionCommentsInfo As MUSTER.Info.InspectionCommentsInfo
        Private oInspectionCommentsDB As MUSTER.DataAccess.InspectionCommentsDB
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

            oInspectionCommentsInfo = New MUSTER.Info.InspectionCommentsInfo
            oInspectionCommentsDB = New MUSTER.DataAccess.InspectionCommentsDB
        End Sub
#End Region
#Region "Exposed Attributes"
        Public ReadOnly Property ID() As Int64
            Get
                Return oInspectionCommentsInfo.ID
            End Get
        End Property
        Public Property InspectionID() As Int64
            Get
                Return oInspectionCommentsInfo.InspectionID
            End Get
            Set(ByVal Value As Int64)
                oInspectionCommentsInfo.InspectionID = Value
            End Set
        End Property
        Public Property InsComments() As String
            Get
                Return oInspectionCommentsInfo.InsComments
            End Get
            Set(ByVal Value As String)
                oInspectionCommentsInfo.InsComments = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oInspectionCommentsInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oInspectionCommentsInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oInspectionCommentsInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oInspectionCommentsInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public ReadOnly Property colIsDirty() As Boolean
            Get
                Dim xInspectionCommentsinfo As MUSTER.Info.InspectionCommentsInfo
                For Each xInspectionCommentsinfo In oInspection.InspectionCommentsCollection.Values
                    If xInspectionCommentsinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
        End Property
        Public Property CreatedBy() As String
            Get
                Return oInspectionCommentsInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oInspectionCommentsInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oInspectionCommentsInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oInspectionCommentsInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oInspectionCommentsInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oInspectionCommentsInfo.ModifiedOn
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
                If ds.Tables("Comments").Rows.Count > 0 Then
                    For Each dr In ds.Tables("Comments").Rows
                        oInspectionCommentsInfo = New MUSTER.Info.InspectionCommentsInfo(dr)
                        oInspection.InspectionCommentsCollection.Add(oInspectionCommentsInfo)
                    Next
                End If
                ds.Tables.Remove("Comments")
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function Retrieve(ByRef inspection As MUSTER.Info.InspectionInfo, ByVal id As Int64) As MUSTER.Info.InspectionCommentsInfo
            Try
                oInspection = inspection
                oInspectionCommentsInfo = oInspection.InspectionCommentsCollection.Item(id)
                If oInspectionCommentsInfo Is Nothing Then
                    Add(id)
                End If
                Return oInspectionCommentsInfo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False) As Boolean
            Dim oldID As Integer
            Try
                If Not bolValidated And Not oInspectionCommentsInfo.Deleted And Not bolDelete Then
                    If Not Me.ValidateData() Then
                        Return False
                    End If
                End If
                If Not (oInspectionCommentsInfo.ID < 0 And oInspectionCommentsInfo.Deleted) Then
                    oldID = oInspectionCommentsInfo.ID
                    oInspectionCommentsDB.Put(oInspectionCommentsInfo, moduleID, staffID, returnVal)
                    If Not bolValidated Then
                        If oldID < 0 Then
                            oInspection.InspectionCommentsCollection.ChangeKey(oldID, oInspectionCommentsInfo.ID)
                        End If
                    End If
                    oInspectionCommentsInfo.Archive()
                    oInspectionCommentsInfo.IsDirty = False
                End If
                If Not bolValidated And bolDelete Then
                    If oInspectionCommentsInfo.Deleted Then
                        ' check if other inspections are present else load new instance
                        Dim strNext As String = Me.GetNext()
                        Dim strPrev As String = Me.GetPrevious()
                        If strNext = oInspectionCommentsInfo.ID Then
                            If strPrev = oInspectionCommentsInfo.ID Then
                                RaiseEvent evtInspectionCommentsErr("Inspection " + oInspectionCommentsInfo.ID.ToString + " deleted")
                                oInspection.InspectionCommentsCollection.Remove(oInspectionCommentsInfo)
                                If bolDelete Then
                                    oInspectionCommentsInfo = New MUSTER.Info.InspectionCommentsInfo
                                Else
                                    oInspectionCommentsInfo = Me.Retrieve(oInspection, 0)
                                End If
                            Else
                                RaiseEvent evtInspectionCommentsErr("Inspection " + oInspectionCommentsInfo.ID.ToString + " deleted")
                                oInspection.InspectionCommentsCollection.Remove(oInspectionCommentsInfo)
                                oInspectionCommentsInfo = Me.Retrieve(oInspection, strPrev)
                            End If
                        Else
                            RaiseEvent evtInspectionCommentsErr("Inspection " + oInspectionCommentsInfo.ID.ToString + " deleted")
                            oInspection.InspectionCommentsCollection.Remove(oInspectionCommentsInfo)
                            oInspectionCommentsInfo = Me.Retrieve(oInspection, strNext)
                        End If
                    End If
                End If
                RaiseEvent evtInspectionCommentsErr(oInspectionCommentsInfo.IsDirty)
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
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal id As Int64)
            Try
                oInspectionCommentsInfo = oInspectionCommentsDB.DBGetByID(id)
                If oInspectionCommentsInfo.ID = 0 Then
                    oInspectionCommentsInfo.ID = nID
                    nID -= 1
                End If
                oInspection.InspectionCommentsCollection.Add(oInspectionCommentsInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oInspectionComments As MUSTER.Info.InspectionCommentsInfo)
            Try
                oInspectionCommentsInfo = oInspectionComments
                If oInspectionCommentsInfo.ID = 0 Then
                    oInspectionCommentsInfo.ID = nID
                    nID -= 1
                End If
                oInspection.InspectionCommentsCollection.Add(oInspectionCommentsInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal id As Int64)
            Try
                If oInspection.InspectionCommentsCollection.Contains(id) Then
                    oInspection.InspectionCommentsCollection.Remove(id)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oInspectionComments As MUSTER.Info.InspectionCommentsInfo)
            Try
                If oInspection.InspectionCommentsCollection.Contains(oInspectionComments) Then
                    oInspection.InspectionCommentsCollection.Remove(oInspectionComments)
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
            Dim commentsInfo As MUSTER.Info.InspectionCommentsInfo
            Try
                For Each commentsInfo In oInspection.InspectionCommentsCollection.Values
                    If commentsInfo.IsDirty Or commentsInfo.ID < 0 Then
                        oInspectionCommentsInfo = commentsInfo
                        If oInspectionCommentsInfo.Deleted Then
                            If oInspectionCommentsInfo.ID < 0 Then
                                delIDs.Add(oInspectionCommentsInfo.ID)
                            Else
                                Me.Save(moduleID, staffID, returnVal, True)
                            End If
                        Else
                            If Me.ValidateData Then
                                If oInspectionCommentsInfo.ID < 0 Then
                                    IDs.Add(oInspectionCommentsInfo.ID)
                                End If
                                Me.Save(moduleID, staffID, returnVal, True)
                            Else : Exit For
                            End If
                        End If
                    End If
                Next
                If Not (delIDs Is Nothing) Then
                    For index = 1 To delIDs.Count
                        commentsInfo = oInspection.InspectionCommentsCollection.Item(CType(delIDs.Item(index), String))
                        oInspection.InspectionCommentsCollection.Remove(commentsInfo)
                    Next
                End If
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        commentsInfo = oInspection.InspectionCommentsCollection.Item(colKey)
                        oInspection.InspectionCommentsCollection.ChangeKey(colKey, commentsInfo.ID)
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
            Dim strArr() As String = oInspection.InspectionCommentsCollection.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            For Each y As String In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return oInspection.InspectionCommentsCollection.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            ElseIf oInspection.InspectionCommentsCollection.Count <> 0 Then
                Return oInspection.InspectionCommentsCollection.Item(nArr.GetValue(colIndex)).ID.ToString
            Else
                Return "0"
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oInspectionCommentsInfo = New MUSTER.Info.InspectionCommentsInfo
        End Sub
        Public Sub Reset()
            oInspectionCommentsInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oInspectionCommentsInfoLocal As New MUSTER.Info.InspectionCommentsInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("InspectionComments ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oInspectionCommentsInfoLocal In oInspection.InspectionCommentsCollection.Values
                    dr = tbEntityTable.NewRow()
                    dr("Replace ID") = oInspectionCommentsInfoLocal.ID
                    dr("Deleted") = oInspectionCommentsInfoLocal.Deleted
                    dr("Created By") = oInspectionCommentsInfoLocal.CreatedBy
                    dr("Date Created") = oInspectionCommentsInfoLocal.CreatedOn
                    dr("Last Edited By") = oInspectionCommentsInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oInspectionCommentsInfoLocal.ModifiedOn
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
        Private Sub InspectionCommentsInfoChanged(ByVal bolValue As Boolean) Handles oInspectionCommentsInfo.evtInspectionCommentsInfoChanged
            RaiseEvent evtInspectionCommentsChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
