'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.InspectionCitation
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
' NOTE: This file to be used as InspectionCitation to build other objects.
'       Replace keyword "InspectionCitation" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pInspectionCitation
#Region "Public Events"
        Public Event evtInspectionCitationErr(ByVal MsgStr As String)
        Public Event evtInspectionCitationChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private oInspection As MUSTER.Info.InspectionInfo
        Private WithEvents oInspectionCitationInfo As MUSTER.Info.InspectionCitationInfo
        Private oInspectionCitationDB As MUSTER.DataAccess.InspectionCitationDB
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
            oInspectionCitationInfo = New MUSTER.Info.InspectionCitationInfo
            oInspectionCitationDB = New MUSTER.DataAccess.InspectionCitationDB
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property OCEID() As Int64
            Get
                Return oInspectionCitationInfo.oceid
            End Get
            Set(ByVal Value As Int64)
                oInspectionCitationInfo.oceid = Value
            End Set
        End Property
        Public ReadOnly Property ID() As Int64
            Get
                Return oInspectionCitationInfo.ID
            End Get
        End Property
        Public Property InspectionID() As Int64
            Get
                Return oInspectionCitationInfo.InspectionID
            End Get
            Set(ByVal Value As Int64)
                oInspectionCitationInfo.InspectionID = Value
            End Set
        End Property
        Public Property QuestionID() As Int64
            Get
                Return oInspectionCitationInfo.QuestionID
            End Get
            Set(ByVal Value As Int64)
                oInspectionCitationInfo.QuestionID = Value
            End Set
        End Property
        Public Property FacilityID() As Int64
            Get
                Return oInspectionCitationInfo.FacilityID
            End Get
            Set(ByVal Value As Int64)
                oInspectionCitationInfo.FacilityID = Value
            End Set
        End Property
        Public Property FCEID() As Int64
            Get
                Return oInspectionCitationInfo.FCEID
            End Get
            Set(ByVal Value As Int64)
                oInspectionCitationInfo.FCEID = Value
            End Set
        End Property
        Public Property CitationID() As Int64
            Get
                Return oInspectionCitationInfo.CitationID
            End Get
            Set(ByVal Value As Int64)
                oInspectionCitationInfo.CitationID = Value
            End Set
        End Property
        Public Property CCAT() As String
            Get
                Return oInspectionCitationInfo.CCAT
            End Get
            Set(ByVal Value As String)
                oInspectionCitationInfo.CCAT = Value
            End Set
        End Property
        Public Property Rescinded() As Boolean
            Get
                Return oInspectionCitationInfo.Rescinded
            End Get
            Set(ByVal Value As Boolean)
                oInspectionCitationInfo.Rescinded = Value
            End Set
        End Property
        Public Property CitationDueDate() As Date
            Get
                Return oInspectionCitationInfo.CitationDueDate
            End Get
            Set(ByVal Value As Date)
                oInspectionCitationInfo.CitationDueDate = Value
            End Set
        End Property
        Public Property CitationReceivedDate() As Date
            Get
                Return oInspectionCitationInfo.CitationReceivedDate
            End Get
            Set(ByVal Value As Date)
                oInspectionCitationInfo.CitationReceivedDate = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oInspectionCitationInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oInspectionCitationInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oInspectionCitationInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oInspectionCitationInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public ReadOnly Property colIsDirty() As Boolean
            Get
                Dim xInspectionCitationinfo As MUSTER.Info.InspectionCitationInfo
                For Each xInspectionCitationinfo In oInspection.CitationsCollection.Values
                    If xInspectionCitationinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
        End Property
        Public Property CreatedBy() As String
            Get
                Return oInspectionCitationInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oInspectionCitationInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oInspectionCitationInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oInspectionCitationInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oInspectionCitationInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oInspectionCitationInfo.ModifiedOn
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
        Public Property InspectionCitationInfo() As MUSTER.Info.InspectionCitationInfo
            Get
                Return oInspectionCitationInfo
            End Get
            Set(ByVal Value As MUSTER.Info.InspectionCitationInfo)
                oInspectionCitationInfo = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Sub Load(ByRef checklistMaster As MUSTER.Info.InspectionInfo, ByRef ds As DataSet)
            Dim dr As DataRow
            oInspection = checklistMaster
            Try
                If ds.Tables("Citation").Rows.Count > 0 Then
                    For Each dr In ds.Tables("Citation").Rows
                        oInspectionCitationInfo = New MUSTER.Info.InspectionCitationInfo(dr)
                        oInspection.CitationsCollection.Add(oInspectionCitationInfo)
                    Next
                End If
                ds.Tables.Remove("Citation")
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function Retrieve(ByRef inspection As MUSTER.Info.InspectionInfo, ByVal id As Int64) As MUSTER.Info.InspectionCitationInfo
            Try
                oInspection = inspection
                oInspectionCitationInfo = oInspection.CitationsCollection.Item(id)
                If oInspectionCitationInfo Is Nothing Then
                    Add(id)
                End If
                Return oInspectionCitationInfo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function RetrieveByOtherID(Optional ByVal inspID As Integer = 0, Optional ByVal fceID As Integer = 0, Optional ByVal oceID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.InspectionCitationsCollection
            Try
                Return oInspectionCitationDB.DBGetByOtherID(inspID, fceID, oceID, showDeleted)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False, Optional ByVal OverrideRights As Boolean = False) As Boolean
            Dim oldID As Integer
            Try
                If Not bolValidated And Not oInspectionCitationInfo.Deleted And Not bolDelete Then
                    If Not Me.ValidateData() Then
                        Return False
                    End If
                End If
                If Not (oInspectionCitationInfo.ID < 0 And oInspectionCitationInfo.Deleted) Then
                    oldID = oInspectionCitationInfo.ID
                    oInspectionCitationDB.Put(oInspectionCitationInfo, moduleID, staffID, returnVal, OverrideRights)
                    If Not bolValidated Then
                        If oldID < 0 Then
                            oInspection.CitationsCollection.ChangeKey(oldID, oInspectionCitationInfo.ID)
                        End If
                    End If
                    ' change the citation id in discrep
                    If oldID <> oInspectionCitationInfo.ID Then
                        For Each discrep As MUSTER.Info.InspectionDiscrepInfo In oInspection.DiscrepsCollection.Values
                            If discrep.InspCitID = oldID Then
                                discrep.InspCitID = oInspectionCitationInfo.ID
                            End If
                        Next
                    End If
                    oInspectionCitationInfo.Archive()
                    oInspectionCitationInfo.IsDirty = False
                End If
                If Not bolValidated And bolDelete Then
                    If oInspectionCitationInfo.Deleted Then
                        ' check if other inspections are present else load new instance
                        Dim strNext As String = Me.GetNext()
                        Dim strPrev As String = Me.GetPrevious()
                        If strNext = oInspectionCitationInfo.ID Then
                            If strPrev = oInspectionCitationInfo.ID Then
                                RaiseEvent evtInspectionCitationErr("Inspection " + oInspectionCitationInfo.ID.ToString + " deleted")
                                oInspection.CitationsCollection.Remove(oInspectionCitationInfo)
                                If bolDelete Then
                                    oInspectionCitationInfo = New MUSTER.Info.InspectionCitationInfo
                                Else
                                    oInspectionCitationInfo = Me.Retrieve(oInspection, 0)
                                End If
                            Else
                                RaiseEvent evtInspectionCitationErr("Inspection " + oInspectionCitationInfo.ID.ToString + " deleted")
                                oInspection.CitationsCollection.Remove(oInspectionCitationInfo)
                                oInspectionCitationInfo = Me.Retrieve(oInspection, strPrev)
                            End If
                        Else
                            RaiseEvent evtInspectionCitationErr("Inspection " + oInspectionCitationInfo.ID.ToString + " deleted")
                            oInspection.CitationsCollection.Remove(oInspectionCitationInfo)
                            oInspectionCitationInfo = Me.Retrieve(oInspection, strNext)
                        End If
                    End If
                End If
                RaiseEvent evtInspectionCitationErr(oInspectionCitationInfo.IsDirty)
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
        Public Function CheckCitationExists(ByVal onDate As Date, Optional ByVal facID As Integer = 0, Optional ByVal showDeleted As Boolean = False, Optional ByVal citationID As Integer = 0, Optional ByVal fceCreated As Int16 = -1, Optional ByVal oceCreated As Int16 = -1, Optional ByVal rescinded As Int16 = -1, Optional ByVal strExcludeOCE As String = "") As Boolean
            Try
                Return oInspectionCitationDB.DBCheckCitationExists(onDate, facID, showDeleted, citationID, fceCreated, oceCreated, rescinded, strExcludeOCE)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal id As Int64)
            Try
                oInspectionCitationInfo = oInspectionCitationDB.DBGetByID(id)
                If oInspectionCitationInfo.ID = 0 Then
                    oInspectionCitationInfo.ID = nID
                    nID -= 1
                End If
                oInspection.CitationsCollection.Add(oInspectionCitationInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oInspectionCitation As MUSTER.Info.InspectionCitationInfo)
            Try
                oInspectionCitationInfo = oInspectionCitation
                If oInspectionCitationInfo.ID = 0 Then
                    oInspectionCitationInfo.ID = nID
                    nID -= 1
                End If
                oInspection.CitationsCollection.Add(oInspectionCitationInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal id As Int64)
            Try
                If oInspection.CitationsCollection.Contains(id) Then
                    oInspection.CitationsCollection.Remove(id)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oInspectionCitation As MUSTER.Info.InspectionCitationInfo)
            Try
                If oInspection.CitationsCollection.Contains(oInspectionCitation) Then
                    oInspection.CitationsCollection.Remove(oInspectionCitation)
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
            Dim citationInfo As MUSTER.Info.InspectionCitationInfo
            Try
                For Each citationInfo In oInspection.CitationsCollection.Values
                    If citationInfo.IsDirty Or citationInfo.ID < 0 Then
                        oInspectionCitationInfo = citationInfo
                        If oInspectionCitationInfo.Deleted Then
                            If oInspectionCitationInfo.ID < 0 Then
                                delIDs.Add(oInspectionCitationInfo.ID)
                            Else
                                Me.Save(moduleID, staffID, returnVal, True)
                            End If
                        Else
                            If Me.ValidateData Then
                                If oInspectionCitationInfo.ID < 0 Then
                                    IDs.Add(oInspectionCitationInfo.ID)
                                End If
                                Me.Save(moduleID, staffID, returnVal, True)
                            Else : Exit For
                            End If
                        End If
                    End If
                Next
                If Not (delIDs Is Nothing) Then
                    For index = 1 To delIDs.Count
                        citationInfo = oInspection.CitationsCollection.Item(CType(delIDs.Item(index), String))
                        oInspection.CitationsCollection.Remove(citationInfo)
                    Next
                End If
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        citationInfo = oInspection.CitationsCollection.Item(colKey)
                        oInspection.CitationsCollection.ChangeKey(colKey, citationInfo.ID)
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
            Dim strArr() As String = oInspection.CitationsCollection.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            For Each y As String In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return oInspection.CitationsCollection.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            ElseIf oInspection.CitationsCollection.Count <> 0 Then
                Return oInspection.CitationsCollection.Item(nArr.GetValue(colIndex)).ID.ToString
            Else
                Return "0"
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oInspectionCitationInfo = New MUSTER.Info.InspectionCitationInfo
        End Sub
        Public Sub Reset()
            oInspectionCitationInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oInspectionCitationInfoLocal As New MUSTER.Info.InspectionCitationInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("InspectionCitation ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oInspectionCitationInfoLocal In oInspection.CitationsCollection.Values
                    dr = tbEntityTable.NewRow()
                    dr("Replace ID") = oInspectionCitationInfoLocal.ID
                    dr("Deleted") = oInspectionCitationInfoLocal.Deleted
                    dr("Created By") = oInspectionCitationInfoLocal.CreatedBy
                    dr("Date Created") = oInspectionCitationInfoLocal.CreatedOn
                    dr("Last Edited By") = oInspectionCitationInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oInspectionCitationInfoLocal.ModifiedOn
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
        Private Sub InspectionCitationInfoChanged(ByVal bolValue As Boolean) Handles oInspectionCitationInfo.evtInspectionCitationInfoChanged
            RaiseEvent evtInspectionCitationChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
