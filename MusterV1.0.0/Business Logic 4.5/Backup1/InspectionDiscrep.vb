'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.InspectionDiscrep
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
' NOTE: This file to be used as InspectionDiscrep to build other objects.
'       Replace keyword "InspectionDiscrep" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pInspectionDiscrep
#Region "Public Events"
        Public Event evtInspectionDescepErr(ByVal MsgStr As String)
        Public Event evtInspectionDescepChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private oInspection As MUSTER.Info.InspectionInfo
        Private WithEvents oInspectionDiscrepInfo As MUSTER.Info.InspectionDiscrepInfo
        Private oInspectionDiscrepDB As New MUSTER.DataAccess.InspectionDiscrepDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
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

            oInspectionDiscrepInfo = New MUSTER.Info.InspectionDiscrepInfo
            oInspectionDiscrepDB = New MUSTER.DataAccess.InspectionDiscrepDB
        End Sub
#End Region
#Region "Exposed Attributes"
        Public ReadOnly Property ID() As Int64
            Get
                Return oInspectionDiscrepInfo.ID
            End Get
        End Property
        Public Property InspectionID() As Int64
            Get
                Return oInspectionDiscrepInfo.InspectionID
            End Get
            Set(ByVal Value As Int64)
                oInspectionDiscrepInfo.InspectionID = Value
            End Set
        End Property
        Public Property QuestionID() As Int64
            Get
                Return oInspectionDiscrepInfo.QuestionID
            End Get
            Set(ByVal Value As Int64)
                oInspectionDiscrepInfo.QuestionID = Value
            End Set
        End Property
        Public Property InspCitID() As Int64
            Get
                Return oInspectionDiscrepInfo.InspCitID
            End Get
            Set(ByVal Value As Int64)
                oInspectionDiscrepInfo.InspCitID = Value
            End Set
        End Property
        Public Property Description() As String
            Get
                Return oInspectionDiscrepInfo.Description
            End Get
            Set(ByVal Value As String)
                oInspectionDiscrepInfo.Description = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oInspectionDiscrepInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oInspectionDiscrepInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oInspectionDiscrepInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oInspectionDiscrepInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public ReadOnly Property colIsDirty() As Boolean
            Get
                Dim xInspectionDescepinfo As MUSTER.Info.InspectionDiscrepInfo
                For Each xInspectionDescepinfo In oInspection.DiscrepsCollection.Values
                    If xInspectionDescepinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
        End Property
        Public Property CreatedBy() As String
            Get
                Return oInspectionDiscrepInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oInspectionDiscrepInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oInspectionDiscrepInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oInspectionDiscrepInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oInspectionDiscrepInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oInspectionDiscrepInfo.ModifiedOn
            End Get
        End Property
        Public Property Rescinded() As Boolean
            Get
                Return oInspectionDiscrepInfo.Rescinded
            End Get
            Set(ByVal Value As Boolean)
                oInspectionDiscrepInfo.Rescinded = Value
            End Set
        End Property
        Public Property DiscrepReceived() As Date
            Get
                Return oInspectionDiscrepInfo.DiscrepReceived
            End Get
            Set(ByVal Value As Date)
                oInspectionDiscrepInfo.DiscrepReceived = Value
            End Set
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
        Public Sub Load(ByRef checklistMaster As MUSTER.Info.InspectionInfo, ByRef ds As DataSet)
            Dim dr As DataRow
            oInspection = checklistMaster
            Try
                If ds.Tables("Discrep").Rows.Count > 0 Then
                    For Each dr In ds.Tables("Discrep").Rows
                        oInspectionDiscrepInfo = New MUSTER.Info.InspectionDiscrepInfo(dr)
                        oInspection.DiscrepsCollection.Add(oInspectionDiscrepInfo)
                    Next
                End If
                ds.Tables.Remove("Discrep")
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function Retrieve(ByRef inspection As MUSTER.Info.InspectionInfo, ByVal id As Int64) As MUSTER.Info.InspectionDiscrepInfo
            Try
                oInspection = inspection
                oInspectionDiscrepInfo = oInspection.DiscrepsCollection.Item(id)
                If oInspectionDiscrepInfo Is Nothing Then
                    Add(id)
                End If
                Return oInspectionDiscrepInfo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function RetrieveByOtherID(Optional ByVal inspID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.InspectionDiscrepsCollection
            Try
                Return oInspectionDiscrepDB.DBGetByOtherID(inspID, showDeleted)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False) As Boolean
            Dim oldID As Integer
            Try
                If Not bolValidated And Not oInspectionDiscrepInfo.Deleted And Not bolDelete Then
                    If Not Me.ValidateData() Then
                        Return False
                    End If
                End If
                If Not (oInspectionDiscrepInfo.ID < 0 And oInspectionDiscrepInfo.Deleted) Then
                    oldID = oInspectionDiscrepInfo.ID
                    oInspectionDiscrepDB.Put(oInspectionDiscrepInfo, moduleID, staffID, returnVal)
                    If Not bolValidated Then
                        If oldID < 0 Then
                            oInspection.DiscrepsCollection.ChangeKey(oldID, oInspectionDiscrepInfo.ID)
                        End If
                    End If
                    oInspectionDiscrepInfo.Archive()
                    oInspectionDiscrepInfo.IsDirty = False
                End If
                If Not bolValidated And bolDelete Then
                    If oInspectionDiscrepInfo.Deleted Then
                        ' check if other inspections are present else load new instance
                        Dim strNext As String = Me.GetNext()
                        Dim strPrev As String = Me.GetPrevious()
                        If strNext = oInspectionDiscrepInfo.ID Then
                            If strPrev = oInspectionDiscrepInfo.ID Then
                                RaiseEvent evtInspectionDescepErr("Inspection " + oInspectionDiscrepInfo.ID.ToString + " deleted")
                                oInspection.DiscrepsCollection.Remove(oInspectionDiscrepInfo)
                                If bolDelete Then
                                    oInspectionDiscrepInfo = New MUSTER.Info.InspectionDiscrepInfo
                                Else
                                    oInspectionDiscrepInfo = Me.Retrieve(oInspection, 0)
                                End If
                            Else
                                RaiseEvent evtInspectionDescepErr("Inspection " + oInspectionDiscrepInfo.ID.ToString + " deleted")
                                oInspection.DiscrepsCollection.Remove(oInspectionDiscrepInfo)
                                oInspectionDiscrepInfo = Me.Retrieve(oInspection, strPrev)
                            End If
                        Else
                            RaiseEvent evtInspectionDescepErr("Inspection " + oInspectionDiscrepInfo.ID.ToString + " deleted")
                            oInspection.DiscrepsCollection.Remove(oInspectionDiscrepInfo)
                            oInspectionDiscrepInfo = Me.Retrieve(oInspection, strNext)
                        End If
                    End If
                End If
                RaiseEvent evtInspectionDescepErr(oInspectionDiscrepInfo.IsDirty)
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
                oInspectionDiscrepInfo = oInspectionDiscrepDB.DBGetByID(id)
                If oInspectionDiscrepInfo.ID = 0 Then
                    oInspectionDiscrepInfo.ID = nID
                    nID -= 1
                End If
                oInspection.DiscrepsCollection.Add(oInspectionDiscrepInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oInspectionDescep As MUSTER.Info.InspectionDiscrepInfo)
            Try
                oInspectionDiscrepInfo = oInspectionDescep
                If oInspectionDiscrepInfo.ID = 0 Then
                    oInspectionDiscrepInfo.ID = nID
                    nID -= 1
                End If
                oInspection.DiscrepsCollection.Add(oInspectionDiscrepInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal id As Int64)
            Try
                If oInspection.DiscrepsCollection.Contains(id) Then
                    oInspection.DiscrepsCollection.Remove(id)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oInspectionDiscrep As MUSTER.Info.InspectionDiscrepInfo)
            Try
                If oInspection.DiscrepsCollection.Contains(oInspectionDiscrep) Then
                    oInspection.DiscrepsCollection.Remove(oInspectionDiscrep)
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
            Dim discrepInfo As MUSTER.Info.InspectionDiscrepInfo
            Try
                For Each discrepInfo In oInspection.DiscrepsCollection.Values
                    If discrepInfo.IsDirty Or discrepInfo.ID < 0 Then
                        oInspectionDiscrepInfo = discrepInfo
                        If oInspectionDiscrepInfo.Deleted Then
                            If oInspectionDiscrepInfo.ID < 0 Then
                                delIDs.Add(oInspectionDiscrepInfo.ID)
                            Else
                                Me.Save(moduleID, staffID, returnVal, True)
                            End If
                        Else
                            If Me.ValidateData Then
                                If oInspectionDiscrepInfo.ID < 0 Then
                                    IDs.Add(oInspectionDiscrepInfo.ID)
                                End If
                                Me.Save(moduleID, staffID, returnVal, True)
                            Else : Exit For
                            End If
                        End If
                    End If
                Next
                If Not (delIDs Is Nothing) Then
                    For index = 1 To delIDs.Count
                        discrepInfo = oInspection.DiscrepsCollection.Item(CType(delIDs.Item(index), String))
                        oInspection.DiscrepsCollection.Remove(discrepInfo)
                    Next
                End If
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        discrepInfo = oInspection.DiscrepsCollection.Item(colKey)
                        oInspection.DiscrepsCollection.ChangeKey(colKey, discrepInfo.ID)
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
            Dim strArr() As String = oInspection.DiscrepsCollection.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            For Each y As String In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return oInspection.DiscrepsCollection.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            ElseIf oInspection.DiscrepsCollection.Count <> 0 Then
                Return oInspection.DiscrepsCollection.Item(nArr.GetValue(colIndex)).ID.ToString
            Else
                Return "0"
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oInspectionDiscrepInfo = New MUSTER.Info.InspectionDiscrepInfo
        End Sub
        Public Sub Reset()
            oInspectionDiscrepInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oInspectionDiscrepInfoLocal As New MUSTER.Info.InspectionDiscrepInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("InspectionDescep ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oInspectionDiscrepInfoLocal In oInspection.DiscrepsCollection.Values
                    dr = tbEntityTable.NewRow()
                    dr("ReplaceID") = oInspectionDiscrepInfoLocal.ID
                    dr("Deleted") = oInspectionDiscrepInfoLocal.Deleted
                    dr("Created By") = oInspectionDiscrepInfoLocal.CreatedBy
                    dr("Date Created") = oInspectionDiscrepInfoLocal.CreatedOn
                    dr("Last Edited By") = oInspectionDiscrepInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oInspectionDiscrepInfoLocal.ModifiedOn
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
        Private Sub InspectionDescepInfoChanged(ByVal bolValue As Boolean) Handles oInspectionDiscrepInfo.evtInspectionDiscrepInfoChanged
            RaiseEvent evtInspectionDescepChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
