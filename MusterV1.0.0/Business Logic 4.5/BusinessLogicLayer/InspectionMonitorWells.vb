'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.InspectionMonitorWells
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      MNR         6/15/05     Original class definition
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
' NOTE: This file to be used as InspectionMonitorWells to build other objects.
'       Replace keyword "InspectionMonitorWells" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pInspectionMonitorWells
#Region "Public Events"
        Public Event evtInspectionMonitorWellsErr(ByVal MsgStr As String)
        Public Event evtInspectionMonitorWellsChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private oInspection As MUSTER.Info.InspectionInfo
        Private WithEvents oInspectionMonitorWellsInfo As MUSTER.Info.InspectionMonitorWellsInfo
        Private oInspectionMonitorWellsDB As MUSTER.DataAccess.InspectionMonitorWellsDB
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
            oInspectionMonitorWellsInfo = New MUSTER.Info.InspectionMonitorWellsInfo
            oInspectionMonitorWellsDB = New MUSTER.DataAccess.InspectionMonitorWellsDB
        End Sub
#End Region
#Region "Exposed Attributes"
        Public ReadOnly Property ID() As Int64
            Get
                Return oInspectionMonitorWellsInfo.ID
            End Get
        End Property
        Public Property InspectionID() As Int64
            Get
                Return oInspectionMonitorWellsInfo.InspectionID
            End Get
            Set(ByVal Value As Int64)
                oInspectionMonitorWellsInfo.InspectionID = Value
            End Set
        End Property
        Public Property QuestionID() As Int64
            Get
                Return oInspectionMonitorWellsInfo.QuestionID
            End Get
            Set(ByVal Value As Int64)
                oInspectionMonitorWellsInfo.QuestionID = Value
            End Set
        End Property
        Public Property LineNumber() As Int64
            Get
                Return oInspectionMonitorWellsInfo.LineNumber
            End Get
            Set(ByVal Value As Int64)
                oInspectionMonitorWellsInfo.LineNumber = Value
            End Set
        End Property
        Public Property TankLine() As Boolean
            Get
                Return oInspectionMonitorWellsInfo.TankLine
            End Get
            Set(ByVal Value As Boolean)
                oInspectionMonitorWellsInfo.TankLine = Value
            End Set
        End Property
        Public Property WellNumber() As Int64
            Get
                Return oInspectionMonitorWellsInfo.WellNumber
            End Get
            Set(ByVal Value As Int64)
                oInspectionMonitorWellsInfo.WellNumber = Value
            End Set
        End Property
        Public Property WellDepth() As String
            Get
                Return oInspectionMonitorWellsInfo.WellDepth
            End Get
            Set(ByVal Value As String)
                oInspectionMonitorWellsInfo.WellDepth = Value
            End Set
        End Property
        Public Property DepthToWater() As String
            Get
                Return oInspectionMonitorWellsInfo.DepthToWater
            End Get
            Set(ByVal Value As String)
                oInspectionMonitorWellsInfo.DepthToWater = Value
            End Set
        End Property
        Public Property DepthToSlots() As String
            Get
                Return oInspectionMonitorWellsInfo.DepthToSlots
            End Get
            Set(ByVal Value As String)
                oInspectionMonitorWellsInfo.DepthToSlots = Value
            End Set
        End Property
        Public Property SurfaceSealed() As Int64
            Get
                Return oInspectionMonitorWellsInfo.SurfaceSealed
            End Get
            Set(ByVal Value As Int64)
                oInspectionMonitorWellsInfo.SurfaceSealed = Value
            End Set
        End Property
        Public Property WellCaps() As Int64
            Get
                Return oInspectionMonitorWellsInfo.WellCaps
            End Get
            Set(ByVal Value As Int64)
                oInspectionMonitorWellsInfo.WellCaps = Value
            End Set
        End Property
        Public Property InspectorsObservations() As String
            Get
                Return oInspectionMonitorWellsInfo.InspectorsObservations
            End Get
            Set(ByVal Value As String)
                oInspectionMonitorWellsInfo.InspectorsObservations = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oInspectionMonitorWellsInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oInspectionMonitorWellsInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oInspectionMonitorWellsInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oInspectionMonitorWellsInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public ReadOnly Property colIsDirty() As Boolean
            Get
                Dim xInspectionMonitorWellsinfo As MUSTER.Info.InspectionMonitorWellsInfo
                For Each xInspectionMonitorWellsinfo In oInspection.MonitorWellsCollection.Values
                    If xInspectionMonitorWellsinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
        End Property
        Public Property CreatedBy() As String
            Get
                Return oInspectionMonitorWellsInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oInspectionMonitorWellsInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oInspectionMonitorWellsInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oInspectionMonitorWellsInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oInspectionMonitorWellsInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oInspectionMonitorWellsInfo.ModifiedOn
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
        Public Sub Load(ByRef checkListMaster As MUSTER.Info.InspectionInfo, ByRef ds As DataSet)
            Dim dr As DataRow
            oInspection = checkListMaster
            Try
                If ds.Tables("MonitorWells").Rows.Count > 0 Then
                    For Each dr In ds.Tables("MonitorWells").Rows
                        oInspectionMonitorWellsInfo = New MUSTER.Info.InspectionMonitorWellsInfo(dr)
                        oInspection.MonitorWellsCollection.Add(oInspectionMonitorWellsInfo)
                    Next
                End If
                ds.Tables.Remove("MonitorWells")
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function Retrieve(ByRef inspection As MUSTER.Info.InspectionInfo, ByVal id As Int64) As MUSTER.Info.InspectionMonitorWellsInfo
            Try
                oInspection = inspection
                oInspectionMonitorWellsInfo = oInspection.MonitorWellsCollection.Item(id)
                If oInspectionMonitorWellsInfo Is Nothing Then
                    Add(id)
                End If
                Return oInspectionMonitorWellsInfo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False) As Boolean
            Dim oldID As Integer
            Try
                If Not bolValidated And Not oInspectionMonitorWellsInfo.Deleted And Not bolDelete Then
                    If Not Me.ValidateData() Then
                        Return False
                    End If
                End If
                If Not (oInspectionMonitorWellsInfo.ID < 0 And oInspectionMonitorWellsInfo.Deleted) Then
                    oldID = oInspectionMonitorWellsInfo.ID
                    oInspectionMonitorWellsDB.Put(oInspectionMonitorWellsInfo, moduleID, staffID, returnVal)
                    If Not bolValidated Then
                        If oldID < 0 Then
                            oInspection.MonitorWellsCollection.ChangeKey(oldID, oInspectionMonitorWellsInfo.ID)
                        End If
                    End If
                    oInspectionMonitorWellsInfo.Archive()
                    oInspectionMonitorWellsInfo.IsDirty = False
                End If
                If Not bolValidated And bolDelete Then
                    If oInspectionMonitorWellsInfo.Deleted Then
                        ' check if other inspections are present else load new instance
                        Dim strNext As String = Me.GetNext()
                        Dim strPrev As String = Me.GetPrevious()
                        If strNext = oInspectionMonitorWellsInfo.ID Then
                            If strPrev = oInspectionMonitorWellsInfo.ID Then
                                RaiseEvent evtInspectionMonitorWellsErr("Inspection " + oInspectionMonitorWellsInfo.ID.ToString + " deleted")
                                oInspection.MonitorWellsCollection.Remove(oInspectionMonitorWellsInfo)
                                If bolDelete Then
                                    oInspectionMonitorWellsInfo = New MUSTER.Info.InspectionMonitorWellsInfo
                                Else
                                    oInspectionMonitorWellsInfo = Me.Retrieve(oInspection, 0)
                                End If
                            Else
                                RaiseEvent evtInspectionMonitorWellsErr("Inspection " + oInspectionMonitorWellsInfo.ID.ToString + " deleted")
                                oInspection.MonitorWellsCollection.Remove(oInspectionMonitorWellsInfo)
                                oInspectionMonitorWellsInfo = Me.Retrieve(oInspection, strPrev)
                            End If
                        Else
                            RaiseEvent evtInspectionMonitorWellsErr("Inspection " + oInspectionMonitorWellsInfo.ID.ToString + " deleted")
                            oInspection.MonitorWellsCollection.Remove(oInspectionMonitorWellsInfo)
                            oInspectionMonitorWellsInfo = Me.Retrieve(oInspection, strNext)
                        End If
                    End If
                End If
                RaiseEvent evtInspectionMonitorWellsErr(oInspectionMonitorWellsInfo.IsDirty)
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
                oInspectionMonitorWellsInfo = oInspectionMonitorWellsDB.DBGetByID(id)
                If oInspectionMonitorWellsInfo.ID = 0 Then
                    oInspectionMonitorWellsInfo.ID = nID
                    nID -= 1
                End If
                oInspection.MonitorWellsCollection.Add(oInspectionMonitorWellsInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oInspectionMonitorWells As MUSTER.Info.InspectionMonitorWellsInfo)
            Try
                oInspectionMonitorWellsInfo = oInspectionMonitorWells
                If oInspectionMonitorWellsInfo.ID = 0 Then
                    oInspectionMonitorWellsInfo.ID = nID
                    nID -= 1
                End If
                oInspection.MonitorWellsCollection.Add(oInspectionMonitorWellsInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal id As Int64)
            Try
                If oInspection.MonitorWellsCollection.Contains(id) Then
                    oInspection.MonitorWellsCollection.Remove(id)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oInspectionMonitorWells As MUSTER.Info.InspectionMonitorWellsInfo)
            Try
                If oInspection.MonitorWellsCollection.Contains(oInspectionMonitorWells) Then
                    oInspection.MonitorWellsCollection.Remove(oInspectionMonitorWells)
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
            Dim mwInfo As MUSTER.Info.InspectionMonitorWellsInfo
            Try
                For Each mwInfo In oInspection.MonitorWellsCollection.Values
                    If mwInfo.IsDirty Or mwInfo.ID < 0 Then
                        oInspectionMonitorWellsInfo = mwInfo
                        If oInspectionMonitorWellsInfo.Deleted Then
                            If oInspectionMonitorWellsInfo.ID < 0 Then
                                delIDs.Add(oInspectionMonitorWellsInfo.ID)
                            Else
                                Me.Save(moduleID, staffID, returnVal, True)
                            End If
                        Else
                            If Me.ValidateData Then
                                If oInspectionMonitorWellsInfo.ID < 0 Then
                                    IDs.Add(oInspectionMonitorWellsInfo.ID)
                                End If
                                Me.Save(moduleID, staffID, returnVal, True)
                            Else : Exit For
                            End If
                        End If
                    End If
                Next
                If Not (delIDs Is Nothing) Then
                    For index = 1 To delIDs.Count
                        mwInfo = oInspection.MonitorWellsCollection.Item(CType(delIDs.Item(index), String))
                        oInspection.MonitorWellsCollection.Remove(mwInfo)
                    Next
                End If
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        mwInfo = oInspection.MonitorWellsCollection.Item(colKey)
                        oInspection.MonitorWellsCollection.ChangeKey(colKey, mwInfo.ID)
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
            Dim strArr() As String = oInspection.MonitorWellsCollection.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            For Each y As String In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return oInspection.MonitorWellsCollection.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            ElseIf oInspection.MonitorWellsCollection.Count <> 0 Then
                Return oInspection.MonitorWellsCollection.Item(nArr.GetValue(colIndex)).ID.ToString
            Else
                Return "0"
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oInspectionMonitorWellsInfo = New MUSTER.Info.InspectionMonitorWellsInfo
        End Sub
        Public Sub Reset()
            oInspectionMonitorWellsInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oInspectionMonitorWellsInfoLocal As New MUSTER.Info.InspectionMonitorWellsInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("InspectionMonitorWells ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oInspectionMonitorWellsInfoLocal In oInspection.MonitorWellsCollection.Values
                    dr = tbEntityTable.NewRow()
                    dr("replace ID") = oInspectionMonitorWellsInfoLocal.ID
                    dr("Deleted") = oInspectionMonitorWellsInfoLocal.Deleted
                    dr("Created By") = oInspectionMonitorWellsInfoLocal.CreatedBy
                    dr("Date Created") = oInspectionMonitorWellsInfoLocal.CreatedOn
                    dr("Last Edited By") = oInspectionMonitorWellsInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oInspectionMonitorWellsInfoLocal.ModifiedOn
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
        Private Sub InspectionMonitorWellsInfoChanged(ByVal bolValue As Boolean) Handles oInspectionMonitorWellsInfo.evtInspectionMonitorWellsInfoChanged
            RaiseEvent evtInspectionMonitorWellsChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
