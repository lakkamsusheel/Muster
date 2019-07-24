'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.InspectionCPReadings
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      MNR         6/10/05     Original class definition
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
' NOTE: This file to be used as InspectionCPReadings to build other objects.
'       Replace keyword "InspectionCPReadings" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pInspectionCPReadings
#Region "Public Events"
        Public Event evtInspectionCPReadingsErr(ByVal MsgStr As String)
        Public Event evtInspectionCPReadingsChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private oInspection As MUSTER.Info.InspectionInfo
        Private WithEvents oInspectionCPReadingsInfo As MUSTER.Info.InspectionCPReadingsInfo
        Private oInspectionCPReadingsDB As New MUSTER.DataAccess.InspectionCPReadingsDB
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
            oInspectionCPReadingsInfo = New MUSTER.Info.InspectionCPReadingsInfo
            oInspectionCPReadingsDB = New MUSTER.DataAccess.InspectionCPReadingsDB
        End Sub
#End Region
#Region "Exposed Attributes"
        Public ReadOnly Property ID() As Integer
            Get
                Return oInspectionCPReadingsInfo.ID
            End Get
        End Property
        Public Property InspectionID() As Integer
            Get
                Return oInspectionCPReadingsInfo.InspectionID
            End Get
            Set(ByVal Value As Integer)
                oInspectionCPReadingsInfo.InspectionID = Value
            End Set
        End Property
        Public Property QuestionID() As Integer
            Get
                Return oInspectionCPReadingsInfo.QuestionID
            End Get
            Set(ByVal Value As Integer)
                oInspectionCPReadingsInfo.QuestionID = Value
            End Set
        End Property
        Public Property TankPipeID() As Integer
            Get
                Return oInspectionCPReadingsInfo.TankPipeID
            End Get
            Set(ByVal Value As Integer)
                oInspectionCPReadingsInfo.TankPipeID = Value
            End Set
        End Property
        Public Property TankPipeIndex() As Integer
            Get
                Return oInspectionCPReadingsInfo.TankPipeIndex
            End Get
            Set(ByVal Value As Integer)
                oInspectionCPReadingsInfo.TankPipeIndex = Value
            End Set
        End Property
        Public Property TankPipeEntityID() As Integer
            Get
                Return oInspectionCPReadingsInfo.TankPipeEntityID
            End Get
            Set(ByVal Value As Integer)
                oInspectionCPReadingsInfo.TankPipeEntityID = Value
            End Set
        End Property
        'Public Property TankDispenser() As Boolean
        '    Get
        '        Return oInspectionCPReadingsInfo.TankDispenser
        '    End Get
        '    Set(ByVal Value As Boolean)
        '        oInspectionCPReadingsInfo.TankDispenser = Value
        '    End Set
        'End Property
        'Public Property Galvanic() As Boolean
        '    Get
        '        Return oInspectionCPReadingsInfo.Galvanic
        '    End Get
        '    Set(ByVal Value As Boolean)
        '        oInspectionCPReadingsInfo.Galvanic = Value
        '    End Set
        'End Property
        'Public Property ImpressedCurrent() As Boolean
        '    Get
        '        Return oInspectionCPReadingsInfo.ImpressedCurrent
        '    End Get
        '    Set(ByVal Value As Boolean)
        '        oInspectionCPReadingsInfo.ImpressedCurrent = Value
        '    End Set
        'End Property
        Public Property ContactPoint() As String
            Get
                Return oInspectionCPReadingsInfo.ContactPoint
            End Get
            Set(ByVal Value As String)
                oInspectionCPReadingsInfo.ContactPoint = Value
            End Set
        End Property
        Public Property LocalReferCellPlacement() As String
            Get
                Return oInspectionCPReadingsInfo.LocalReferCellPlacement
            End Get
            Set(ByVal Value As String)
                oInspectionCPReadingsInfo.LocalReferCellPlacement = Value
            End Set
        End Property
        Public Property LocalOn() As String
            Get
                Return oInspectionCPReadingsInfo.LocalOn
            End Get
            Set(ByVal Value As String)
                oInspectionCPReadingsInfo.LocalOn = Value
            End Set
        End Property
        Public Property RemoteOff() As String
            Get
                Return oInspectionCPReadingsInfo.RemoteOff
            End Get
            Set(ByVal Value As String)
                oInspectionCPReadingsInfo.RemoteOff = Value
            End Set
        End Property
        Public Property PassFailIncon() As Integer
            Get
                Return oInspectionCPReadingsInfo.PassFailIncon
            End Get
            Set(ByVal Value As Integer)
                oInspectionCPReadingsInfo.PassFailIncon = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oInspectionCPReadingsInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oInspectionCPReadingsInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oInspectionCPReadingsInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oInspectionCPReadingsInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public ReadOnly Property colIsDirty() As Boolean
            Get
                Dim xInspectionCPReadingsinfo As MUSTER.Info.InspectionCPReadingsInfo
                For Each xInspectionCPReadingsinfo In oInspection.CPReadingsCollection.Values
                    If xInspectionCPReadingsinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
        End Property
        Public Property CreatedBy() As String
            Get
                Return oInspectionCPReadingsInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oInspectionCPReadingsInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oInspectionCPReadingsInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oInspectionCPReadingsInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oInspectionCPReadingsInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oInspectionCPReadingsInfo.ModifiedOn
            End Get
        End Property
        Public Property RemoteReferCellPlacement() As Boolean
            Get
                Return oInspectionCPReadingsInfo.RemoteReferCellPlacement
            End Get
            Set(ByVal Value As Boolean)
                oInspectionCPReadingsInfo.RemoteReferCellPlacement = Value
            End Set
        End Property
        Public Property GalvanicIC() As Boolean
            Get
                Return oInspectionCPReadingsInfo.GalvanicIC
            End Get
            Set(ByVal Value As Boolean)
                oInspectionCPReadingsInfo.GalvanicIC = Value
            End Set
        End Property
        Public Property GalvanicICResponse() As Integer
            Get
                Return oInspectionCPReadingsInfo.GalvanicICResponse
            End Get
            Set(ByVal Value As Integer)
                oInspectionCPReadingsInfo.GalvanicICResponse = Value
            End Set
        End Property
        Public Property TestedByInspector() As Boolean
            Get
                Return oInspectionCPReadingsInfo.TestedByInspector
            End Get
            Set(ByVal Value As Boolean)
                oInspectionCPReadingsInfo.TestedByInspector = Value
            End Set
        End Property
        Public Property TestedByInspectorResponse() As Boolean
            Get
                Return oInspectionCPReadingsInfo.TestedByInspectorResponse
            End Get
            Set(ByVal Value As Boolean)
                oInspectionCPReadingsInfo.TestedByInspectorResponse = Value
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
        Public Sub Load(ByRef checkListMaster As MUSTER.Info.InspectionInfo, ByRef ds As DataSet)
            Dim dr As DataRow
            oInspection = checkListMaster
            Try
                If ds.Tables("CPReadings").Rows.Count > 0 Then
                    For Each dr In ds.Tables("CPReadings").Rows
                        oInspectionCPReadingsInfo = New MUSTER.Info.InspectionCPReadingsInfo(dr)
                        oInspection.CPReadingsCollection.Add(oInspectionCPReadingsInfo)
                    Next
                End If
                ds.Tables.Remove("CPReadings")
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function Retrieve(ByRef inspection As MUSTER.Info.InspectionInfo, ByVal id As Integer) As MUSTER.Info.InspectionCPReadingsInfo
            Try
                oInspection = inspection
                oInspectionCPReadingsInfo = oInspection.CPReadingsCollection.Item(id)
                If oInspectionCPReadingsInfo Is Nothing Then
                    Add(id)
                End If
                Return oInspectionCPReadingsInfo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False) As Boolean
            Dim oldID As Integer
            Try
                If Not bolValidated And Not oInspectionCPReadingsInfo.Deleted And Not bolDelete Then
                    If Not Me.ValidateData() Then
                        Return False
                    End If
                End If
                If Not (oInspectionCPReadingsInfo.ID < 0 And oInspectionCPReadingsInfo.Deleted) Then
                    oldID = oInspectionCPReadingsInfo.ID
                    oInspectionCPReadingsDB.Put(oInspectionCPReadingsInfo, moduleID, staffID, returnVal)
                    If Not bolValidated Then
                        If oldID < 0 Then
                            oInspection.CPReadingsCollection.ChangeKey(oldID, oInspectionCPReadingsInfo.ID)
                        End If
                    End If
                    oInspectionCPReadingsInfo.Archive()
                    oInspectionCPReadingsInfo.IsDirty = False
                End If
                If Not bolValidated And bolDelete Then
                    If oInspectionCPReadingsInfo.Deleted Then
                        ' check if other inspections are present else load new instance
                        Dim strNext As String = Me.GetNext()
                        Dim strPrev As String = Me.GetPrevious()
                        If strNext = oInspectionCPReadingsInfo.ID Then
                            If strPrev = oInspectionCPReadingsInfo.ID Then
                                RaiseEvent evtInspectionCPReadingsErr("Inspection " + oInspectionCPReadingsInfo.ID.ToString + " deleted")
                                oInspection.CPReadingsCollection.Remove(oInspectionCPReadingsInfo)
                                If bolDelete Then
                                    oInspectionCPReadingsInfo = New MUSTER.Info.InspectionCPReadingsInfo
                                Else
                                    oInspectionCPReadingsInfo = Me.Retrieve(oInspection, 0)
                                End If
                            Else
                                RaiseEvent evtInspectionCPReadingsErr("Inspection " + oInspectionCPReadingsInfo.ID.ToString + " deleted")
                                oInspection.CPReadingsCollection.Remove(oInspectionCPReadingsInfo)
                                oInspectionCPReadingsInfo = Me.Retrieve(oInspection, strPrev)
                            End If
                        Else
                            RaiseEvent evtInspectionCPReadingsErr("Inspection " + oInspectionCPReadingsInfo.ID.ToString + " deleted")
                            oInspection.CPReadingsCollection.Remove(oInspectionCPReadingsInfo)
                            oInspectionCPReadingsInfo = Me.Retrieve(oInspection, strNext)
                        End If
                    End If
                End If
                RaiseEvent evtInspectionCPReadingsErr(oInspectionCPReadingsInfo.IsDirty)
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
        Public Sub Add(ByVal id As Integer)
            Try
                oInspectionCPReadingsInfo = oInspectionCPReadingsDB.DBGetByID(id)
                If oInspectionCPReadingsInfo.ID = 0 Then
                    oInspectionCPReadingsInfo.ID = nID
                    nID -= 1
                End If
                oInspection.CPReadingsCollection.Add(oInspectionCPReadingsInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oInspectionCPReadings As MUSTER.Info.InspectionCPReadingsInfo)
            Try
                oInspectionCPReadingsInfo = oInspectionCPReadings
                If oInspectionCPReadingsInfo.ID = 0 Then
                    oInspectionCPReadingsInfo.ID = nID
                    nID -= 1
                End If
                oInspection.CPReadingsCollection.Add(oInspectionCPReadingsInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal id As Integer)
            Try
                If oInspection.CPReadingsCollection.Contains(id) Then
                    oInspection.CPReadingsCollection.Remove(id)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oInspectionCPReadings As MUSTER.Info.InspectionCPReadingsInfo)
            Try
                If oInspection.CPReadingsCollection.Contains(oInspectionCPReadings) Then
                    oInspection.CPReadingsCollection.Remove(oInspectionCPReadings)
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
            Dim cpInfo As MUSTER.Info.InspectionCPReadingsInfo
            Try
                For Each cpInfo In oInspection.CPReadingsCollection.Values
                    If cpInfo.IsDirty Or cpInfo.ID < 0 Then
                        oInspectionCPReadingsInfo = cpInfo
                        If oInspectionCPReadingsInfo.Deleted Then
                            If oInspectionCPReadingsInfo.ID < 0 Then
                                delIDs.Add(oInspectionCPReadingsInfo.ID)
                            Else
                                Me.Save(moduleID, staffID, returnVal, True)
                            End If
                        Else
                            If Me.ValidateData Then
                                If oInspectionCPReadingsInfo.ID < 0 Then
                                    IDs.Add(oInspectionCPReadingsInfo.ID)
                                End If
                                Me.Save(moduleID, staffID, returnVal, True)
                            Else : Exit For
                            End If
                        End If
                    End If
                Next
                If Not (delIDs Is Nothing) Then
                    For index = 1 To delIDs.Count
                        cpInfo = oInspection.CPReadingsCollection.Item(CType(delIDs.Item(index), String))
                        oInspection.CPReadingsCollection.Remove(cpInfo)
                    Next
                End If
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        cpInfo = oInspection.CPReadingsCollection.Item(colKey)
                        oInspection.CPReadingsCollection.ChangeKey(colKey, cpInfo.ID)
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
            Dim strArr() As String = oInspection.CPReadingsCollection.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            For Each y As String In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return oInspection.CPReadingsCollection.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            ElseIf oInspection.CPReadingsCollection.Count <> 0 Then
                Return oInspection.CPReadingsCollection.Item(nArr.GetValue(colIndex)).ID.ToString
            Else
                Return "0"
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oInspectionCPReadingsInfo = New MUSTER.Info.InspectionCPReadingsInfo
        End Sub
        Public Sub Reset()
            oInspectionCPReadingsInfo.Reset()
        End Sub
        Public Sub Archive()
            oInspectionCPReadingsInfo.Archive()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oInspectionCPReadingsInfoLocal As New MUSTER.Info.InspectionCPReadingsInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("InspectionCPReadings ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oInspectionCPReadingsInfoLocal In oInspection.CPReadingsCollection.Values
                    dr = tbEntityTable.NewRow()
                    dr("ReplaceID") = oInspectionCPReadingsInfoLocal.ID
                    dr("Deleted") = oInspectionCPReadingsInfoLocal.Deleted
                    dr("Created By") = oInspectionCPReadingsInfoLocal.CreatedBy
                    dr("Date Created") = oInspectionCPReadingsInfoLocal.CreatedOn
                    dr("Last Edited By") = oInspectionCPReadingsInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oInspectionCPReadingsInfoLocal.ModifiedOn
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
        Private Sub InspectionCPReadingsInfoChanged(ByVal bolValue As Boolean) Handles oInspectionCPReadingsInfo.evtInspectionCPReadingsInfoChanged
            RaiseEvent evtInspectionCPReadingsChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
