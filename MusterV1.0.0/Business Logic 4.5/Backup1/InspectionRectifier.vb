'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.InspectionRectifier
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
' NOTE: This file to be used as InspectionRectifier to build other objects.
'       Replace keyword "InspectionRectifier" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pInspectionRectifier
#Region "Public Events"
        Public Event evtInspectionRectifierErr(ByVal MsgStr As String)
        Public Event evtInspectionRectifierChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private oInspection As MUSTER.Info.InspectionInfo
        Private WithEvents oInspectionRectifierInfo As MUSTER.Info.InspectionRectifierInfo
        Private oInspectionRectifierDB As MUSTER.DataAccess.InspectionRectifierDB
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
            oInspectionRectifierInfo = New MUSTER.Info.InspectionRectifierInfo
            oInspectionRectifierDB = New MUSTER.DataAccess.InspectionRectifierDB
        End Sub
#End Region
#Region "Exposed Attributes"
        Public ReadOnly Property ID() As Int64
            Get
                Return oInspectionRectifierInfo.ID
            End Get
        End Property
        Public Property InspectionID() As Int64
            Get
                Return oInspectionRectifierInfo.InspectionID
            End Get
            Set(ByVal Value As Int64)
                oInspectionRectifierInfo.InspectionID = Value
            End Set
        End Property
        Public Property QuestionID() As Int64
            Get
                Return oInspectionRectifierInfo.QuestionID
            End Get
            Set(ByVal Value As Int64)
                oInspectionRectifierInfo.QuestionID = Value
            End Set
        End Property
        Public Property RectifierOn() As Boolean
            Get
                Return oInspectionRectifierInfo.RectifierOn
            End Get
            Set(ByVal Value As Boolean)
                oInspectionRectifierInfo.RectifierOn = Value
            End Set
        End Property
        Public Property InopHowLong() As String
            Get
                Return oInspectionRectifierInfo.InopHowLong
            End Get
            Set(ByVal Value As String)
                oInspectionRectifierInfo.InopHowLong = Value
            End Set
        End Property
        Public Property Volts() As Double
            Get
                Return oInspectionRectifierInfo.Volts
            End Get
            Set(ByVal Value As Double)
                oInspectionRectifierInfo.Volts = Value
            End Set
        End Property
        Public Property Amps() As Double
            Get
                Return oInspectionRectifierInfo.Amps
            End Get
            Set(ByVal Value As Double)
                oInspectionRectifierInfo.Amps = Value
            End Set
        End Property
        Public Property Hours() As Double
            Get
                Return oInspectionRectifierInfo.Hours
            End Get
            Set(ByVal Value As Double)
                oInspectionRectifierInfo.Hours = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oInspectionRectifierInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oInspectionRectifierInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oInspectionRectifierInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oInspectionRectifierInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public ReadOnly Property colIsDirty() As Boolean
            Get
                Dim xInspectionRectifierinfo As MUSTER.Info.InspectionRectifierInfo
                For Each xInspectionRectifierinfo In oInspection.RectifiersCollection.Values
                    If xInspectionRectifierinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
        End Property
        Public Property CreatedBy() As String
            Get
                Return oInspectionRectifierInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oInspectionRectifierInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oInspectionRectifierInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oInspectionRectifierInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oInspectionRectifierInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oInspectionRectifierInfo.ModifiedOn
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
        Public Sub Load(ByRef checklistMaster As MUSTER.Info.InspectionInfo, ByRef ds As DataSet)
            Dim dr As DataRow
            oInspection = checklistMaster
            Try
                If ds.Tables("Rectifier").Rows.Count > 0 Then
                    For Each dr In ds.Tables("Rectifier").Rows
                        oInspectionRectifierInfo = New MUSTER.Info.InspectionRectifierInfo(dr)
                        oInspection.RectifiersCollection.Add(oInspectionRectifierInfo)
                    Next
                End If
                ds.Tables.Remove("Rectifier")
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function Retrieve(ByRef inspection As MUSTER.Info.InspectionInfo, ByVal id As Int64) As MUSTER.Info.InspectionRectifierInfo
            Try
                oInspection = inspection
                oInspectionRectifierInfo = oInspection.RectifiersCollection.Item(id)
                If oInspectionRectifierInfo Is Nothing Then
                    Add(id)
                End If
                Return oInspectionRectifierInfo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False) As Boolean
            Dim oldID As Integer
            Try
                If Not bolValidated And Not oInspectionRectifierInfo.Deleted And Not bolDelete Then
                    If Not Me.ValidateData() Then
                        Return False
                    End If
                End If
                If Not (oInspectionRectifierInfo.ID < 0 And oInspectionRectifierInfo.Deleted) Then
                    oldID = oInspectionRectifierInfo.ID
                    oInspectionRectifierDB.Put(oInspectionRectifierInfo, moduleID, staffID, returnVal)
                    If Not bolValidated Then
                        If oldID < 0 Then
                            oInspection.RectifiersCollection.ChangeKey(oldID, oInspectionRectifierInfo.ID)
                        End If
                    End If
                    oInspectionRectifierInfo.Archive()
                    oInspectionRectifierInfo.IsDirty = False
                End If
                If Not bolValidated And bolDelete Then
                    If oInspectionRectifierInfo.Deleted Then
                        ' check if other inspections are present else load new instance
                        Dim strNext As String = Me.GetNext()
                        Dim strPrev As String = Me.GetPrevious()
                        If strNext = oInspectionRectifierInfo.ID Then
                            If strPrev = oInspectionRectifierInfo.ID Then
                                RaiseEvent evtInspectionRectifierErr("Inspection " + oInspectionRectifierInfo.ID.ToString + " deleted")
                                oInspection.RectifiersCollection.Remove(oInspectionRectifierInfo)
                                If bolDelete Then
                                    oInspectionRectifierInfo = New MUSTER.Info.InspectionRectifierInfo
                                Else
                                    oInspectionRectifierInfo = Me.Retrieve(oInspection, 0)
                                End If
                            Else
                                RaiseEvent evtInspectionRectifierErr("Inspection " + oInspectionRectifierInfo.ID.ToString + " deleted")
                                oInspection.RectifiersCollection.Remove(oInspectionRectifierInfo)
                                oInspectionRectifierInfo = Me.Retrieve(oInspection, strPrev)
                            End If
                        Else
                            RaiseEvent evtInspectionRectifierErr("Inspection " + oInspectionRectifierInfo.ID.ToString + " deleted")
                            oInspection.RectifiersCollection.Remove(oInspectionRectifierInfo)
                            oInspectionRectifierInfo = Me.Retrieve(oInspection, strNext)
                        End If
                    End If
                End If
                RaiseEvent evtInspectionRectifierErr(oInspectionRectifierInfo.IsDirty)
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
                oInspectionRectifierInfo = oInspectionRectifierDB.DBGetByID(id)
                If oInspectionRectifierInfo.ID = 0 Then
                    oInspectionRectifierInfo.ID = nID
                    nID -= 1
                End If
                oInspection.RectifiersCollection.Add(oInspectionRectifierInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oInspectionRectifier As MUSTER.Info.InspectionRectifierInfo)
            Try
                oInspectionRectifierInfo = oInspectionRectifier
                If oInspectionRectifierInfo.ID = 0 Then
                    oInspectionRectifierInfo.ID = nID
                    nID -= 1
                End If
                oInspection.RectifiersCollection.Add(oInspectionRectifierInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal id As Int64)
            Try
                If oInspection.RectifiersCollection.Contains(id) Then
                    oInspection.RectifiersCollection.Remove(id)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oInspectionRectifier As MUSTER.Info.InspectionRectifierInfo)
            Try
                If oInspection.RectifiersCollection.Contains(oInspectionRectifier) Then
                    oInspection.RectifiersCollection.Remove(oInspectionRectifier)
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
            Dim rectInfo As MUSTER.Info.InspectionRectifierInfo
            Try
                For Each rectInfo In oInspection.RectifiersCollection.Values
                    If rectInfo.IsDirty Or rectInfo.ID < 0 Then
                        oInspectionRectifierInfo = rectInfo
                        If oInspectionRectifierInfo.Deleted Then
                            If oInspectionRectifierInfo.ID < 0 Then
                                delIDs.Add(oInspectionRectifierInfo.ID)
                            Else
                                Me.Save(moduleID, staffID, returnVal, True)
                            End If
                        Else
                            If Me.ValidateData Then
                                If oInspectionRectifierInfo.ID < 0 Then
                                    IDs.Add(oInspectionRectifierInfo.ID)
                                End If
                                Me.Save(moduleID, staffID, returnVal, True)
                            Else : Exit For
                            End If
                        End If
                    End If
                Next
                If Not (delIDs Is Nothing) Then
                    For index = 1 To delIDs.Count
                        rectInfo = oInspection.RectifiersCollection.Item(CType(delIDs.Item(index), String))
                        oInspection.RectifiersCollection.Remove(rectInfo)
                    Next
                End If
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        rectInfo = oInspection.RectifiersCollection.Item(colKey)
                        oInspection.RectifiersCollection.ChangeKey(colKey, rectInfo.ID)
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
            Dim strArr() As String = oInspection.RectifiersCollection.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            For Each y As String In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return oInspection.RectifiersCollection.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            ElseIf oInspection.RectifiersCollection.Count <> 0 Then
                Return oInspection.RectifiersCollection.Item(nArr.GetValue(colIndex)).ID.ToString
            Else
                Return "0"
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oInspectionRectifierInfo = New MUSTER.Info.InspectionRectifierInfo
        End Sub
        Public Sub Reset()
            oInspectionRectifierInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oInspectionRectifierInfoLocal As New MUSTER.Info.InspectionRectifierInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("InspectionRectifier ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oInspectionRectifierInfoLocal In oInspection.RectifiersCollection.Values
                    dr = tbEntityTable.NewRow()
                    dr("Replace ID") = oInspectionRectifierInfoLocal.ID
                    dr("Deleted") = oInspectionRectifierInfoLocal.Deleted
                    dr("Created By") = oInspectionRectifierInfoLocal.CreatedBy
                    dr("Date Created") = oInspectionRectifierInfoLocal.CreatedOn
                    dr("Last Edited By") = oInspectionRectifierInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oInspectionRectifierInfoLocal.ModifiedOn
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
        Private Sub InspectionRectifierInfoChanged(ByVal bolValue As Boolean) Handles oInspectionRectifierInfo.evtInspectionRectifierInfoChanged
            RaiseEvent evtInspectionRectifierChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
