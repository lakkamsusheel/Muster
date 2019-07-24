'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.InspectionSketch
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
' NOTE: This file to be used as InspectionSketch to build other objects.
'       Replace keyword "InspectionSketch" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pInspectionSketch
#Region "Public Events"
        Public Event evtInspectionSketchErr(ByVal MsgStr As String)
        Public Event evtInspectionSketchChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private oInspection As MUSTER.Info.InspectionInfo
        Private WithEvents oInspectionSketchInfo As MUSTER.Info.InspectionSketchInfo
        Private oInspectionSketchDB As MUSTER.DataAccess.InspectionSketchDB
        Private MusterException As MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private alPics As Collections.ArrayList
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

            oInspectionSketchInfo = New MUSTER.Info.InspectionSketchInfo
            oInspectionSketchDB = New MUSTER.DataAccess.InspectionSketchDB
        End Sub

        Sub dispose()
            If Not Me.alPics Is Nothing Then
                Me.alPics.Clear()
                Me.alPics = Nothing
            End If

        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property Pics() As Collections.ArrayList
            Get
                If alPics Is Nothing Then
                    alPics = New Collections.ArrayList
                End If
                Return alPics
            End Get

            Set(ByVal Value As Collections.ArrayList)
                alPics = Value
            End Set

        End Property
        Public ReadOnly Property ID() As Int64
            Get
                Return oInspectionSketchInfo.ID
            End Get
        End Property
        Public Property InspectionID() As Int64
            Get
                Return oInspectionSketchInfo.InspectionID
            End Get
            Set(ByVal Value As Int64)
                oInspectionSketchInfo.InspectionID = Value
            End Set
        End Property
        Public Property SketchFileName() As String
            Get
                Return oInspectionSketchInfo.SketchFileName
            End Get
            Set(ByVal Value As String)
                oInspectionSketchInfo.SketchFileName = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oInspectionSketchInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oInspectionSketchInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oInspectionSketchInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oInspectionSketchInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public ReadOnly Property colIsDirty() As Boolean
            Get
                Dim xInspectionSketchinfo As MUSTER.Info.InspectionSketchInfo
                For Each xInspectionSketchinfo In oInspection.SketchsCollection.Values
                    If xInspectionSketchinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
        End Property
        Public Property CreatedBy() As String
            Get
                Return oInspectionSketchInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oInspectionSketchInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oInspectionSketchInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oInspectionSketchInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oInspectionSketchInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oInspectionSketchInfo.ModifiedOn
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
                If ds.Tables("Sketch").Rows.Count > 0 Then
                    For Each dr In ds.Tables("Sketch").Rows
                        oInspectionSketchInfo = New MUSTER.Info.InspectionSketchInfo(dr)
                        oInspection.SketchsCollection.Add(oInspectionSketchInfo)
                    Next
                End If
                ds.Tables.Remove("Sketch")
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Sub LoadPics(ByVal ds As DataSet, ByVal path As String)

            Dim dr As DataRow

            Try
                If ds.Tables("FacPics").Rows.Count > 0 Then
                    For Each dr In ds.Tables("FacPics").Rows
                        If System.IO.File.Exists(String.Format("{0}\{1}", path, dr.Item("FacPic"))) Then
                            Me.Pics.Add(String.Format("{0}\{1}", path, dr.Item("FacPic")))
                        End If
                    Next
                End If
                ds.Tables.Remove("facPics")
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Sub

        Public Function Retrieve(ByRef inspection As MUSTER.Info.InspectionInfo, ByVal id As Int64) As MUSTER.Info.InspectionSketchInfo
            Try
                oInspection = inspection
                oInspectionSketchInfo = oInspection.SketchsCollection.Item(id)
                If oInspectionSketchInfo Is Nothing Then
                    Add(id)
                End If
                Return oInspectionSketchInfo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function Retrieve(ByRef inspection As MUSTER.Info.InspectionInfo) As MUSTER.Info.InspectionSketchInfo
            Try
                oInspection = inspection
                oInspection.SketchsCollection.Clear()

                Add(inspection)

                Return oInspectionSketchInfo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False) As Boolean
            Dim oldID As Integer
            Try
                If Not bolValidated And Not oInspectionSketchInfo.Deleted And Not bolDelete Then
                    If Not Me.ValidateData() Then
                        Return False
                    End If
                End If
                If Not (oInspectionSketchInfo.ID < 0 And oInspectionSketchInfo.Deleted) Then
                    oldID = oInspectionSketchInfo.ID
                    oInspectionSketchDB.Put(oInspectionSketchInfo, moduleID, staffID, returnVal)
                    If Not bolValidated Then
                        If oldID < 0 Then
                            oInspection.SketchsCollection.ChangeKey(oldID, oInspectionSketchInfo.ID)
                        End If
                    End If
                    oInspectionSketchInfo.Archive()
                    oInspectionSketchInfo.IsDirty = False
                End If
                If Not bolValidated And bolDelete Then
                    If oInspectionSketchInfo.Deleted Then
                        ' check if other inspections are present else load new instance
                        Dim strNext As String = Me.GetNext()
                        Dim strPrev As String = Me.GetPrevious()
                        If strNext = oInspectionSketchInfo.ID Then
                            If strPrev = oInspectionSketchInfo.ID Then
                                RaiseEvent evtInspectionSketchErr("Inspection " + oInspectionSketchInfo.ID.ToString + " deleted")
                                oInspection.SketchsCollection.Remove(oInspectionSketchInfo)
                                If bolDelete Then
                                    oInspectionSketchInfo = New MUSTER.Info.InspectionSketchInfo
                                Else
                                    oInspectionSketchInfo = Me.Retrieve(oInspection, 0)
                                End If
                            Else
                                RaiseEvent evtInspectionSketchErr("Inspection " + oInspectionSketchInfo.ID.ToString + " deleted")
                                oInspection.SketchsCollection.Remove(oInspectionSketchInfo)
                                oInspectionSketchInfo = Me.Retrieve(oInspection, strPrev)
                            End If
                        Else
                            RaiseEvent evtInspectionSketchErr("Inspection " + oInspectionSketchInfo.ID.ToString + " deleted")
                            oInspection.SketchsCollection.Remove(oInspectionSketchInfo)
                            oInspectionSketchInfo = Me.Retrieve(oInspection, strNext)
                        End If
                    End If
                End If
                RaiseEvent evtInspectionSketchErr(oInspectionSketchInfo.IsDirty)
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
                oInspectionSketchInfo = oInspectionSketchDB.DBGetByID(id)
                If oInspectionSketchInfo.ID = 0 Then
                    oInspectionSketchInfo.ID = nID
                    nID -= 1
                End If
                oInspection.SketchsCollection.Add(oInspectionSketchInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        'Adds an entity to the collection as called for by Inspection
        Public Sub Add(ByVal inspection As Info.InspectionInfo)
            Try
                oInspectionSketchInfo = oInspectionSketchDB.DBGetByInspectionID(inspection.ID)
                If oInspectionSketchInfo.ID = 0 Then
                    oInspectionSketchInfo.ID = nID
                    nID -= 1
                End If
                oInspection.SketchsCollection.Add(oInspectionSketchInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oInspectionSketch As MUSTER.Info.InspectionSketchInfo)
            Try
                oInspectionSketchInfo = oInspectionSketch
                If oInspectionSketchInfo.ID = 0 Then
                    oInspectionSketchInfo.ID = nID
                    nID -= 1
                End If
                oInspection.SketchsCollection.Add(oInspectionSketchInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal id As Int64)
            Try
                If oInspection.SketchsCollection.Contains(id) Then
                    oInspection.SketchsCollection.Remove(id)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oInspectionSketch As MUSTER.Info.InspectionSketchInfo)
            Try
                If oInspection.SketchsCollection.Contains(oInspectionSketch) Then
                    oInspection.SketchsCollection.Remove(oInspectionSketch)
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
            Dim sketchInfo As MUSTER.Info.InspectionSketchInfo
            Try
                For Each sketchInfo In oInspection.SketchsCollection.Values
                    If sketchInfo.IsDirty Or sketchInfo.ID < 0 Then
                        oInspectionSketchInfo = sketchInfo
                        If oInspectionSketchInfo.InspectionID <= 0 Then
                            MsgBox(String.Format("Warning: A Sketch of the facility has not been drawn yet."), MsgBoxStyle.OKOnly)

                        Else

                            If oInspectionSketchInfo.Deleted Then
                                If oInspectionSketchInfo.ID < 0 Then
                                    delIDs.Add(oInspectionSketchInfo.ID)
                                Else
                                    Me.Save(moduleID, staffID, returnVal, True)
                                End If
                            Else
                                If Me.ValidateData Then
                                    If oInspectionSketchInfo.ID < 0 Then
                                        IDs.Add(oInspectionSketchInfo.ID)
                                    End If
                                    Me.Save(moduleID, staffID, returnVal, True)
                                Else : Exit For
                                End If
                            End If
                        End If

                    End If
                Next
                If Not (delIDs Is Nothing) Then
                    For index = 1 To delIDs.Count
                        sketchInfo = oInspection.SketchsCollection.Item(CType(delIDs.Item(index), String))
                        oInspection.SketchsCollection.Remove(sketchInfo)
                    Next
                End If
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        sketchInfo = oInspection.SketchsCollection.Item(colKey)
                        oInspection.SketchsCollection.ChangeKey(colKey, sketchInfo.ID)
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
            Dim strArr() As String = oInspection.SketchsCollection.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            For Each y As String In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return oInspection.SketchsCollection.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            ElseIf oInspection.SketchsCollection.Count <> 0 Then
                Return oInspection.SketchsCollection.Item(nArr.GetValue(colIndex)).ID.ToString
            Else
                Return "0"
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oInspectionSketchInfo = New MUSTER.Info.InspectionSketchInfo
        End Sub
        Public Sub Reset()
            oInspectionSketchInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oInspectionSketchInfoLocal As New MUSTER.Info.InspectionSketchInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("InspectionSketch ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oInspectionSketchInfoLocal In oInspection.SketchsCollection.Values
                    dr = tbEntityTable.NewRow()
                    dr("ReplaceID") = oInspectionSketchInfoLocal.ID
                    dr("Deleted") = oInspectionSketchInfoLocal.Deleted
                    dr("Created By") = oInspectionSketchInfoLocal.CreatedBy
                    dr("Date Created") = oInspectionSketchInfoLocal.CreatedOn
                    dr("Last Edited By") = oInspectionSketchInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oInspectionSketchInfoLocal.ModifiedOn
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
        Private Sub InspectionSketchInfoChanged(ByVal bolValue As Boolean) Handles oInspectionSketchInfo.evtInspectionSketchInfoChanged
            RaiseEvent evtInspectionSketchChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
