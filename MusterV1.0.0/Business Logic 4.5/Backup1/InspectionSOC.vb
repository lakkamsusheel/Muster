'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.InspectionSOC
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
' NOTE: This file to be used as InspectionSOC to build other objects.
'       Replace keyword "InspectionSOC" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pInspectionSOC
#Region "Public Events"
        Public Event evtInspectionSOCErr(ByVal MsgStr As String)
        Public Event evtInspectionSOCChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private oInspection As MUSTER.Info.InspectionInfo
        Private WithEvents oInspectionSOCInfo As MUSTER.Info.InspectionSOCInfo
        Private oInspectionSOCDB As New MUSTER.DataAccess.InspectionSOCDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
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

            oInspectionSOCInfo = New MUSTER.Info.InspectionSOCInfo
            oInspectionSOCDB = New MUSTER.DataAccess.InspectionSOCDB
        End Sub
#End Region
#Region "Exposed Attributes"
        Public ReadOnly Property ID() As Int64
            Get
                Return oInspectionSOCInfo.ID
            End Get
        End Property
        Public Property InspectionID() As Int64
            Get
                Return oInspectionSOCInfo.InspectionID
            End Get
            Set(ByVal Value As Int64)
                oInspectionSOCInfo.InspectionID = Value
            End Set
        End Property
        Public Property LeakPrevention() As Int64
            Get
                Return oInspectionSOCInfo.LeakPrevention
            End Get
            Set(ByVal Value As Int64)
                oInspectionSOCInfo.LeakPrevention = Value
            End Set
        End Property
        Public Property LeakPreventionCitation() As String
            Get
                Return oInspectionSOCInfo.LeakPreventionCitation
            End Get
            Set(ByVal Value As String)
                oInspectionSOCInfo.LeakPreventionCitation = Value
            End Set
        End Property
        Public Property LeakPreventionLineNumbers() As String
            Get
                Return oInspectionSOCInfo.LeakPreventionLineNumbers
            End Get
            Set(ByVal Value As String)
                oInspectionSOCInfo.LeakPreventionLineNumbers = Value
            End Set
        End Property
        Public Property LeakDetection() As Int64
            Get
                Return oInspectionSOCInfo.LeakDetection
            End Get
            Set(ByVal Value As Int64)
                oInspectionSOCInfo.LeakDetection = Value
            End Set
        End Property
        Public Property LeakDetectionCitation() As String
            Get
                Return oInspectionSOCInfo.LeakDetectionCitation
            End Get
            Set(ByVal Value As String)
                oInspectionSOCInfo.LeakDetectionCitation = Value
            End Set
        End Property
        Public Property LeakDetectionLineNumbers() As String
            Get
                Return oInspectionSOCInfo.LeakDetectionLineNumbers
            End Get
            Set(ByVal Value As String)
                oInspectionSOCInfo.LeakDetectionLineNumbers = Value
            End Set
        End Property
        Public Property LeakPreventionDetection() As Int64
            Get
                Return oInspectionSOCInfo.LeakPreventionDetection
            End Get
            Set(ByVal Value As Int64)
                oInspectionSOCInfo.LeakPreventionDetection = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oInspectionSOCInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oInspectionSOCInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property CAEOverride() As Boolean
            Get
                Return oInspectionSOCInfo.CAEOverride
            End Get
            Set(ByVal Value As Boolean)
                oInspectionSOCInfo.CAEOverride = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oInspectionSOCInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oInspectionSOCInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public ReadOnly Property colIsDirty() As Boolean
            Get
                Dim xInspectionSOCinfo As MUSTER.Info.InspectionSOCInfo
                For Each xInspectionSOCinfo In oInspection.SOCsCollection.Values
                    If xInspectionSOCinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
        End Property
        Public Property CreatedBy() As String
            Get
                Return oInspectionSOCInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oInspectionSOCInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oInspectionSOCInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oInspectionSOCInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oInspectionSOCInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oInspectionSOCInfo.ModifiedOn
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
                If ds.Tables("SOC").Rows.Count > 0 Then
                    For Each dr In ds.Tables("SOC").Rows
                        oInspectionSOCInfo = New MUSTER.Info.InspectionSOCInfo(dr)
                        oInspection.SOCsCollection.Add(oInspectionSOCInfo)
                    Next
                End If
                ds.Tables.Remove("SOC")
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function Retrieve(ByRef inspection As MUSTER.Info.InspectionInfo, ByVal id As Int64) As MUSTER.Info.InspectionSOCInfo
            Try
                oInspection = inspection
                oInspectionSOCInfo = oInspection.SOCsCollection.Item(id)
                If oInspectionSOCInfo Is Nothing Then
                    Add(id)
                End If
                Return oInspectionSOCInfo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False) As Boolean
            Dim oldID As Integer
            Try
                If Not bolValidated And Not oInspectionSOCInfo.Deleted And Not bolDelete Then
                    If Not Me.ValidateData() Then
                        Return False
                    End If
                End If
                If Not (oInspectionSOCInfo.ID < 0 And oInspectionSOCInfo.Deleted) Then
                    oldID = oInspectionSOCInfo.ID
                    oInspectionSOCDB.Put(oInspectionSOCInfo, moduleID, staffID, returnVal)
                    If Not bolValidated Then
                        If oldID < 0 Then
                            oInspection.SOCsCollection.ChangeKey(oldID, oInspectionSOCInfo.ID)
                        End If
                    End If
                    oInspectionSOCInfo.Archive()
                    oInspectionSOCInfo.IsDirty = False
                End If
                If Not bolValidated And bolDelete Then
                    If oInspectionSOCInfo.Deleted Then
                        ' check if other inspections are present else load new instance
                        Dim strNext As String = Me.GetNext()
                        Dim strPrev As String = Me.GetPrevious()
                        If strNext = oInspectionSOCInfo.ID Then
                            If strPrev = oInspectionSOCInfo.ID Then
                                RaiseEvent evtInspectionSOCErr("Inspection " + oInspectionSOCInfo.ID.ToString + " deleted")
                                oInspection.SOCsCollection.Remove(oInspectionSOCInfo)
                                If bolDelete Then
                                    oInspectionSOCInfo = New MUSTER.Info.InspectionSOCInfo
                                Else
                                    oInspectionSOCInfo = Me.Retrieve(oInspection, 0)
                                End If
                            Else
                                RaiseEvent evtInspectionSOCErr("Inspection " + oInspectionSOCInfo.ID.ToString + " deleted")
                                oInspection.SOCsCollection.Remove(oInspectionSOCInfo)
                                oInspectionSOCInfo = Me.Retrieve(oInspection, strPrev)
                            End If
                        Else
                            RaiseEvent evtInspectionSOCErr("Inspection " + oInspectionSOCInfo.ID.ToString + " deleted")
                            oInspection.SOCsCollection.Remove(oInspectionSOCInfo)
                            oInspectionSOCInfo = Me.Retrieve(oInspection, strNext)
                        End If
                    End If
                End If
                RaiseEvent evtInspectionSOCErr(oInspectionSOCInfo.IsDirty)
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
                oInspectionSOCInfo = oInspectionSOCDB.DBGetByID(id)
                If oInspectionSOCInfo.ID = 0 Then
                    oInspectionSOCInfo.ID = nID
                    nID -= 1
                End If
                oInspection.SOCsCollection.Add(oInspectionSOCInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oInspectionSOC As MUSTER.Info.InspectionSOCInfo)
            Try
                oInspectionSOCInfo = oInspectionSOC
                If oInspectionSOCInfo.ID = 0 Then
                    oInspectionSOCInfo.ID = nID
                    nID -= 1
                End If
                oInspection.SOCsCollection.Add(oInspectionSOCInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal id As Int64)
            Try
                If oInspection.SOCsCollection.Contains(id) Then
                    oInspection.SOCsCollection.Remove(id)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oInspectionSOC As MUSTER.Info.InspectionSOCInfo)
            Try
                If oInspection.SOCsCollection.Contains(oInspectionSOC) Then
                    oInspection.SOCsCollection.Remove(oInspectionSOC)
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
            Dim socInfo As MUSTER.Info.InspectionSOCInfo
            Try
                For Each socInfo In oInspection.SOCsCollection.Values
                    If socInfo.IsDirty Or socInfo.ID < 0 Then
                        oInspectionSOCInfo = socInfo
                        If oInspectionSOCInfo.Deleted Then
                            If oInspectionSOCInfo.ID < 0 Then
                                delIDs.Add(oInspectionSOCInfo.ID)
                            Else
                                Me.Save(moduleID, staffID, returnVal, True)
                            End If
                        Else
                            If Me.ValidateData Then
                                If oInspectionSOCInfo.ID < 0 Then
                                    IDs.Add(oInspectionSOCInfo.ID)
                                End If
                                Me.Save(moduleID, staffID, returnVal, True)
                            Else : Exit For
                            End If
                        End If
                    End If
                Next
                If Not (delIDs Is Nothing) Then
                    For index = 1 To delIDs.Count
                        socInfo = oInspection.SOCsCollection.Item(CType(delIDs.Item(index), String))
                        oInspection.SOCsCollection.Remove(socInfo)
                    Next
                End If
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        socInfo = oInspection.SOCsCollection.Item(colKey)
                        oInspection.SOCsCollection.ChangeKey(colKey, socInfo.ID)
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
            Dim strArr() As String = oInspection.SOCsCollection.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            For Each y As String In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return oInspection.SOCsCollection.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            ElseIf oInspection.SOCsCollection.Count <> 0 Then
                Return oInspection.SOCsCollection.Item(nArr.GetValue(colIndex)).ID.ToString
            Else
                Return "0"
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oInspectionSOCInfo = New MUSTER.Info.InspectionSOCInfo
        End Sub
        Public Sub Reset()
            oInspectionSOCInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        'Public Function EntityTable() As DataTable
        '    Dim oInspectionSOCInfoLocal As New MUSTER.Info.InspectionSOCInfo
        '    Dim dr As DataRow
        '    Dim tbEntityTable As New DataTable
        '    Try
        '        tbEntityTable.Columns.Add("InspectionSOC ID")
        '        tbEntityTable.Columns.Add("Deleted")
        '        tbEntityTable.Columns.Add("Created By")
        '        tbEntityTable.Columns.Add("Date Created")
        '        tbEntityTable.Columns.Add("Last Edited By")
        '        tbEntityTable.Columns.Add("Date Last Edited")

        '        For Each oInspectionSOCInfoLocal In oInspection.SOCsCollection.Values
        '            dr = tbEntityTable.NewRow()
        '            dr("REPLACE ID") = oInspectionSOCInfoLocal.ID
        '            dr("Deleted") = oInspectionSOCInfoLocal.Deleted
        '            dr("Created By") = oInspectionSOCInfoLocal.CreatedBy
        '            dr("Date Created") = oInspectionSOCInfoLocal.CreatedOn
        '            dr("Last Edited By") = oInspectionSOCInfoLocal.ModifiedBy
        '            dr("Date Last Edited") = oInspectionSOCInfoLocal.ModifiedOn
        '            tbEntityTable.Rows.Add(dr)
        '        Next
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub InspectionSOCInfoChanged(ByVal bolValue As Boolean) Handles oInspectionSOCInfo.evtInspectionSOCInfoChanged
            RaiseEvent evtInspectionSOCChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
