'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.InspectionCCAT
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
' NOTE: This file to be used as InspectionCCAT to build other objects.
'       Replace keyword "InspectionCCAT" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pInspectionCCAT
#Region "Public Events"
        Public Event evtInspectionCCATErr(ByVal MsgStr As String)
        Public Event evtInspectionCCATChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private oInspection As MUSTER.Info.InspectionInfo
        Private WithEvents oInspectionCCATInfo As MUSTER.Info.InspectionCCATInfo
        Private oInspectionCCATDB As MUSTER.DataAccess.InspectionCCATDB
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

            oInspectionCCATInfo = New MUSTER.Info.InspectionCCATInfo
            oInspectionCCATDB = New MUSTER.DataAccess.InspectionCCATDB
        End Sub
#End Region
#Region "Exposed Attributes"

        Public Property FirstCompartment() As Boolean
            Get
                Return oInspectionCCATInfo.FirstCompartment
            End Get
            Set(ByVal Value As Boolean)
                oInspectionCCATInfo.FirstCompartment = Value
            End Set
        End Property

        Public Property CompartmentID() As Integer
            Get
                Return oInspectionCCATInfo.CompartmentID
            End Get

            Set(ByVal Value As Integer)
                oInspectionCCATInfo.CompartmentID = Value
            End Set
        End Property

        Public ReadOnly Property ID() As Integer
            Get
                Return oInspectionCCATInfo.ID
            End Get
        End Property
        Public Property InspectionID() As Integer
            Get
                Return oInspectionCCATInfo.InspectionID
            End Get
            Set(ByVal Value As Integer)
                oInspectionCCATInfo.InspectionID = Value
            End Set
        End Property
        Public Property QuestionID() As Integer
            Get
                Return oInspectionCCATInfo.QuestionID
            End Get
            Set(ByVal Value As Integer)
                oInspectionCCATInfo.QuestionID = Value
            End Set
        End Property
        Public Property TankPipeID() As Integer
            Get
                Return oInspectionCCATInfo.TankPipeID
            End Get
            Set(ByVal Value As Integer)
                oInspectionCCATInfo.TankPipeID = Value
            End Set
        End Property
        Public Property TankPipeEntityID() As Integer
            Get
                Return oInspectionCCATInfo.TankPipeEntityID
            End Get
            Set(ByVal Value As Integer)
                oInspectionCCATInfo.TankPipeEntityID = Value
            End Set
        End Property
        Public Property TankPipeResponse() As Boolean
            Get
                Return oInspectionCCATInfo.TankPipeResponse
            End Get
            Set(ByVal Value As Boolean)
                oInspectionCCATInfo.TankPipeResponse = Value
            End Set
        End Property
        Public Property Termination() As Boolean
            Get
                Return oInspectionCCATInfo.Termination
            End Get
            Set(ByVal Value As Boolean)
                oInspectionCCATInfo.Termination = Value
            End Set
        End Property
        Public Property TankPipeResponseDetail() As String
            Get
                Return oInspectionCCATInfo.TankPipeResponseDetail
            End Get
            Set(ByVal Value As String)
                oInspectionCCATInfo.TankPipeResponseDetail = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oInspectionCCATInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oInspectionCCATInfo.IsDirty = value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oInspectionCCATInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oInspectionCCATInfo.Deleted = Value
            End Set
        End Property
        Public ReadOnly Property colIsDirty() As Boolean
            Get
                Dim xInspectionCCATinfo As MUSTER.Info.InspectionCCATInfo
                For Each xInspectionCCATinfo In oInspection.CCATsCollection.Values
                    If xInspectionCCATinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
        End Property
        Public Property CreatedBy() As String
            Get
                Return oInspectionCCATInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oInspectionCCATInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oInspectionCCATInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oInspectionCCATInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oInspectionCCATInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oInspectionCCATInfo.ModifiedOn
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
                If ds.Tables("CCAT").Rows.Count > 0 Then
                    For Each dr In ds.Tables("CCAT").Rows
                        oInspectionCCATInfo = New MUSTER.Info.InspectionCCATInfo(dr)
                        oInspection.CCATsCollection.Add(oInspectionCCATInfo)
                    Next
                End If
                ds.Tables.Remove("CCAT")
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Function GetCCATTankPipeTermListForInspection(ByVal inspID As Integer, Optional ByVal facilityID As Integer = 0) As DataTable
            Try
                Return oInspectionCCATDB.DBGetCCATListForInspection(inspID, facilityID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function
        Public Function Retrieve(ByRef inspection As MUSTER.Info.InspectionInfo, ByVal id As Integer) As MUSTER.Info.InspectionCCATInfo
            Try
                oInspection = inspection
                oInspectionCCATInfo = oInspection.CCATsCollection.Item(id)
                If oInspectionCCATInfo Is Nothing Then
                    Add(id)
                End If
                Return oInspectionCCATInfo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function RetrieveByQID(ByRef inspection As MUSTER.Info.InspectionInfo, ByVal qid As Integer) As MUSTER.Info.InspectionCCATInfo
            Try
                oInspection = inspection
                For Each ccat As MUSTER.Info.InspectionCCATInfo In oInspection.CCATsCollection.Values
                    If ccat.QuestionID = qid Then
                        oInspectionCCATInfo = ccat
                    End If
                Next
                If oInspectionCCATInfo Is Nothing Or oInspectionCCATInfo.QuestionID <> qid Then
                    Add(New MUSTER.Info.InspectionCCATInfo)
                    oInspectionCCATInfo.QuestionID = qid
                End If
                Return oInspectionCCATInfo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False) As Boolean
            Dim oldID As Integer
            Try
                If Not bolValidated And Not oInspectionCCATInfo.Deleted And Not bolDelete Then
                    If Not Me.ValidateData() Then
                        Return False
                    End If
                End If
                If Not (oInspectionCCATInfo.ID < 0 And oInspectionCCATInfo.Deleted) Then
                    oldID = oInspectionCCATInfo.ID
                    oInspectionCCATDB.Put(oInspectionCCATInfo, moduleID, staffID, returnVal, FirstCompartment)
                    If Not bolValidated Then
                        If oldID < 0 Then
                            oInspection.CCATsCollection.ChangeKey(oldID, oInspectionCCATInfo.ID)
                        End If
                    End If
                    oInspectionCCATInfo.Archive()
                    oInspectionCCATInfo.IsDirty = False
                End If
                If Not bolValidated And bolDelete Then
                    If oInspectionCCATInfo.Deleted Then
                        ' check if other inspections are present else load new instance
                        Dim strNext As String = Me.GetNext()
                        Dim strPrev As String = Me.GetPrevious()
                        If strNext = oInspectionCCATInfo.ID Then
                            If strPrev = oInspectionCCATInfo.ID Then
                                RaiseEvent evtInspectionCCATErr("Inspection " + oInspectionCCATInfo.ID.ToString + " deleted")
                                oInspection.CCATsCollection.Remove(oInspectionCCATInfo)
                                If bolDelete Then
                                    oInspectionCCATInfo = New MUSTER.Info.InspectionCCATInfo
                                Else
                                    oInspectionCCATInfo = Me.Retrieve(oInspection, 0)
                                End If
                            Else
                                RaiseEvent evtInspectionCCATErr("Inspection " + oInspectionCCATInfo.ID.ToString + " deleted")
                                oInspection.CCATsCollection.Remove(oInspectionCCATInfo)
                                oInspectionCCATInfo = Me.Retrieve(oInspection, strPrev)
                            End If
                        Else
                            RaiseEvent evtInspectionCCATErr("Inspection " + oInspectionCCATInfo.ID.ToString + " deleted")
                            oInspection.CCATsCollection.Remove(oInspectionCCATInfo)
                            oInspectionCCATInfo = Me.Retrieve(oInspection, strNext)
                        End If
                    End If
                End If
                RaiseEvent evtInspectionCCATErr(oInspectionCCATInfo.IsDirty)
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
                oInspectionCCATInfo = oInspectionCCATDB.DBGetByID(id)
                If oInspectionCCATInfo.ID = 0 Then
                    oInspectionCCATInfo.InspectionID = oInspection.ID
                    oInspectionCCATInfo.ID = nID
                    nID -= 1
                End If
                oInspection.CCATsCollection.Add(oInspectionCCATInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oInspectionCCAT As MUSTER.Info.InspectionCCATInfo)
            Try
                oInspectionCCATInfo = oInspectionCCAT
                If oInspectionCCATInfo.ID = 0 Then
                    oInspectionCCATInfo.InspectionID = oInspection.ID
                    oInspectionCCATInfo.ID = nID
                    nID -= 1
                End If
                oInspection.CCATsCollection.Add(oInspectionCCATInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal id As Integer)
            Try
                If oInspection.CCATsCollection.Contains(id) Then
                    oInspection.CCATsCollection.Remove(id)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("InspectionCCAT " & id.ToString & " is not in the collection of InspectionCCATs.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oInspectionCCAT As MUSTER.Info.InspectionCCATInfo)
            Try
                If oInspection.CCATsCollection.Contains(oInspectionCCAT) Then
                    oInspection.CCATsCollection.Remove(oInspectionCCAT)
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
            Dim ccatInfo As MUSTER.Info.InspectionCCATInfo
            Try
                For Each ccatInfo In oInspection.CCATsCollection.Values
                    If ccatInfo.IsDirty Then
                        oInspectionCCATInfo = ccatInfo
                        If oInspectionCCATInfo.Deleted Then
                            If oInspectionCCATInfo.ID < 0 Then
                                delIDs.Add(oInspectionCCATInfo.ID)
                            Else
                                Me.Save(moduleID, staffID, returnVal, True)
                            End If
                        Else
                            If Me.ValidateData Then
                                If oInspectionCCATInfo.ID < 0 Then
                                    IDs.Add(oInspectionCCATInfo.ID)
                                End If
                                Me.Save(moduleID, staffID, returnVal, True)
                            Else : Exit For
                            End If
                        End If
                    End If
                Next
                If Not (delIDs Is Nothing) Then
                    For index = 1 To delIDs.Count
                        ccatInfo = oInspection.CCATsCollection.Item(CType(delIDs.Item(index), String))
                        oInspection.CCATsCollection.Remove(ccatInfo)
                    Next
                End If
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        ccatInfo = oInspection.CCATsCollection.Item(colKey)
                        oInspection.CCATsCollection.ChangeKey(colKey, ccatInfo.ID)
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
            Dim strArr() As String = oInspection.CCATsCollection.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            For Each y As String In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return oInspection.CCATsCollection.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            ElseIf oInspection.CCATsCollection.Count <> 0 Then
                Return oInspection.CCATsCollection.Item(nArr.GetValue(colIndex)).ID.ToString
            Else
                Return "0"
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oInspectionCCATInfo = New MUSTER.Info.InspectionCCATInfo
        End Sub
        Public Sub Reset()
            oInspectionCCATInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oInspectionCCATInfoLocal As New MUSTER.Info.InspectionCCATInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("InspectionCCAT ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oInspectionCCATInfoLocal In oInspection.CCATsCollection.Values
                    dr = tbEntityTable.NewRow()
                    dr("Replace ID") = oInspectionCCATInfoLocal.ID
                    dr("Created By") = oInspectionCCATInfoLocal.CreatedBy
                    dr("Date Created") = oInspectionCCATInfoLocal.CreatedOn
                    dr("Last Edited By") = oInspectionCCATInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oInspectionCCATInfoLocal.ModifiedOn
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
        Private Sub InspectionCCATInfoChanged(ByVal bolValue As Boolean) Handles oInspectionCCATInfo.evtInspectionCCATInfoChanged
            RaiseEvent evtInspectionCCATChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
