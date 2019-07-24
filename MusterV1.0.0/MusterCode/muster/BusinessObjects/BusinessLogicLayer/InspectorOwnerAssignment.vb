'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.InspectorOwnerAssignment
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      RFF/KKM     06/21/2005  Original class definition
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
' NOTE: This file to be used as InspectorOwnerAssignment to build other objects.
'       Replace keyword "InspectorOwnerAssignment" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pInspectorOwnerAssignment
#Region "Public Events"
        Public Event InspectorOwnerAssignmentErr(ByVal MsgStr As String)
        Public Event InspectorOwnerAssignmentChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private WithEvents oInspectorOwnerAssignmentInfo As Muster.Info.InspectorOwnerAssignmentInfo
        Private WithEvents colInspectorOwnerAssignments As Muster.Info.InspectorOwnerAssignmentsCollection
        Private oInspectorOwnerAssignmentDB As New Muster.DataAccess.InspectorOwnerAssignmentDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
#End Region
#Region "Constructors"
        Public Sub New()
            oInspectorOwnerAssignmentInfo = New MUSTER.Info.InspectorOwnerAssignmentInfo
            colInspectorOwnerAssignments = New MUSTER.Info.InspectorOwnerAssignmentsCollection
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property colInspectorOwners() As MUSTER.Info.InspectorOwnerAssignmentsCollection
            Get
                Return colInspectorOwnerAssignments
            End Get
            Set(ByVal Value As MUSTER.Info.InspectorOwnerAssignmentsCollection)
                colInspectorOwnerAssignments = Value
            End Set
        End Property
        Public Property ID() As Integer
            Get
                Return oInspectorOwnerAssignmentInfo.ID
            End Get
            Set(ByVal Value As Integer)
                oInspectorOwnerAssignmentInfo.ID = Integer.Parse(Value)
            End Set
        End Property
        Public Property STAFF_ID() As Integer
            Get
                Return oInspectorOwnerAssignmentInfo.STAFF_ID
            End Get
            Set(ByVal Value As Integer)
                oInspectorOwnerAssignmentInfo.STAFF_ID = Value
            End Set
        End Property
        Public Property OWNER_ID() As Integer
            Get
                Return oInspectorOwnerAssignmentInfo.OWNER_ID
            End Get
            Set(ByVal value As Integer)
                oInspectorOwnerAssignmentInfo.OWNER_ID = value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oInspectorOwnerAssignmentInfo.DELETED
            End Get
            Set(ByVal Value As Boolean)
                oInspectorOwnerAssignmentInfo.DELETED = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oInspectorOwnerAssignmentInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oInspectorOwnerAssignmentInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xInspectorOwnerAssignmentinfo As MUSTER.Info.InspectorOwnerAssignmentInfo
                For Each xInspectorOwnerAssignmentinfo In colInspectorOwnerAssignments.Values
                    If xInspectorOwnerAssignmentinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oInspectorOwnerAssignmentInfo.IsDirty = Value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oInspectorOwnerAssignmentInfo.CREATED_BY
            End Get
            Set(ByVal Value As String)
                oInspectorOwnerAssignmentInfo.CREATED_BY = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oInspectorOwnerAssignmentInfo.DATE_CREATED
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oInspectorOwnerAssignmentInfo.LAST_EDITED_BY
            End Get
            Set(ByVal Value As String)
                oInspectorOwnerAssignmentInfo.LAST_EDITED_BY = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oInspectorOwnerAssignmentInfo.DATE_LAST_EDITED
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As MUSTER.Info.InspectorOwnerAssignmentInfo
            Try
                oInspectorOwnerAssignmentInfo = colInspectorOwnerAssignments.Item(ID)
                If Not oInspectorOwnerAssignmentInfo Is Nothing Then
                    Return oInspectorOwnerAssignmentInfo
                End If
                oInspectorOwnerAssignmentInfo = oInspectorOwnerAssignmentDB.DBGetByID(ID)
                If oInspectorOwnerAssignmentInfo.ID = 0 Then
                    oInspectorOwnerAssignmentInfo.ID = nID
                    nID -= 1
                End If
                colInspectorOwnerAssignments.Add(oInspectorOwnerAssignmentInfo)
                Return oInspectorOwnerAssignmentInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Saves the data in the current Info object
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False) As Boolean
            Try
                If Not bolValidated Then
                    If Not Me.ValidateData() Then
                        Return False
                    End If
                End If
                If Not (oInspectorOwnerAssignmentInfo.ID < 0 And oInspectorOwnerAssignmentInfo.DELETED) Then
                    Dim OldKey As String = oInspectorOwnerAssignmentInfo.ID.ToString
                    oInspectorOwnerAssignmentDB.Put(oInspectorOwnerAssignmentInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Function
                    End If

                    If Not bolValidated Then
                        If oInspectorOwnerAssignmentInfo.ID.ToString <> OldKey Then
                            colInspectorOwnerAssignments.ChangeKey(OldKey, oInspectorOwnerAssignmentInfo.ID.ToString)
                        End If
                    End If
                    oInspectorOwnerAssignmentInfo.Archive()
                    oInspectorOwnerAssignmentInfo.IsDirty = False
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Registration") As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = False
            Try

                If errStr.Length > 0 Or Not validateSuccess Then
                    RaiseEvent InspectorOwnerAssignmentErr(errStr)
                End If
                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        'Gets all the info
        Public Sub GetAll(ByVal inspID As Integer)
            Try
                colInspectorOwnerAssignments.Clear()
                colInspectorOwnerAssignments = oInspectorOwnerAssignmentDB.GetAllInfo(inspID)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oInspectorOwnerAssignmentInfo = oInspectorOwnerAssignmentDB.DBGetByID(ID)
                If oInspectorOwnerAssignmentInfo.ID <= 0 Then
                    oInspectorOwnerAssignmentInfo.ID = nID
                    nID -= 1
                End If
                colInspectorOwnerAssignments.Add(oInspectorOwnerAssignmentInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Function Add(ByRef oInspectorOwnerAssignment As MUSTER.Info.InspectorOwnerAssignmentInfo) As Integer
            Try
                oInspectorOwnerAssignmentInfo = oInspectorOwnerAssignment
                If oInspectorOwnerAssignmentInfo.ID <= 0 Then
                    oInspectorOwnerAssignmentInfo.ID = nID
                    nID -= 1
                End If
                colInspectorOwnerAssignments.Add(oInspectorOwnerAssignmentInfo)
                Return oInspectorOwnerAssignmentInfo.ID
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oInspectorOwnerAssignmentInfoLocal As MUSTER.Info.InspectorOwnerAssignmentInfo
            Try
                For Each oInspectorOwnerAssignmentInfoLocal In colInspectorOwnerAssignments.Values
                    If oInspectorOwnerAssignmentInfoLocal.ID = ID Then
                        colInspectorOwnerAssignments.Remove(oInspectorOwnerAssignmentInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("InspectorOwnerAssignment " & ID.ToString & " is not in the collection of InspectorOwnerAssignments.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oInspectorOwnerAssignment As MUSTER.Info.InspectorOwnerAssignmentInfo)
            Try
                colInspectorOwnerAssignments.Remove(oInspectorOwnerAssignment)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("InspectorOwnerAssignment " & oInspectorOwnerAssignment.ID & " is not in the collection of InspectorOwnerAssignments.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xInspectorOwnerAssignmentInfo As MUSTER.Info.InspectorOwnerAssignmentInfo
            For Each xInspectorOwnerAssignmentInfo In colInspectorOwnerAssignments.Values
                If xInspectorOwnerAssignmentInfo.IsDirty Then
                    oInspectorOwnerAssignmentInfo = xInspectorOwnerAssignmentInfo
                    Me.Save(moduleID, staffID, returnVal, True)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                End If
            Next
        End Sub
#End Region
#Region "General Operations"
        Public Sub Clear()
            oInspectorOwnerAssignmentInfo = New MUSTER.Info.InspectorOwnerAssignmentInfo
        End Sub
        Public Sub Reset()
            oInspectorOwnerAssignmentInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function getAvailableOwnerFacilities(ByVal inspID As Integer) As DataTable
            Try
                Dim dtReturn As DataTable = oInspectorOwnerAssignmentDB.DBGetAvailableOwnerFacilities(inspID).Tables(0)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function EntityTable() As DataTable
            Dim oInspectorOwnerAssignmentInfoLocal As New MUSTER.Info.InspectorOwnerAssignmentInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("OWNER")
                tbEntityTable.Columns.Add("OWNER_ID")
                tbEntityTable.Columns.Add("FACILITIES")
                tbEntityTable.Columns.Add("ID")
                For Each oInspectorOwnerAssignmentInfoLocal In colInspectorOwnerAssignments.Values
                    dr = tbEntityTable.NewRow()
                    dr("ID") = oInspectorOwnerAssignmentInfoLocal.ID
                    dr("OWNER") = oInspectorOwnerAssignmentInfoLocal.Owner
                    dr("OWNER_ID") = oInspectorOwnerAssignmentInfoLocal.OWNER_ID
                    dr("FACILITIES") = oInspectorOwnerAssignmentInfoLocal.Facilities
                    tbEntityTable.Rows.Add(dr)
                Next
                Return tbEntityTable
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function getCNEManagers() As DataTable
            Try
                Return oInspectorOwnerAssignmentDB.DBGetCNEmanagers
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function

        Public Function getOwnersInConflictOfManagerTerritory(Optional ByVal managerID As Integer = 0) As DataTable
            Try
                Return oInspectorOwnerAssignmentDB.DBgetOwnersInConflictOfManagerTerritory(managerID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function

        Public Function getInspectorsUnderManager(Optional ByVal managerID As Integer = 0) As DataTable
            Try
                Return oInspectorOwnerAssignmentDB.DBgetInspectorsUnderManager(managerID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function

        Public Function getAvailableInspectorsForManagement() As DataTable
            Try
                Return oInspectorOwnerAssignmentDB.DBgetAvailableInspectorsForManagement
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function


        Public Sub AssignInspectorToManager(ByVal head_staff_id As Integer, ByVal inspector_id As Integer)
            Try
                oInspectorOwnerAssignmentDB.DBSetAssignmentToCNEmanager(head_staff_id, inspector_id, 0)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Sub



        Public Sub RemoveInspectorFromManager(ByVal head_staff_id As Integer, ByVal inspector_id As Integer)
            Try
                oInspectorOwnerAssignmentDB.DBSetAssignmentToCNEManager(head_staff_id, inspector_id, 1)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Sub


        Public Function getOwnerFacilities() As DataSet
            Try
                Return oInspectorOwnerAssignmentDB.DBGetOwnerFacilities()
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region

#End Region
#Region "External Event Handlers"
        Private Sub InspectorOwnerAssignmentInfoChanged(ByVal bolValue As Boolean) Handles oInspectorOwnerAssignmentInfo.InspectorOwnerAssignmentInfoChanged
            RaiseEvent InspectorOwnerAssignmentChanged(bolValue)
        End Sub
        Private Sub InspectorOwnerAssignmentColChanged(ByVal bolValue As Boolean) Handles colInspectorOwnerAssignments.InspectorOwnerAssignmentColChanged
            RaiseEvent ColChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
