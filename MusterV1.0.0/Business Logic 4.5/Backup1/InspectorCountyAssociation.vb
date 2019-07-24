'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.InspectorCountyAssociation
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0                              Original class definition
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
' NOTE: This file to be used as InspectorCountyAssociation to build other objects.
'       Replace keyword "InspectorCountyAssociation" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pInspectorCountyAssociation
#Region "Public Events"
        Public Event InspectorCountyAssociationErr(ByVal MsgStr As String)
        Public Event InspectorCountyAssociationChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private WithEvents oInspectorCountyAssociationInfo As Muster.Info.InspectorCountyAssociationInfo
        Private WithEvents colInspectorCountyAssociations As Muster.Info.InspectorCountyAssociationsCollection
        Private oInspectorCountyAssociationDB As New Muster.DataAccess.InspectorCountyAssociationDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
#End Region
#Region "Constructors"
        Public Sub New()
            oInspectorCountyAssociationInfo = New MUSTER.Info.InspectorCountyAssociationInfo
            colInspectorCountyAssociations = New MUSTER.Info.InspectorCountyAssociationsCollection
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return oInspectorCountyAssociationInfo.ID
            End Get
            Set(ByVal Value As Integer)
                oInspectorCountyAssociationInfo.ID = Integer.Parse(Value)
            End Set
        End Property
        Public Property STAFF_ID() As Integer
            Get
                Return oInspectorCountyAssociationInfo.STAFF_ID
            End Get
            Set(ByVal Value As Integer)
                oInspectorCountyAssociationInfo.STAFF_ID = Value
            End Set
        End Property
        Public Property FIPS_CODE() As Integer
            Get
                Return oInspectorCountyAssociationInfo.FIPS_CODE
            End Get
            Set(ByVal value As Integer)
                oInspectorCountyAssociationInfo.FIPS_CODE = value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oInspectorCountyAssociationInfo.DELETED
            End Get
            Set(ByVal Value As Boolean)
                oInspectorCountyAssociationInfo.DELETED = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oInspectorCountyAssociationInfo.IsDirty
            End Get
            Set(ByVal value As Boolean)
                oInspectorCountyAssociationInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xInspectorCountyAssignmentinfo As MUSTER.Info.InspectorCountyAssociationInfo
                For Each xInspectorCountyAssignmentinfo In colInspectorCountyAssociations.Values
                    If xInspectorCountyAssignmentinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oInspectorCountyAssociationInfo.IsDirty = Value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oInspectorCountyAssociationInfo.CREATED_BY
            End Get
            Set(ByVal Value As String)
                oInspectorCountyAssociationInfo.CREATED_BY = Value
            End Set
        End Property
        Public Property CreatedOn() As Date
            Get
                Return oInspectorCountyAssociationInfo.DATE_CREATED
            End Get
            Set(ByVal Value As Date)
                oInspectorCountyAssociationInfo.DATE_CREATED = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedBy() As String
            Get
                Return oInspectorCountyAssociationInfo.LAST_EDITED_BY
            End Get
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oInspectorCountyAssociationInfo.DATE_LAST_EDITED
            End Get
        End Property
        Public Property colInspectorCounties() As MUSTER.Info.InspectorCountyAssociationsCollection
            Get
                Return colInspectorCountyAssociations
            End Get
            Set(ByVal Value As MUSTER.Info.InspectorCountyAssociationsCollection)
                colInspectorCountyAssociations = Value
            End Set
        End Property
        Public Property CountyInfo() As MUSTER.Info.InspectorCountyAssociationInfo
            Get
                Return oInspectorCountyAssociationInfo
            End Get
            Set(ByVal Value As MUSTER.Info.InspectorCountyAssociationInfo)
                oInspectorCountyAssociationInfo = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As MUSTER.Info.InspectorCountyAssociationInfo
            Try
                oInspectorCountyAssociationInfo = colInspectorCountyAssociations.Item(ID)
                If Not oInspectorCountyAssociationInfo Is Nothing Then
                    Return oInspectorCountyAssociationInfo
                End If
                oInspectorCountyAssociationInfo = oInspectorCountyAssociationDB.DBGetByID(ID)
                If oInspectorCountyAssociationInfo.ID = 0 Then
                    oInspectorCountyAssociationInfo.ID = nID
                    nID -= 1
                End If
                colInspectorCountyAssociations.Add(oInspectorCountyAssociationInfo)
                Return oInspectorCountyAssociationInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Saves the data in the current Info object
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False) As Boolean
            Try
                If Not bolValidated Then
                    If Me.ValidateData() Then
                        Return False
                    End If
                End If
                If Not (oInspectorCountyAssociationInfo.ID < 0 And oInspectorCountyAssociationInfo.DELETED) Then
                    Dim OldKey As String = oInspectorCountyAssociationInfo.ID.ToString
                    oInspectorCountyAssociationDB.Put(oInspectorCountyAssociationInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Function
                    End If

                    If Not bolValidated Then
                        If oInspectorCountyAssociationInfo.ID.ToString <> OldKey Then
                            colInspectorCountyAssociations.ChangeKey(OldKey, oInspectorCountyAssociationInfo.ID.ToString)
                        End If
                    End If
                    oInspectorCountyAssociationInfo.Archive()
                    oInspectorCountyAssociationInfo.IsDirty = False
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Saves the data into the database
        Public Function fillCollection(ByVal dt As DataTable) As Boolean
            Dim dr As DataRow
            Dim nID As Integer = -1
            Dim oCountyInfo As MUSTER.Info.InspectorCountyAssociationInfo
            Try
                For Each dr In dt.Rows
                    oCountyInfo.FIPS_CODE = dr("COUNTY")
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        'Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Registration") As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = False
            Try
                If errStr.Length > 0 Or Not validateSuccess Then
                    RaiseEvent InspectorCountyAssociationErr(errStr)
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
        Public Sub GetAll(Optional ByVal InspID As Integer = Nothing)
            Try
                colInspectorCountyAssociations.Clear()
                colInspectorCountyAssociations = oInspectorCountyAssociationDB.GetbyInspectorID(InspID)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oInspectorCountyAssociationInfo = oInspectorCountyAssociationDB.DBGetByID(ID)
                If oInspectorCountyAssociationInfo.ID <= 0 Then
                    oInspectorCountyAssociationInfo.ID = nID
                    nID -= 1
                End If
                colInspectorCountyAssociations.Add(oInspectorCountyAssociationInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Function Add(ByRef oInspectorCountyAssociation As MUSTER.Info.InspectorCountyAssociationInfo) As Integer
            Try
                oInspectorCountyAssociationInfo = oInspectorCountyAssociation
                If oInspectorCountyAssociationInfo.ID <= 0 Then
                    oInspectorCountyAssociationInfo.ID = nID
                    nID -= 1
                End If
                colInspectorCountyAssociations.Add(oInspectorCountyAssociationInfo)
                Return oInspectorCountyAssociationInfo.ID
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oInspectorCountyAssociationInfoLocal As MUSTER.Info.InspectorCountyAssociationInfo

            Try
                For Each oInspectorCountyAssociationInfoLocal In colInspectorCountyAssociations.Values
                    If oInspectorCountyAssociationInfoLocal.ID = ID Then
                        colInspectorCountyAssociations.Remove(oInspectorCountyAssociationInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("InspectorCountyAssociation " & ID.ToString & " is not in the collection of InspectorCountyAssociations.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oInspectorCountyAssociation As MUSTER.Info.InspectorCountyAssociationInfo)
            Try
                colInspectorCountyAssociations.Remove(oInspectorCountyAssociation)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("InspectorCountyAssociation " & oInspectorCountyAssociation.ID & " is not in the collection of InspectorCountyAssociations.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xInspectorCountyAssociationInfo As MUSTER.Info.InspectorCountyAssociationInfo
            For Each xInspectorCountyAssociationInfo In colInspectorCountyAssociations.Values
                If xInspectorCountyAssociationInfo.IsDirty Then
                    oInspectorCountyAssociationInfo = xInspectorCountyAssociationInfo
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
            oInspectorCountyAssociationInfo = New MUSTER.Info.InspectorCountyAssociationInfo
        End Sub
        Public Sub Reset()
            oInspectorCountyAssociationInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oCountyInfoLocal As New MUSTER.Info.InspectorCountyAssociationInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("COUNTY")
                tbEntityTable.Columns.Add("FIPS")
                tbEntityTable.Columns.Add("FACILITIES")
                tbEntityTable.Columns.Add("ID")

                For Each oCountyInfoLocal In colInspectorCountyAssociations.Values
                    dr = tbEntityTable.NewRow()
                    dr("COUNTY") = oCountyInfoLocal.County
                    dr("FIPS") = oCountyInfoLocal.FIPS_CODE
                    dr("FACILITIES") = oCountyInfoLocal.Facilities
                    dr("ID") = oCountyInfoLocal.ID
                    tbEntityTable.Rows.Add(dr)
                Next
                Return tbEntityTable
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function getAvailableCountyFacilities(ByVal inspID As Integer) As DataTable
            Try
                Dim dtReturn As DataTable = oInspectorCountyAssociationDB.DBGetAvailableCountyFacilities(inspID).Tables(0)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function getInspectors() As DataTable
            Dim dtTable As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vCAEInspectors")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        'Public Function getFacilityCount(ByVal inspID As Integer) As Integer
        '    Try
        '        Return CInt(oInspectorCountyAssociationDB.DBGetTotalCountyOwnerFacilities(inspID).Tables(0).Rows.Item(0).Item(0))
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function
#End Region
#Region "LookUp Operations"
        Private Function GetDataTable(ByVal DBViewName As String) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Try
                strSQL = "SELECT * FROM " & DBViewName

                dsReturn = oInspectorCountyAssociationDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                Else
                    dtReturn = Nothing
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


#End Region
#End Region
#Region "External Event Handlers"
        Private Sub InspectorCountyAssociationInfoChanged(ByVal bolValue As Boolean) Handles oInspectorCountyAssociationInfo.InspectorCountyAssociationInfoChanged
            RaiseEvent InspectorCountyAssociationChanged(bolValue)
        End Sub
        Private Sub InspectorCountyAssociationColChanged(ByVal bolValue As Boolean) Handles colInspectorCountyAssociations.InspectorCountyAssociationColChanged
            RaiseEvent ColChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
