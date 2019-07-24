'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Entity
'   Provides the info and collection objects to the client for manipulating
'    an EntityInfo object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         JVC2    11/19/04    Original class definition.
'   1.1         AN      01/03/05    Added Try catch and Exception Handling/Logging
'   1.2         JC      01/03/05    
'   1.3         JVC2    02/02/05    Added overloaded NEW function to populate based on 
'                                       entity name passed to NEW.  Also added exposure
'                                       of ID attribute (readonly).  Also overloaded NEW
'                                       function to populate based on entity type ID.
'   1.4         JVC2    02/07/05    Added readonly property Name.
'
' Function          Description
' GetEntity(NAME)   Returns the Entity requested by the string arg NAME
' GetEntity(ID)     Returns the Entity requested by the int arg ID
' GetEntityAll()    Returns an EntityCollection with all Entity objects
' Add(ID)           Adds the Entity identified by arg ID to the 
'                           internal EntityCollection
' Add(Name)         Adds the Entity identified by arg NAME to the internal 
'                           EntityCollection
' Add(Entity)       Adds the Entity passed as the argument to the internal 
'                           EntityCollection
' Remove(ID)        Removes the Entity identified by arg ID from the internal 
'                           EntityCollection
' Remove(NAME)      Removes the Entity identified by arg NAME from the 
'                           internal EntityCollection
' EntityTable()     Returns a datatable containing all columns for the Entity 
'                           objects in the internal EntityCollection.
' EntityCombo()     Returns a two-column datatable containing Name and ID for 
'                           the Entity objects in the internal EntityCollection.
' colIsDirty()      Returns a boolean indicating whether any of the EntityInfo
'                    objects in the EntityCollection has been modified since the
'                    last time it was retrieved from/saved to the repository.
' Flush()           Marshalls all modified/added EntityInfo objects in the 
'                    EntityCollection to the repository.
' Save()            Marshalls the internal EntityInfo object to the repository.
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pEntity
#Region "Private Member Variables"
        Private colEntities As Muster.Info.EntityCollection
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private oEntityInfo As Muster.Info.EntityInfo
        Private oEntityDB As New Muster.DataAccess.EntityDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Public Sub New()
            oEntityInfo = New Muster.Info.EntityInfo
            colEntities = New Muster.Info.EntityCollection
        End Sub
        Public Sub New(ByVal EntName As String)
            colEntities = New Muster.Info.EntityCollection
            oEntityInfo = Me.GetEntity(EntName)
        End Sub
        Public Sub New(ByVal EntTypeID As Int64)
            colEntities = New MUSTER.Info.EntityCollection
            oEntityInfo = Me.GetEntity(EntTypeID)
        End Sub
#End Region
#Region "Exposed Attributes"
        Public ReadOnly Property ID() As Integer
            Get
                Return oEntityInfo.ID
            End Get
        End Property
        Public ReadOnly Property Name() As String
            Get
                Return oEntityInfo.Name
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Sub Load(ByVal ds As DataSet)
            Dim dr As DataRow
            Try
                For Each dr In ds.Tables("Entity").Rows
                    oEntityInfo = New MUSTER.Info.EntityInfo(dr)
                    colEntities.Add(oEntityInfo)
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        ' Obtains and returns an entity as called for by name
        Public Function GetEntity(ByVal Name As String) As MUSTER.Info.EntityInfo
            Try
                If colEntities.Contains(Name) Then
                    oEntityInfo = colEntities.Item(Name)
                    Return oEntityInfo
                Else
                    colEntities.Add(oEntityDB.DBGetByName(Name))
                    Return GetEntity(Name)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        'Obtains and returns an entity as called for by ID
        Public Function GetEntity(ByVal ID As Int64) As MUSTER.Info.EntityInfo

            Dim oTempInfo As MUSTER.Info.EntityInfo

            Try
                For Each oTempInfo In colEntities.Values
                    If oTempInfo.ID = ID Then
                        oEntityInfo = oTempInfo
                        Return oEntityInfo
                    End If
                Next
                colEntities.Add(oEntityDB.DBGetByID(ID))
                Return GetEntity(ID)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
#End Region
#Region "Collection Operations"
        Function GetEntityAll() As MUSTER.Info.EntityCollection
            Try
                colEntities.Clear()
                colEntities = oEntityDB.GetAllInfo()
                oEntityInfo = colEntities(colEntities.GetKeys(0))
                Return colEntities
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Int64)
            Try
                oEntityInfo = oEntityDB.DBGetByID(ID)
                colEntities.Add(oEntityInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Adds an entity to the collection as called for by Name
        Public Sub Add(ByVal Name As String)
            Try
                oEntityInfo = oEntityDB.DBGetByName(Name)
                colEntities.Add(oEntityInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oEntity As MUSTER.Info.EntityInfo)

            Try
                oEntityInfo = oEntity
                colEntities.Add(oEntityInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Int64)

            Dim myIndex As Int16 = 1

            Try
                For Each oEntityInfo In colEntities.Values
                    If oEntityInfo.ID = ID Then
                        colEntities.Remove(oEntityInfo)
                        oEntityInfo = New MUSTER.Info.EntityInfo
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            Throw New Exception("Entity " & ID.ToString & " is not in the collection of entities.")

        End Sub
        'Removes the entity called for by Name from the collection
        Public Sub Remove(ByVal Name As String)
            Dim myIndex As Int16 = 1

            Try
                For Each oEntityInfo In colEntities
                    If oEntityInfo.Name = Name Then
                        colEntities.Remove(oEntityInfo)
                        oEntityInfo = New MUSTER.Info.EntityInfo
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            Throw New Exception("Entity " & Name & " is not in the collection of entities.")

        End Sub
        Public Function colIsDirty() As Boolean
            Dim oTempInfo As MUSTER.Info.EntityInfo

            For Each oTempInfo In colEntities.Values
                If oTempInfo.IsDirty Then
                    Return True
                End If
            Next

            Return False

        End Function
        Public Sub Flush()
            Dim oTempInfo As MUSTER.Info.EntityInfo

            For Each oTempInfo In colEntities.Values
                If oTempInfo.IsDirty Then
                    oEntityInfo = oTempInfo
                    Me.Save()
                End If
            Next

        End Sub
        Public Sub Save()
            Try
                oEntityDB.Put(oEntityInfo)
                oEntityInfo.IsDirty = False
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#Region "Miscellaneous Operations"
        'Returns a datatable of the entities in the collection
        Public Function EntityTable() As DataTable

            Dim oEntityInfoLocal As MUSTER.Info.EntityInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable

            Try
                tbEntityTable.Columns.Add("Entity ID")
                tbEntityTable.Columns.Add("Entity Name")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Created On")
                tbEntityTable.Columns.Add("Modified By")
                tbEntityTable.Columns.Add("Modified On")

                For Each oEntityInfoLocal In colEntities.Values
                    dr = tbEntityTable.NewRow()
                    dr("Entity ID") = oEntityInfoLocal.ID
                    dr("Entity Name") = oEntityInfoLocal.Name
                    dr("Created By") = oEntityInfoLocal.CreatedBy
                    dr("Created On") = oEntityInfoLocal.CreatedOn
                    dr("Modified By") = oEntityInfoLocal.ModifiedBy
                    dr("Modified On") = oEntityInfoLocal.ModifiedOn
                    tbEntityTable.Rows.Add(dr)
                Next
                Return tbEntityTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        'Returns a two-column datatable of the entities in the collection column names Entity ID and Entity Name
        Public Function EntityCombo() As DataTable

            Dim oEntityInfoLocal As MUSTER.Info.EntityInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable

            Try
                tbEntityTable.Columns.Add("Entity ID")
                tbEntityTable.Columns.Add("Entity Name")

                For Each oEntityInfoLocal In colEntities.Values
                    dr = tbEntityTable.NewRow()
                    dr("Entity ID") = oEntityInfoLocal.ID
                    dr("Entity Name") = oEntityInfoLocal.Name
                    tbEntityTable.Rows.Add(dr)
                Next
                Return tbEntityTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
#End Region
#End Region
    End Class
End Namespace
