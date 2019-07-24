'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Flag
'   Provides the operations required to manipulate an Flag object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         MNR     12/10/04    Original class definition.
'   1.1         AN      01/03/05    Added Try catch and Exception Handling/Logging
'   1.2         MNR     01/05/05    Added IsDirty() Property
'   1.3         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   1.4         JVC2    01/26/2005  Added SourceUserID attribute as pass-thru to info object.
'   1.5         JVC2    01/31/2005  Added call to ChangeKey if ID changes during save.
'                                       Also added Source User ID to EntityTable()
'   1.6         JVC2    02/02/05    Added EntityTypeID to private members and initialize to "Flag" type.
'                                       Also added overloaded retrieve which takes the entity ID,
'                                       entity type ID and text to be matched to retrieve the FIRST
'                                       non-deleted flag from the collection.
'   1.7         JVC2    02/07/05    Altered EntityTable to add the entity name and to instantiate
'                                       a local entity object so that multiple rows result in minimal
'                                       database hits.
'   1.8         JVC2    02/14/05    Added operation FindFromCalendarID which retrieves a flag
'                                       based on its associated calendar ID
'                                   Added FlagInfo attribute that returns current FlagInfo object
'   1.9         AB      02/18/05    Added DataAge check to the Retrieve function
'
'
' Function          Description
' Retrieve(ID)      Returns an Info Object requested by the int arg ID
' Save()            Saves the Info Object
' GetAll()          Returns a collection with all the relevant information
' Add(ID)           Adds an Info Object identified by the int arg ID
'                   to the Flags Collection
' Add(Entity)       Adds the Entity passed as an argument
'                   to the Flags Collection
' Remove(ID)        Removes an Info Object identified by the int arg ID
'                   from the Flags Collection
' Remove(Entity)    Removes the Entity passed as an argument
'                   from the Flags Collection
' Flush()           Marshalls all modified/new Onwer Info objects in the
'                   Flag Collection to the repository
' EntityTable()     Returns a datatable containing all columns for the Entity
'                   objects in the Flags Collection
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
        Public Class pFlag
#Region "Public Events"
        Public Event FlagErr(ByVal MsgStr As String)
#End Region
#Region "Private Member Variables"
        Private oFlagInfo As New Muster.Info.FlagInfo
        Private colFlags As New Muster.Info.FlagsCollection
        Private oFlagDB As New Muster.DataAccess.FlagDB
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Flag").ID
#End Region
#Region "Constructors"
        Public Sub New()
            oFlagInfo = New Muster.Info.FlagInfo
            colFlags = New Muster.Info.FlagsCollection
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return oFlagInfo.ID
            End Get
            Set(ByVal Value As Integer)
                oFlagInfo.ID = Value
            End Set
        End Property
        Public Property EntityID() As Integer
            Get
                Return oFlagInfo.EntityID
            End Get
            Set(ByVal Value As Integer)
                oFlagInfo.EntityID = Value
            End Set
        End Property
        Public Property EntityType() As Integer
            Get
                Return oFlagInfo.EntityType
            End Get
            Set(ByVal Value As Integer)
                oFlagInfo.EntityType = Value
            End Set
        End Property
        Public Property FlagDescription() As String
            Get
                Return oFlagInfo.FlagDescription
            End Get
            Set(ByVal Value As String)
                oFlagInfo.FlagDescription = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oFlagInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oFlagInfo.Deleted = Value
            End Set
        End Property
        Public Property DueDate() As Date
            Get
                Return oFlagInfo.DueDate
            End Get
            Set(ByVal Value As Date)
                oFlagInfo.DueDate = Value
            End Set
        End Property
        Public Property ModuleID() As String
            Get
                Return oFlagInfo.ModuleID
            End Get
            Set(ByVal Value As String)
                oFlagInfo.ModuleID = Value
            End Set
        End Property
        Public Property CalendarInfoID() As Integer
            Get
                Return oFlagInfo.CalendarInfoID
            End Get
            Set(ByVal Value As Integer)
                oFlagInfo.CalendarInfoID = Value
            End Set
        End Property
        Public Property SourceUserID() As String
            Get
                Return oFlagInfo.SourceUserID
            End Get
            Set(ByVal Value As String)
                oFlagInfo.SourceUserID = Value
            End Set
        End Property
        Public Property FlagColor() As String
            Get
                Return oFlagInfo.FlagColor
            End Get
            Set(ByVal Value As String)
                oFlagInfo.FlagColor = Value
            End Set
        End Property
        Public Property TurnsRedOn() As Date
            Get
                Return oFlagInfo.TurnsRedOn
            End Get
            Set(ByVal Value As Date)
                oFlagInfo.TurnsRedOn = Value
            End Set
        End Property
        Public Property FlagInfo() As MUSTER.Info.FlagInfo
            Get
                Return Me.oFlagInfo
            End Get
            Set(ByVal Value As MUSTER.info.FlagInfo)
                Me.oFlagInfo = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oFlagInfo.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oFlagInfo.IsDirty = Value
            End Set
        End Property
        Public Property FlagsCol() As MUSTER.Info.FlagsCollection
            Get
                Return colFlags
            End Get
            Set(ByVal Value As MUSTER.Info.FlagsCollection)
                colFlags = Value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oFlagInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oFlagInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oFlagInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oFlagInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oFlagInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oFlagInfo.ModifiedOn
            End Get
        End Property

#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As Muster.Info.FlagInfo
            Dim oFlagInfoLocal As MUSTER.Info.FlagInfo
            Dim bolDataAged As Boolean = False

            Try
                For Each oFlagInfoLocal In colFlags.Values
                    If oFlagInfoLocal.ID = ID Then
                        If oFlagInfoLocal.IsDirty = False And oFlagInfoLocal.IsAgedData = True Then
                            bolDataAged = True
                        Else
                            oFlagInfo = oFlagInfoLocal
                            Return oFlagInfo
                        End If
                    End If
                Next
                If bolDataAged = True Then
                    colFlags.Remove(oFlagInfoLocal)
                End If
                oFlagInfo = oFlagDB.DBGetByID(ID)
                If oFlagInfo.ID = 0 Then
                    oFlagInfo.ID = nID
                    nID -= 1
                End If
                colFlags.Add(oFlagInfo)
                Return oFlagInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function Retrieve(ByVal ID As Integer, ByVal EntityTypeID As Integer, ByVal strMatch As String) As Muster.Info.FlagInfo
            Dim oFlagInfoLocal As MUSTER.Info.FlagInfo
            Dim nTempID As Integer
            Dim bolFlagInfoFound As Boolean = False
            Try
                For Each oFlagInfoLocal In colFlags.Values
                    If oFlagInfoLocal.EntityType = EntityTypeID And _
                        oFlagInfoLocal.EntityID = ID And _
                        oFlagInfoLocal.FlagDescription.IndexOf(strMatch) > -1 And _
                        Not oFlagInfoLocal.Deleted Then
                        If oFlagInfoLocal.IsDirty = False And oFlagInfoLocal.IsAgedData = True Then
                            nTempID = oFlagInfoLocal.ID
                        Else
                            oFlagInfo = oFlagInfoLocal
                            bolFlagInfoFound = True
                        End If
                        Exit For
                    End If
                Next
                ' Data was found, but it was old, get a refresh of data and continue
                If nTempID > 0 Then
                    colFlags.Remove(oFlagInfoLocal)
                    oFlagInfoLocal = oFlagDB.DBGetByID(ID)
                    colFlags.Add(oFlagInfoLocal)
                    'Check to ensure the data still meets the proper requirements
                    If oFlagInfoLocal.EntityType = EntityTypeID And _
                                            oFlagInfoLocal.EntityID = ID And _
                                            oFlagInfoLocal.FlagDescription.IndexOf(strMatch) > -1 And _
                                            Not oFlagInfoLocal.Deleted Then
                        oFlagInfo = oFlagInfoLocal
                        bolFlagInfoFound = True
                    End If
                End If

                If bolFlagInfoFound Then
                    Return oFlagInfo
                Else
                    Return Nothing
                End If

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        Public Function RetrieveFlags(Optional ByVal entityID As Integer = 0, Optional ByVal entityType As Integer = 0, Optional ByVal showDeleted As Boolean = False, Optional ByVal [Module] As String = "", Optional ByVal flagID As Integer = 0, Optional ByVal calID As Integer = 0, Optional ByVal userID As String = "", Optional ByVal flagDesc As String = "") As MUSTER.Info.FlagsCollection
            Try
                colFlags.Clear()
                colFlags = oFlagDB.DBGetFlags(entityID, entityType, showDeleted, [Module], flagID, calID, userID, flagDesc)
                Return colFlags
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        'Saves the data in the current Info object
        Public Sub Save(Optional ByVal moduleID As Integer = 0, Optional ByVal staffID As Integer = 0, Optional ByRef returnVal As String = "")
            Try
                Dim OldKey As String = oFlagInfo.ID.ToString
                oFlagDB.Put(oFlagInfo, moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If
                If oFlagInfo.ID.ToString <> OldKey Then
                    colFlags.ChangeKey(OldKey, oFlagInfo.ID.ToString)
                End If
                oFlagInfo.Archive()
                oFlagInfo.IsDirty = False
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function GetFlagsDS(Optional ByVal entityID As Integer = 0, Optional ByVal entityType As Integer = 0, Optional ByVal showDeleted As Boolean = False, Optional ByVal [Module] As String = "", Optional ByVal flagID As Integer = 0, Optional ByVal calID As Integer = 0, Optional ByVal userID As String = "", Optional ByVal flagDesc As String = "") As DataSet
            Try
                Return oFlagDB.DBGetFlagsDS(entityID, entityType, showDeleted, [Module], flagID, calID, userID, flagDesc)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub FindFromCalendarID(ByVal CalendarID As Int64)
            Dim xFlag As MUSTER.Info.FlagInfo
            For Each xFlag In colFlags.Values
                If xFlag.CalendarInfoID = CalendarID Then
                    Me.Retrieve(xFlag.ID)
                    Exit Sub
                End If
            Next
            Me.oFlagInfo = Nothing
        End Sub
        Public Function GetBarometerColors(Optional ByVal entityID As Integer = 0, Optional ByVal entityType As Integer = 0, Optional ByVal eventID As Integer = 0, Optional ByVal eventType As Integer = 0) As DataSet
            Try
                Return oFlagDB.GetBarometerColors(entityID, entityType, eventID, eventType)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        Function GetAll() As MUSTER.Info.FlagsCollection
            Try
                colFlags.Clear()
                colFlags = oFlagDB.GetAllInfo()
                Return colFlags
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oFlagInfo = oFlagDB.DBGetByID(ID)
                If oFlagInfo.ID = 0 Then
                    oFlagInfo.ID = nID
                    nID -= 1
                End If
                colFlags.Add(oFlagInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oFlag As MUSTER.Info.FlagInfo)
            Try
                oFlagInfo = oFlag
                colFlags.Add(oFlagInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oFlagInfoLocal As MUSTER.Info.FlagInfo

            Try
                For Each oFlagInfoLocal In colFlags.Values
                    If oFlagInfoLocal.ID = ID Then
                        colFlags.Remove(oFlagInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'Throw New Exception("Flag " & ID.ToString & " is not in the collection of Flags.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oFlag As MUSTER.Info.FlagInfo)
            Try
                colFlags.Remove(oFlag)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'Throw New Exception("Flag " & oFlag.ID & " is not in the collection of Flags.")
        End Sub
        Public Sub Flush(Optional ByVal moduleID As Integer = 0, Optional ByVal staffID As Integer = 0, Optional ByRef returnVal As String = "")
            Dim xFlagInfo As MUSTER.Info.FlagInfo
            For Each xFlagInfo In colFlags.Values
                If xFlagInfo.IsDirty Then
                    oFlagInfo = xFlagInfo
                    Me.Save(moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                End If
            Next
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colFlags.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colFlags.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colFlags.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oFlagInfo = New MUSTER.Info.FlagInfo
        End Sub
        Public Sub Reset()
            oFlagInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oFlagInfoLocal As New MUSTER.Info.FlagInfo
            Dim oEntityLocal As New MUSTER.BusinessLogic.pEntity
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("Flag ID", Type.GetType("System.Int64"))
                tbEntityTable.Columns.Add("Entity ID", Type.GetType("System.Int64"))
                tbEntityTable.Columns.Add("Entity Type", Type.GetType("System.Int64"))
                tbEntityTable.Columns.Add("Entity Name", Type.GetType("System.String"))
                tbEntityTable.Columns.Add("Flag Desc", Type.GetType("System.String"))
                tbEntityTable.Columns.Add("Deleted", Type.GetType("System.Boolean"))
                tbEntityTable.Columns.Add("Due Date", Type.GetType("System.DateTime"))
                tbEntityTable.Columns.Add("Created By", Type.GetType("System.String"))
                tbEntityTable.Columns.Add("Date Created", Type.GetType("System.DateTime"))
                tbEntityTable.Columns.Add("Last Edited By", Type.GetType("System.String"))
                tbEntityTable.Columns.Add("Date Last Edited", Type.GetType("System.DateTime"))
                tbEntityTable.Columns.Add("Module ID", Type.GetType("System.String"))
                tbEntityTable.Columns.Add("Source User ID", Type.GetType("System.String"))
                tbEntityTable.Columns.Add("Calendar Info ID", Type.GetType("System.Int64"))

                For Each oFlagInfoLocal In colFlags.Values
                    If oFlagInfoLocal.ID > 0 Then
                        dr = tbEntityTable.NewRow()
                        dr("Flag ID") = oFlagInfoLocal.ID
                        dr("Entity ID") = oFlagInfoLocal.EntityID
                        dr("Entity Type") = oFlagInfoLocal.EntityType
                        oEntityLocal.GetEntity(oFlagInfoLocal.EntityType)
                        dr("Entity Name") = oEntityLocal.Name
                        dr("Flag Desc") = oFlagInfoLocal.FlagDescription
                        dr("Deleted") = oFlagInfoLocal.Deleted
                        dr("Due Date") = oFlagInfoLocal.DueDate
                        dr("Created By") = oFlagInfoLocal.CreatedBy
                        dr("Date Created") = oFlagInfoLocal.CreatedOn
                        dr("Last Edited By") = oFlagInfoLocal.ModifiedBy
                        dr("Date Last Edited") = oFlagInfoLocal.ModifiedOn
                        dr("Module ID") = oFlagInfoLocal.ModuleID
                        dr("Calendar Info ID") = oFlagInfoLocal.CalendarInfoID
                        dr("Source User ID") = oFlagInfoLocal.SourceUserID
                        tbEntityTable.Rows.Add(dr)
                    End If
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
