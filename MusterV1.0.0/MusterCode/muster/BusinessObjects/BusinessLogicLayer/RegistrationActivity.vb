'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.RegistrationActivity
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0                              Original class definition
'  1.1        AB        02/22/05    Added DataAge check to the Retrieve function
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
' NOTE: This file to be used as RegistrationActivity to build other objects.
'       Replace keyword "RegistrationActivity" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pRegistrationActivity
#Region "Public Events"
        Public Event RegistrationActivityErr(ByVal MsgStr As String)
        Public Event RegistrationActivityChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private WithEvents oRegistrationActivityInfo As MUSTER.Info.RegistrationActivityInfo
        Private WithEvents colRegistrationActivitys As MUSTER.Info.RegistrationActivityCollection
        Private oRegistrationActivityDB As New MUSTER.DataAccess.RegistrationActivityDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
#End Region
#Region "Constructors"
        Public Sub New()
            oRegistrationActivityInfo = New MUSTER.Info.RegistrationActivityInfo
            colRegistrationActivitys = New MUSTER.Info.RegistrationActivityCollection
        End Sub
        '********************************************************
        '
        ' Overloaded NEW which will populate with a single instance
        '   of the named RegistrationActivity object.
        '
        '********************************************************
        Public Sub New(ByVal RegistrationActivityName As String)
            oRegistrationActivityInfo = New MUSTER.Info.RegistrationActivityInfo
            colRegistrationActivitys = New MUSTER.Info.RegistrationActivityCollection
            Me.Retrieve(RegistrationActivityName)
        End Sub
#End Region
#Region "Exposed Attributes"
        ' Gets/Sets the Registration Activity Value
        Public Property ActivityDesc() As Integer
            Get
                Return oRegistrationActivityInfo.ActivityDesc
            End Get
            Set(ByVal Value As Integer)
                oRegistrationActivityInfo.ActivityDesc = Value
            End Set
        End Property
        Public Property RegistrationID() As Integer
            Get
                Return oRegistrationActivityInfo.RegistrationID
            End Get
            Set(ByVal Value As Integer)
                oRegistrationActivityInfo.RegistrationID = Value
            End Set
        End Property
        ' Gets/Sets the Entity Type for the registration activity
        Public Property EntityType() As Integer
            Get
                Return oRegistrationActivityInfo.EntityType
            End Get
            Set(ByVal Value As Integer)
                oRegistrationActivityInfo.EntityType = Value
            End Set
        End Property
        ' Gets/Sets the Entity ID associated with the registration activity
        Public Property EntityId() As Integer
            Get
                Return oRegistrationActivityInfo.EntityId
            End Get
            Set(ByVal Value As Integer)
                oRegistrationActivityInfo.EntityId = Value
            End Set
        End Property
        ' Gets/Sets the User ID associated with the registration activity
        Public Property UserID() As Integer
            Get
                Return oRegistrationActivityInfo.UserID
            End Get
            Set(ByVal Value As Integer)
                oRegistrationActivityInfo.UserID = Value
            End Set
        End Property
        ' Gets/Sets the boolean indicating whether or not the registration activity has been processed
        Public Property Processed() As Boolean
            Get
                Return oRegistrationActivityInfo.Processed
            End Get
            Set(ByVal Value As Boolean)
                oRegistrationActivityInfo.Processed = Value
            End Set
        End Property
        ' Gets/Sets the date on which the registration activity was processed
        Public Property DateAdded() As DateTime
            Get
                Return oRegistrationActivityInfo.DateAdded
            End Get
            Set(ByVal Value As DateTime)
                oRegistrationActivityInfo.DateAdded = Value
            End Set
        End Property

        ' Gets the string indicating the user that created the activity
        Public ReadOnly Property CreatedBy() As String
            Get
                Return oRegistrationActivityInfo.CreatedBy
            End Get
        End Property
        ' Gets the date on which the activity was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oRegistrationActivityInfo.CreatedOn
            End Get
        End Property
        ' Gets/Sets the date on which the registration activity was added
        Public Property DateAdded(ByVal dtDate As Date) As Date
            Get
                Return oRegistrationActivityInfo.DateAdded
            End Get
            Set(ByVal Value As Date)
                oRegistrationActivityInfo.DateAdded = Value
            End Set
        End Property

        ' Gets/Sets the deleted state for the registration activity
        Public Property Deleted() As Boolean
            Get
                Return oRegistrationActivityInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oRegistrationActivityInfo.Deleted = Value
            End Set
        End Property

        ' Gets the dirty state of the info object
        Public ReadOnly Property IsDirty() As Boolean
            Get
                Return oRegistrationActivityInfo.IsDirty
            End Get
        End Property
        ' Gets the string indicating the last user to modify the activity
        Public ReadOnly Property ModifiedBy() As String
            Get
                Return oRegistrationActivityInfo.ModifiedBy
            End Get
        End Property
        ' Gets the date on which the activity was last modified
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oRegistrationActivityInfo.ModifiedOn
            End Get
        End Property


        'Gets/Sets the Registration Action Index
        Public Property RegActionIndex() As Long
            Get
                Return oRegistrationActivityInfo.RegActionIndex
            End Get
            Set(ByVal Value As Long)
                oRegistrationActivityInfo.RegActionIndex = Value
            End Set
        End Property

        Public Property colIsDirty() As Boolean
            Get
                Dim xRegistrationActivityinfo As MUSTER.Info.RegistrationActivityInfo
                For Each xRegistrationActivityinfo In colRegistrationActivitys.Values
                    If xRegistrationActivityinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                'oRegistrationActivityInfo.IsDirty = Value
            End Set
        End Property

        Public Function Values() As ICollection
            Return Me.colRegistrationActivitys.Values
        End Function

        Public Property Col() As MUSTER.Info.RegistrationActivityCollection
            Get
                Return colRegistrationActivitys
            End Get
            Set(ByVal Value As MUSTER.Info.RegistrationActivityCollection)
                colRegistrationActivitys = Value
            End Set
        End Property
        Public Property CalendarID() As Integer
            Get
                Return oRegistrationActivityInfo.CalendarID
            End Get
            Set(ByVal Value As Integer)
                oRegistrationActivityInfo.CalendarID = Value
            End Set
        End Property
        Public Property RegActivityInfo() As MUSTER.Info.RegistrationActivityInfo
            Get
                Return oRegistrationActivityInfo
            End Get
            Set(ByVal Value As MUSTER.Info.RegistrationActivityInfo)
                oRegistrationActivityInfo = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As MUSTER.Info.RegistrationActivityInfo
            Dim oRegistrationActivityInfoLocal As MUSTER.Info.RegistrationActivityInfo
            Dim bolDataAged As Boolean = False
            Try
                For Each oRegistrationActivityInfoLocal In colRegistrationActivitys.Values
                    If oRegistrationActivityInfoLocal.RegActionIndex = ID Then
                        If oRegistrationActivityInfoLocal.IsAgedData = True And oRegistrationActivityInfoLocal.IsDirty = False Then
                            bolDataAged = True
                            Exit For
                        Else
                            oRegistrationActivityInfo = oRegistrationActivityInfoLocal
                            Return oRegistrationActivityInfo
                        End If
                    End If
                Next
                If bolDataAged Then
                    colRegistrationActivitys.Remove(oRegistrationActivityInfo)
                End If
                oRegistrationActivityInfo = oRegistrationActivityDB.DBGetByID(ID)
                If oRegistrationActivityInfo.RegActionIndex = 0 Then
                    oRegistrationActivityInfo.RegActionIndex = nID
                    nID -= 1
                End If
                colRegistrationActivitys.Add(oRegistrationActivityInfo)
                Return oRegistrationActivityInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Retrieve by EntityID
        Public Function Retrieve(ByVal EntityID As Integer, ByVal EntityType As Integer, ByVal Activity As String) As MUSTER.Info.RegistrationActivityInfo
            Dim oRegistrationActivityInfoLocal As MUSTER.Info.RegistrationActivityInfo = Nothing
            Dim bolDataAged As Boolean = False
            Try
                For Each oRegistrationActivityInfoLocal In colRegistrationActivitys.Values
                    If oRegistrationActivityInfoLocal.EntityId = EntityID And _
                         oRegistrationActivityInfoLocal.EntityType = EntityType And _
                         oRegistrationActivityInfoLocal.ActivityDesc = Activity Then
                        If oRegistrationActivityInfoLocal.IsAgedData = True And oRegistrationActivityInfoLocal.IsDirty = False Then
                            bolDataAged = True
                        Else
                            oRegistrationActivityInfo = oRegistrationActivityInfoLocal
                            Return oRegistrationActivityInfo
                        End If
                    End If
                Next
                If bolDataAged Then
                    colRegistrationActivitys.Remove(oRegistrationActivityInfoLocal)
                End If
                oRegistrationActivityInfo = oRegistrationActivityDB.DBGetByEntity(EntityID, EntityType, Activity)
                If Not oRegistrationActivityInfo.RegActionIndex = 0 Then
                    colRegistrationActivitys.Add(oRegistrationActivityInfo)
                End If

                Return oRegistrationActivityInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Saves the data in the current Info object
        Public Sub Save()
            Dim strModuleName As String = String.Empty
            Dim oldID As Integer
            Try
                ' If Me.ValidateData(strModuleName) Then
                oldID = oRegistrationActivityInfo.RegActionIndex
                oRegistrationActivityDB.Put(oRegistrationActivityInfo)
                If oldID <= 0 Then
                    colRegistrationActivitys.ChangeKey(oldID, oRegistrationActivityInfo.RegActionIndex)
                End If
                oRegistrationActivityInfo.Archive()
                ' End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Registration") As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = False
            '********************************************************
            '
            ' Sample validation code below.  Modify as necessary to
            '  handle rules for the object.  Note that errStr should
            '  be built with all failed validation reasons as it is
            '  raised to the consumer for display to the user.  The
            '  same goes for the boolean validateSuccess.  Since
            '  validateSuccess ASSUMES failure, it must be set to
            '  TRUE if all validations are passed successfully.
            '********************************************************
            Try
                Select Case [module]
                    Case "Registration"

                        ' if any validations failed
                        Exit Select
                    Case "Technical"

                        ' if any validations failed
                        Exit Select
                End Select
                If errStr.Length > 0 Or Not validateSuccess Then
                    RaiseEvent RegistrationActivityErr(errStr)
                End If
                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function RetrieveByRegID(ByVal regID As Integer) As MUSTER.Info.RegistrationActivityCollection
            Try
                Return oRegistrationActivityDB.DBGetByRegistration(regID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

#End Region
#Region "Collection Operations"
        'Gets all the info for a registration
        Function GetAllForRegistration(ByVal strval As String) As MUSTER.Info.RegistrationActivityCollection
            Try
                colRegistrationActivitys.Clear()
                colRegistrationActivitys = oRegistrationActivityDB.DBGetByRegistration(strval)
                Return colRegistrationActivitys
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Gets all the info
        Function GetAll() As MUSTER.Info.RegistrationActivityCollection
            Try
                colRegistrationActivitys.Clear()
                colRegistrationActivitys = oRegistrationActivityDB.GetAllInfo
                Return colRegistrationActivitys
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oRegistrationActivityInfo = oRegistrationActivityDB.DBGetByID(ID)
                If oRegistrationActivityInfo.RegActionIndex = 0 Then
                    oRegistrationActivityInfo.RegActionIndex = nID
                    nID -= 1
                End If
                colRegistrationActivitys.Add(oRegistrationActivityInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oRegistrationActivity As MUSTER.Info.RegistrationActivityInfo)
            Try
                oRegistrationActivityInfo = oRegistrationActivity
                If oRegistrationActivityInfo.RegActionIndex = 0 Then
                    oRegistrationActivityInfo.RegActionIndex = nID
                    nID -= 1
                End If
                colRegistrationActivitys.Add(oRegistrationActivityInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oRegistrationActivityInfoLocal As MUSTER.Info.RegistrationActivityInfo

            Try
                For Each oRegistrationActivityInfoLocal In colRegistrationActivitys.Values
                    If oRegistrationActivityInfoLocal.RegActionIndex = ID Then
                        colRegistrationActivitys.Remove(oRegistrationActivityInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Pipe " & ID.ToString & " is not in the collection of Pipes.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oRegistrationActivity As MUSTER.Info.RegistrationActivityInfo)
            Try
                colRegistrationActivitys.Remove(oRegistrationActivity)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("RegistrationActivity " & oRegistrationActivity.RegActionIndex & " is not in the collection of RegistrationActivitys.")
        End Sub
        Public Sub Flush()
            Dim xRegistrationActivityInfo As MUSTER.Info.RegistrationActivityInfo
            For Each xRegistrationActivityInfo In colRegistrationActivitys.Values
                If xRegistrationActivityInfo.IsDirty Then
                    oRegistrationActivityInfo = xRegistrationActivityInfo
                    '********************************************************
                    '
                    ' Note that if there are contained objects and the respective
                    '  contained object collections are dirty, then the contained
                    '  collections MUST BE FLUSHED before this object can be
                    '  saved.  Otherwise, there is a risk that an attempt will
                    '  be made to insert a new object to the repository without
                    '  corresponding contained information being present which 
                    '  may, in turn, cause a foreign key violation!
                    '
                    '********************************************************
                    Me.Save()
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
            Dim strArr() As String = colRegistrationActivitys.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.RegActionIndex.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colRegistrationActivitys.Item(nArr.GetValue(colIndex + direction)).RegActionIndex.ToString
            Else
                Return colRegistrationActivitys.Item(nArr.GetValue(colIndex)).RegActionIndex.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oRegistrationActivityInfo = New MUSTER.Info.RegistrationActivityInfo
        End Sub
        Public Sub Reset()
            oRegistrationActivityInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oRegistrationActivityInfoLocal As New MUSTER.Info.PipeInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("RegistrationActivity ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oRegistrationActivityInfoLocal In colRegistrationActivitys.Values
                    dr = tbEntityTable.NewRow()
                    dr("Pipe ID") = oRegistrationActivityInfoLocal.ID
                    dr("Deleted") = oRegistrationActivityInfoLocal.Deleted
                    dr("Created By") = oRegistrationActivityInfoLocal.CreatedBy
                    dr("Date Created") = oRegistrationActivityInfoLocal.CreatedOn
                    dr("Last Edited By") = oRegistrationActivityInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oRegistrationActivityInfoLocal.ModifiedOn
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
        Private Sub RegistrationActivityInfoChanged(ByVal bolValue As Boolean) Handles oRegistrationActivityInfo.RegistrationInfoChanged
            RaiseEvent RegistrationActivityChanged(bolValue)
        End Sub
        Private Sub RegistrationActivityColChanged(ByVal bolValue As Boolean) Handles colRegistrationActivitys.RegistrationColChanged
            RaiseEvent ColChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
