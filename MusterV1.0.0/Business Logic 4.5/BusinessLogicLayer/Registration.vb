'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Registration
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
' NOTE: This file to be used as Registration to build other objects.
'       Replace keyword "Registration" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pRegistration
#Region "Public Events"
        Public Event RegistrationErr(ByVal MsgStr As String)
        Public Event RegistrationChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)

        Public Event RegistrationActivityChanged(ByVal bolValue As Boolean)
        Public Event RegistrationActivityColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Public Owner Events"
        Public Event evtOwnerErr(ByVal MsgStr As String, ByVal strSrc As String)
        Public Event evtOwnerChanged(ByVal bolValue As Boolean, ByVal strSrc As String)
        Public Event evtOwnersChanged(ByVal bolValue As Boolean, ByVal strSrc As String)
        Public Event evtValidationErr(ByVal ID As Integer, ByVal MsgStr As String, ByVal strSrc As String)
        Public Event evtOwnerCommentsChanged(ByVal bolValue As Boolean)

        'facility
        Public Event evtFacilityChanged(ByVal bolValue As Boolean, ByVal strSrc As String)
        Public Event evtFacilitiesChanged(ByVal bolValue As Boolean, ByVal strSrc As String)
        Public Event evtFacilityCommentsChanged(ByVal bolValue As Boolean)
        'Added By Elango 
        Public Event evtOwnFacilityCAPStatusChanged(ByVal BolValue As Boolean, ByVal nFacId As Integer, ByVal strSrc As String)

        'Tank
        Public Event evtTankCommentsChanged(ByVal bolValue As Boolean)
        Public Event evtTankValidationErr(ByVal tnkID As Integer, ByVal strMessage As String, ByVal strSrc As String)

        'Pipe
        Public Event evtPipeCommentsChanged(ByVal bolValue As Boolean)

        'address
        Public Event evtAddressChanged(ByVal bolValue As Boolean, ByVal strSrc As String)
        Public Event evtAddressesChanged(ByVal bolValue As Boolean, ByVal strSrc As String)

        'persona
        Public Event evtPersonaChanged(ByVal bolValue As Boolean, ByVal strSrc As String)
        Public Event evtPersonasChanged(ByVal bolValue As Boolean, ByVal strSrc As String)

#End Region
#Region "Private Member Variables"
        Private WithEvents oRegistrationInfo As MUSTER.Info.RegistrationInfo
        Private WithEvents colRegistrations As MUSTER.Info.RegistrationsCollection
        Private WithEvents RegistrationActivity As MUSTER.BusinessLogic.pRegistrationActivity
        Private oRegistrationDB As New MUSTER.DataAccess.RegistrationDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
#End Region
#Region "Constructors"
        Public Sub New()
            oRegistrationInfo = New MUSTER.Info.RegistrationInfo
            colRegistrations = New MUSTER.Info.RegistrationsCollection
            RegistrationActivity = New MUSTER.BusinessLogic.pRegistrationActivity
        End Sub
        '********************************************************
        '
        ' Overloaded NEW which will populate with a single instance
        '   of the named Registration object.
        '
        '********************************************************
        Public Sub New(ByVal RegistrationID As Integer)
            oRegistrationInfo = New MUSTER.Info.RegistrationInfo
            colRegistrations = New MUSTER.Info.RegistrationsCollection
            RegistrationActivity = New MUSTER.BusinessLogic.pRegistrationActivity
            Me.Retrieve(RegistrationID)
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return oRegistrationInfo.ID
            End Get
            Set(ByVal Value As Integer)
                oRegistrationInfo.ID = Integer.Parse(Value)
            End Set
        End Property
        Public Property nIDProperty() As Integer
            Get
                Return nID
            End Get
            Set(ByVal Value As Integer)
                nid = Value
            End Set
        End Property
        Public Property OWNER_ID() As Int64
            Get
                Return oRegistrationInfo.OWNER_ID
            End Get
            Set(ByVal Value As Int64)
                oRegistrationInfo.OWNER_ID = Value
            End Set
        End Property

        Public Property DATE_STARTED() As DateTime
            Get
                Return oRegistrationInfo.DATE_STARTED
            End Get
            Set(ByVal Value As DateTime)
                oRegistrationInfo.DATE_STARTED = Value
            End Set
        End Property

        Public Property DATE_COMPLETED() As DateTime
            Get
                Return oRegistrationInfo.DATE_COMPLETED
            End Get
            Set(ByVal Value As DateTime)
                oRegistrationInfo.DATE_COMPLETED = Value
            End Set
        End Property
        Public Property COMPLETED() As Boolean
            Get
                Return oRegistrationInfo.COMPLETED
            End Get

            Set(ByVal value As Boolean)
                oRegistrationInfo.COMPLETED = value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oRegistrationInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oRegistrationInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oRegistrationInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oRegistrationInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xRegistrationinfo As MUSTER.Info.RegistrationInfo
                For Each xRegistrationinfo In colRegistrations.Values
                    If xRegistrationinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oRegistrationInfo.IsDirty = Value
            End Set
        End Property

        Public Property Activity() As MUSTER.BusinessLogic.pRegistrationActivity
            Get
                Return RegistrationActivity
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pRegistrationActivity)
                Me.RegistrationActivity = Value
            End Set
        End Property
#End Region

#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As MUSTER.Info.RegistrationInfo
            Dim oRegistrationInfoLocal As MUSTER.Info.RegistrationInfo
            Dim bolDataAged As Boolean = False
            Try
                '----- look in the collections for the ID value -----------------------
                For Each oRegistrationInfoLocal In colRegistrations.Values
                    If oRegistrationInfoLocal.ID = ID Then
                        If oRegistrationInfoLocal.IsAgedData = True And oRegistrationInfoLocal.IsDirty = False Then
                            bolDataAged = True
                            Exit For
                        Else
                            oRegistrationInfo = oRegistrationInfoLocal
                            RegistrationActivity.GetAllForRegistration(oRegistrationInfo.ID.ToString)
                            Return oRegistrationInfo
                        End If
                    End If
                Next
                '-----------------------------------------------------------------------

                If bolDataAged = True Then
                    colRegistrations.Remove(oRegistrationInfoLocal)
                End If
                '-----------------------------------------------------------------------

                oRegistrationInfo = oRegistrationDB.DBGetByID(ID)
                If oRegistrationInfo.ID = 0 Then
                    oRegistrationInfo.ID = nID
                    nID -= 1
                End If
                colRegistrations.Add(oRegistrationInfo)
                RegistrationActivity.GetAllForRegistration(oRegistrationInfo.ID.ToString)
                Return oRegistrationInfo
                '------------------------------------------------------------------------

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Obtains and returns an entity as called for by OwnerID
        Public Function RetrieveByOwnerID(ByVal nOwnerID As Integer) As MUSTER.Info.RegistrationInfo
            Try
                oRegistrationInfo = oRegistrationDB.DBGetByOwnerID(nOwnerID)
                Return oRegistrationInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal EntityID As String) As MUSTER.Info.RegistrationInfo
            Dim oRegistrationInfoLocal As MUSTER.Info.RegistrationInfo
            Try
                For Each oRegistrationInfoLocal In colRegistrations.Values
                    If oRegistrationInfoLocal.OWNER_ID = Integer.Parse(EntityID) _
                        And Not oRegistrationInfoLocal.COMPLETED Then

                        oRegistrationInfo = oRegistrationInfoLocal

                        RegistrationActivity.GetAllForRegistration(oRegistrationInfo.ID.ToString)
                        Return oRegistrationInfo
                    End If
                Next

                Return Nothing
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
                oldID = oRegistrationInfo.ID
                oRegistrationDB.Put(oRegistrationInfo)
                If oldID <= 0 Then
                    colRegistrations.ChangeKey(oldID, oRegistrationInfo.ID)
                End If

                RegistrationActivity.Flush()
                oRegistrationInfo.Archive()
                oRegistrationInfo.IsDirty = False
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
                    RaiseEvent RegistrationErr(errStr)
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
        Function GetAll() As MUSTER.Info.RegistrationsCollection
            Try
                colRegistrations.Clear()
                colRegistrations = oRegistrationDB.GetAllInfo
                Return colRegistrations
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oRegistrationInfo = oRegistrationDB.DBGetByID(ID)
                If oRegistrationInfo.ID = 0 Then
                    oRegistrationInfo.ID = nID
                    nID -= 1
                End If
                RegistrationActivity.GetAllForRegistration(oRegistrationInfo.ID.ToString)
                colRegistrations.Add(oRegistrationInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oRegistration As MUSTER.Info.RegistrationInfo)
            Try
                oRegistrationInfo = oRegistration
                If oRegistrationInfo.ID = 0 Then
                    oRegistrationInfo.ID = nID
                    nID -= 1
                End If
                RegistrationActivity.GetAllForRegistration(oRegistrationInfo.ID.ToString)
                colRegistrations.Add(oRegistrationInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oRegistrationInfoLocal As MUSTER.Info.RegistrationInfo

            Try
                For Each oRegistrationInfoLocal In colRegistrations.Values
                    If oRegistrationInfoLocal.ID = ID Then
                        colRegistrations.Remove(oRegistrationInfoLocal)
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
        Public Sub Remove(ByVal oRegistration As MUSTER.Info.RegistrationInfo)
            Try
                colRegistrations.Remove(oRegistration)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Registration " & oRegistration.ID & " is not in the collection of Registrations.")
        End Sub
        Public Sub Flush()
            Dim xRegistrationInfo As MUSTER.Info.RegistrationInfo
            For Each xRegistrationInfo In colRegistrations.Values
                If xRegistrationInfo.IsDirty Then
                    oRegistrationInfo = xRegistrationInfo
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
            Dim strArr() As String = colRegistrations.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colRegistrations.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colRegistrations.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oRegistrationInfo = New MUSTER.Info.RegistrationInfo
        End Sub
        Public Sub Reset()
            oRegistrationInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oRegistrationInfoLocal As New MUSTER.Info.PipeInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("Registration ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oRegistrationInfoLocal In colRegistrations.Values
                    dr = tbEntityTable.NewRow()
                    dr("Pipe ID") = oRegistrationInfoLocal.ID
                    dr("Deleted") = oRegistrationInfoLocal.Deleted
                    dr("Created By") = oRegistrationInfoLocal.CreatedBy
                    dr("Date Created") = oRegistrationInfoLocal.CreatedOn
                    dr("Last Edited By") = oRegistrationInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oRegistrationInfoLocal.ModifiedOn
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
        Private Sub RegistrationInfoChanged(ByVal bolValue As Boolean) Handles oRegistrationInfo.RegistrationInfoChanged
            RaiseEvent RegistrationChanged(bolValue)
        End Sub
        Private Sub RegistrationColChanged(ByVal bolValue As Boolean) Handles colRegistrations.RegistrationColChanged
            RaiseEvent ColChanged(bolValue)
        End Sub

        Private Sub RegistrationActivityInfoChanged(ByVal bolValue As Boolean) Handles oRegistrationInfo.RegistrationInfoChanged
            RaiseEvent RegistrationActivityChanged(bolValue)
        End Sub
        Private Sub RegistrationActivityColChangedsub(ByVal bolValue As Boolean) Handles colRegistrations.RegistrationColChanged
            RaiseEvent RegistrationActivityColChanged(bolValue)
        End Sub
#End Region

    End Class
End Namespace
