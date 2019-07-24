'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Calendar
'   Provides the info and collection objects to the client for manipulating
'   an CalendarInfo object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         KJ      12/07/04    Original class definition.
'   1.1         KJ      12/23/04    Made changes to header. 
'   1.2         AN      01/03/05    Added Try catch and Exception Handling/Logging
'   1.3         KJ      01/03/05    Changed the exposed attributes properties to have collection.
'   1.4         KJ      01/06/05    Added Events. Converted Function ColIsDirty to a Property.
'                                   Changed Retrieve and Add Function. Removed Add(ID) and GetCalendar(ID).
'   1.5         JVC2    01/17/05    Added GetLoadedCalendar() which loads the calendar entries
'                                   for the specified entity.
'   1.6         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'
'   1.7         MR      01/27/05    Added CreateTable() common function and Modified Save,Remove Functions.
'   1.8         JVC2    02/02/05    Added EntityTypeID to private members and initialize to "Calendar" type.
'                                       Also added EntityType attribute to expose the typeID.
'   1.9         AB      02/17/05    Added DataAge check to the Retrieve function
'   2.0         JVC2    02/21/05    Added GetByComment Function.  
'   2.1         JVC2    03/25/05    Added Owning_Entity_Type and Owning_Entity_ID
'   2.2         MNR     04/07/05    Added Retrive(ByVal nEntityType As Int16, ByVal nEntityID As Int64, Optional ByVal strVal As String = "", Optional ByVal strType As String = "USER")
'   2.3         JVC2    04/12/05    Added MarkToDoDeleted and MarkDueToMeDeleted operations
'
' Function                      Description
' New()                 Initializes the CalendarCollection and CalendarInfo objects.
' Retrieve(ID)          Sets the internal CalendarInfo to the CaledanrInfo matching the supplied key.
' GetCalendarAll()      Returns an CalendarCollection with all CalendarInfo objects
' Add(CalendarInfo)     Adds the CalendarInfo passed as the argument to the internal 
'                           CalendarCollection
' Remove(ID)            Removes the CalendarInfo identified by arg ID from the internal 
'                           CalendarCollection
' colIsDirty()          Returns a boolean indicating whether any of the CalendarInfo
'                           objects in the CalendarCollection has been modified since the
'                           last time it was retrieved from/saved to the repository.
' Flush()               Marshalls all modified/added CalendarInfo objects in the 
'                           CalendarCollection to the repository.
' Save()                Marshalls the internal CalendarInfo object to the repository.
' Save(CalendarInfo)    Marshalls the CalendarInfo object to the repository as supplied.
' CalendarTable()       Returns a datatable containing all columns for the CalendarInfo
'                           objects in the internal CalendarCollection.
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pCalendar
#Region "Private Member Variables"
        Private WithEvents colCalendar As Muster.Info.CalendarCollection
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private WithEvents oCalendarInfo As Muster.Info.CalendarInfo
        Private oCalendarDB As New Muster.DataAccess.CalendarDB
        Private nNewIndex As Integer = 0
        Private blnShowDeleted As Boolean = False
        Private nID As Int64 = -1
        Private MusterException As New MUSTER.Exceptions.MusterExceptions

        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Calendar").ID
#End Region
#Region "Public Event Handlers"
        Public Event CalendarChanged(ByVal IsDirtyState As Boolean)
        Public Event evtCalErr(ByVal MsgStr As String)
#End Region
#Region "Constructors"
        Public Sub New()
            oCalendarInfo = New Muster.Info.CalendarInfo
            colCalendar = New Muster.Info.CalendarCollection
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property CalendarId() As Integer
            Get
                Return oCalendarInfo.CalendarInfoId
            End Get

            Set(ByVal value As Integer)
                oCalendarInfo.CalendarInfoId = Integer.Parse(value)
                colCalendar(oCalendarInfo.CalendarInfoId) = oCalendarInfo
            End Set
        End Property

        Public Property Owning_Entity_Type() As Integer
            Get
                Return oCalendarInfo.OwningEntityType
            End Get
            Set(ByVal Value As Integer)
                oCalendarInfo.OwningEntityType = Value
            End Set
        End Property

        Public Property Owning_Entity_ID() As Int64
            Get
                Return oCalendarInfo.OwningEntityID
            End Get
            Set(ByVal Value As Int64)
                oCalendarInfo.OwningEntityID = Value
            End Set
        End Property

        Public Property NotificationDate() As DateTime
            Get
                Return oCalendarInfo.NotificationDate
            End Get

            Set(ByVal value As DateTime)
                oCalendarInfo.NotificationDate = value
                colCalendar(oCalendarInfo.CalendarInfoId) = oCalendarInfo
            End Set
        End Property

        Public Property DateDue() As DateTime
            Get
                Return oCalendarInfo.DateDue
            End Get

            Set(ByVal value As DateTime)
                oCalendarInfo.DateDue = value
                colCalendar(oCalendarInfo.CalendarInfoId) = oCalendarInfo
            End Set
        End Property

        Public Property CurrentColorCode() As Integer
            Get
                Return oCalendarInfo.CurrentColorCode
            End Get

            Set(ByVal value As Integer)
                oCalendarInfo.CurrentColorCode = Integer.Parse(value)
                colCalendar(oCalendarInfo.CalendarInfoId) = oCalendarInfo
            End Set
        End Property

        Public Property TaskDescription() As String
            Get
                Return oCalendarInfo.TaskDescription
            End Get

            Set(ByVal value As String)
                oCalendarInfo.TaskDescription = value
                colCalendar(oCalendarInfo.CalendarInfoId) = oCalendarInfo
            End Set
        End Property

        Public Property UserId() As String
            Get
                Return oCalendarInfo.UserId
            End Get

            Set(ByVal value As String)
                oCalendarInfo.UserId = value
                colCalendar(oCalendarInfo.CalendarInfoId) = oCalendarInfo
            End Set
        End Property

        Public Property SourceUserId() As String
            Get
                Return oCalendarInfo.SourceUserId
            End Get

            Set(ByVal value As String)
                oCalendarInfo.SourceUserId = value
                colCalendar(oCalendarInfo.CalendarInfoId) = oCalendarInfo
            End Set
        End Property

        Public Property GroupId() As String
            Get
                Return oCalendarInfo.GroupId
            End Get

            Set(ByVal value As String)
                oCalendarInfo.GroupId = value
                colCalendar(oCalendarInfo.CalendarInfoId) = oCalendarInfo
            End Set
        End Property

        Public Property DueToMe() As Boolean
            Get
                Return oCalendarInfo.DueToMe
            End Get

            Set(ByVal value As Boolean)
                oCalendarInfo.DueToMe = value
                colCalendar(oCalendarInfo.CalendarInfoId) = oCalendarInfo
            End Set
        End Property

        Public Property ToDo() As Boolean
            Get
                Return oCalendarInfo.ToDo
            End Get

            Set(ByVal value As Boolean)
                oCalendarInfo.ToDo = value
                colCalendar(oCalendarInfo.CalendarInfoId) = oCalendarInfo
            End Set
        End Property

        Public Property Completed() As Boolean
            Get
                Return oCalendarInfo.Completed
            End Get

            Set(ByVal value As Boolean)
                oCalendarInfo.Completed = value
                colCalendar(oCalendarInfo.CalendarInfoId) = oCalendarInfo
            End Set
        End Property

        'Public ReadOnly Property EntityType() As Integer
        '    Get
        '        Return nEntityTypeID
        '    End Get
        'End Property

        Public Property Deleted() As Boolean
            Get
                Return oCalendarInfo.Deleted
            End Get

            Set(ByVal value As Boolean)
                oCalendarInfo.Deleted = value
                colCalendar(oCalendarInfo.CalendarInfoId) = oCalendarInfo
            End Set
        End Property

        Public Property IsDirty() As Boolean
            Get
                Return oCalendarInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oCalendarInfo.IsDirty = value
                colCalendar(oCalendarInfo.CalendarInfoId) = oCalendarInfo
            End Set
        End Property

        Public Property colIsDirty() As Boolean
            Get
                Dim xCalendarInfo As MUSTER.Info.CalendarInfo
                For Each xCalendarInfo In colCalendar.Values
                    If xCalendarInfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)

            End Set
        End Property

        Public Property ShowDeleted() As Boolean
            Get
                Return blnShowDeleted
            End Get

            Set(ByVal value As Boolean)
                blnShowDeleted = value
            End Set
        End Property

        Public Property CreatedBy() As String
            Get
                Return oCalendarInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oCalendarInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oCalendarInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oCalendarInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oCalendarInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oCalendarInfo.ModifiedOn
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Function Retrieve(ByVal ID As Int64) As MUSTER.Info.CalendarInfo
            Dim oCalendarInfoLocal As MUSTER.Info.CalendarInfo
            Dim bolDataAged As Boolean = False

            Try
                For Each oCalendarInfoLocal In colCalendar.Values
                    If oCalendarInfoLocal.CalendarInfoId = ID Then
                        If oCalendarInfoLocal.IsAgedData = True And oCalendarInfoLocal.IsDirty = False Then
                            bolDataAged = True
                            Exit For
                        Else
                            oCalendarInfo = oCalendarInfoLocal
                            Return oCalendarInfo
                        End If
                    End If
                Next
                If bolDataAged Then
                    colCalendar.Remove(oCalendarInfoLocal)
                End If

                oCalendarInfo = oCalendarDB.DBGetByID(ID)
                If oCalendarInfo.CalendarInfoId = 0 Then
                    oCalendarInfo.CalendarInfoId = nID
                    nID -= 1
                End If
                colCalendar.Add(oCalendarInfo)
                Return oCalendarInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetByComment(ByVal OtherID As String) As MUSTER.Info.CalendarCollection
            Dim oCalendarInfoLocal As MUSTER.Info.CalendarInfo
            Dim oCalendarColLocal As MUSTER.Info.CalendarCollection
            Dim CalInfToAdd As MUSTER.Info.CalendarInfo
            oCalendarColLocal = New MUSTER.Info.CalendarCollection
            Try
                Return oCalendarDB.DBGetByOtherID(OtherID, "DESCRIPTION")
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function Retrieve(ByVal OtherID As String, Optional ByVal strType As String = "USER")
            Dim oCalendarInfoLocal As MUSTER.Info.CalendarInfo
            Dim oCalendarColLocal As MUSTER.Info.CalendarCollection
            oCalendarColLocal = New MUSTER.Info.CalendarCollection
            Try
                oCalendarColLocal = oCalendarDB.DBGetByOtherID(OtherID, strType)
                For Each oCalendarInfoLocal In oCalendarColLocal.Values
                    colCalendar.Add(oCalendarInfoLocal)
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function Retrieve(ByVal nEntityType As Int16, ByVal nEntityID As Int64, Optional ByVal strVal As String = "", Optional ByVal strType As String = "USER")
            Dim oCalendarInfoLocal As MUSTER.Info.CalendarInfo
            Dim oCalendarColLocal As MUSTER.Info.CalendarCollection
            oCalendarColLocal = New MUSTER.Info.CalendarCollection
            Try
                oCalendarColLocal = oCalendarDB.DBGetByOtherID(strVal, strType, nEntityType, nEntityID)
                For Each oCalendarInfoLocal In oCalendarColLocal.Values
                    colCalendar.Add(oCalendarInfoLocal)
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function RetrieveByOtherID(ByVal nEntityType As Int16, ByVal nEntityID As Int64, Optional ByVal strVal As String = "", Optional ByVal strType As String = "USER") As MUSTER.Info.CalendarCollection
            Try
                Return oCalendarDB.DBGetByOtherID(strVal, strType, nEntityType, nEntityID)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Function GetDataSet(ByVal strSQL As String) As DataSet
            Try
                Dim ds As DataSet
                ds = oCalendarDB.DBGetDS(strSQL)
                Return ds
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetParentEntityID(ByVal fromEntityID As Int64, ByVal fromEntityType As Int64, ByVal toEntityType As Int64) As Int64
            Try
                Return oCalendarDB.DBGetParentEntityID(fromEntityID, fromEntityType, toEntityType)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        Function GetLoadedCalendar(ByVal strID As String, Optional ByVal strType As String = "USER") As DataTable
            Dim oCalInfo As MUSTER.Info.CalendarInfo
            Dim oCalCol As New MUSTER.Info.CalendarCollection
            Dim bolAddRow As Boolean = False

            Dim dtCal As New DataTable
            dtCal = CreateTable()

            Dim drRow As DataRow

            '
            'First, force a load of the Calendar entries for the requested entity
            '
            Me.Retrieve(strID, strType)
            '
            'Now, build the table with the mathcing entries from the collection
            '
            For Each oCalInfo In colCalendar.Values
                bolAddRow = False
                Select Case strType
                    Case "USER"
                        If oCalInfo.UserId = strID Then
                            bolAddRow = True
                        End If
                    Case "GROUP"
                        If oCalInfo.GroupId = strID Then
                            bolAddRow = True
                        End If
                End Select
                If bolAddRow Then
                    drRow = dtCal.NewRow
                    With oCalInfo
                        drRow.Item("CALENDAR_INFO_ID") = .CalendarInfoId
                        If .UserId <> "" Then
                            drRow.Item("TARGET") = .UserId
                        Else
                            drRow.Item("TARGET") = .GroupId
                        End If
                        drRow.Item("USER_ID") = .UserId
                        drRow.Item("NOTIFICATION_DATE") = .NotificationDate
                        drRow.Item("DATE_DUE") = .DateDue
                        drRow.Item("TASK_DESCRIPTION") = .TaskDescription
                        drRow.Item("SOURCE_USER_ID") = .SourceUserId
                        drRow.Item("GROUP_ID") = .GroupId
                        drRow.Item("DUE_TO_ME") = .DueToMe
                        drRow.Item("TO_DO") = .ToDo
                        drRow.Item("COMPLETED") = .Completed
                        drRow.Item("DELETED") = .Deleted
                        drRow.Item("CREATED_BY") = .CreatedBy
                        drRow.Item("DATE_CREATED") = .CreatedOn
                        drRow.Item("LAST_EDITED_BY") = .ModifiedBy
                        drRow.Item("DATE_LAST_EDITED") = .ModifiedOn
                        drRow.Item("Owning_Entity_Type") = .OwningEntityType
                        drRow.Item("Owning_Entity_ID") = .OwningEntityID
                    End With
                    dtCal.Rows.Add(drRow)
                End If
            Next

            Return dtCal

        End Function
        Function GetCalendarAll(ByVal str As String, ByVal blnShowDeleted As Boolean) As MUSTER.Info.CalendarCollection
            Try
                colCalendar.Clear()
                colCalendar = oCalendarDB.GetAllInfo(str, blnShowDeleted)
                Return colCalendar
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub Add(ByRef oCalendar As MUSTER.Info.CalendarInfo)
            Try
                oCalendarInfo = oCalendar

                If oCalendarInfo.CalendarInfoId = 0 Then
                    oCalendarInfo.CalendarInfoId = nID
                    nID -= 1
                End If
                If Not colCalendar.Contains(oCalendarInfo.CalendarInfoId) Then
                    colCalendar.Add(oCalendarInfo)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Marks Calendar Entries for the indicated Entity as deleted
        Public Sub MarkToDoDeleted(ByVal nOwningEntityType As Int64, ByVal nOwningEntityID As Int64, Optional ByVal moduleID As Integer = 0, Optional ByVal staffID As Integer = 0, Optional ByRef returnVal As String = "", Optional ByVal UserID As String = "")
            Try
                Dim oCalendarInfoLocal As MUSTER.Info.CalendarInfo
                Dim myIndex As Int16 = 1
                For Each oCalendarInfoLocal In colCalendar.Values
                    If oCalendarInfoLocal.OwningEntityID = nOwningEntityID And _
                        oCalendarInfoLocal.OwningEntityType = nOwningEntityType And _
                        oCalendarInfoLocal.ToDo Then
                        oCalendarInfoLocal.Deleted = True
                    End If
                Next
                If Me.colIsDirty Then
                    Me.Flush(moduleID, staffID, returnVal, UserID)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Marks Calendar Entries for the indicated Entity as deleted
        Public Sub MarkDueToMeDeleted(ByVal nOwningEntityType As Int64, ByVal nOwningEntityID As Int64, Optional ByVal moduleID As Integer = 0, Optional ByVal staffID As Integer = 0, Optional ByRef returnVal As String = "", Optional ByVal UserID As String = "")
            Try
                Dim oCalendarInfoLocal As MUSTER.Info.CalendarInfo
                Dim myIndex As Int16 = 1
                For Each oCalendarInfoLocal In colCalendar.Values
                    If oCalendarInfoLocal.OwningEntityID = nOwningEntityID And _
                        oCalendarInfoLocal.OwningEntityType = nOwningEntityType And _
                        oCalendarInfoLocal.DueToMe Then
                        oCalendarInfoLocal.Deleted = True
                    End If
                Next
                If Me.colIsDirty Then
                    Me.Flush(moduleID, staffID, returnVal, UserID)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Marks Calendar Entries for the indicated Entity as deleted
        Public Sub MarkToDoCompleted(ByVal nOwningEntityType As Int64, ByVal nOwningEntityID As Int64, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Try
                Dim oCalendarInfoLocal As MUSTER.Info.CalendarInfo
                Dim myIndex As Int16 = 1
                For Each oCalendarInfoLocal In colCalendar.Values
                    If oCalendarInfoLocal.OwningEntityID = nOwningEntityID And _
                        oCalendarInfoLocal.OwningEntityType = nOwningEntityType And _
                        oCalendarInfoLocal.ToDo Then
                        oCalendarInfoLocal.Completed = True
                    End If
                Next
                If Me.colIsDirty Then
                    Me.Flush(moduleID, staffID, returnVal, UserID)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Marks Calendar Entries for the indicated Entity as deleted
        Public Sub MarkDueToMeCompleted(ByVal nOwningEntityType As Int64, ByVal nOwningEntityID As Int64, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Try
                Dim oCalendarInfoLocal As MUSTER.Info.CalendarInfo
                Dim myIndex As Int16 = 1
                For Each oCalendarInfoLocal In colCalendar.Values
                    If oCalendarInfoLocal.OwningEntityID = nOwningEntityID And _
                        oCalendarInfoLocal.OwningEntityType = nOwningEntityType And _
                        oCalendarInfoLocal.DueToMe Then
                        oCalendarInfoLocal.Completed = True
                    End If
                Next
                If Me.colIsDirty Then
                    Me.Flush(moduleID, staffID, returnVal, UserID)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Removes the Calendar called for by ID from the collection
        Public Sub Remove(ByVal ID As Int64)
            Try
                If colCalendar.Contains(ID) Then
                    colCalendar.Remove(ID)
                End If
                'Dim oCalendarInfoLocal As MUSTER.Info.CalendarInfo
                'Dim myIndex As Int16 = 1
                'For Each oCalendarInfoLocal In colCalendar.Values
                '    If oCalendarInfoLocal.CalendarInfoId = ID Then
                '        colCalendar.Remove(oCalendarInfoLocal)
                '        oCalendarInfo = New MUSTER.Info.CalendarInfo
                '        Exit Sub
                '    End If
                '    myIndex += 1
                'Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'Throw New Exception("Calendar " & ID.ToString & " is not in the collection of Calendares.")
        End Sub
        Public Sub Flush(Optional ByVal moduleID As Integer = 0, Optional ByVal staffID As Integer = 0, Optional ByRef returnVal As String = "", Optional ByVal UserID As String = "")
            Dim oTempInfo As MUSTER.Info.CalendarInfo
            Dim IDs As New Collection
            Dim index As Integer = 0
            For Each oTempInfo In colCalendar.Values
                If oTempInfo.IsDirty Then
                    'oCalendarInfo = oTempInfo
                    IDs.Add(oTempInfo.CalendarInfoId)
                    'Me.Save()
                    If oTempInfo.CalendarInfoId <= 0 Then
                        oTempInfo.CreatedBy = UserID
                    Else
                        oTempInfo.ModifiedBy = UserID
                    End If
                End If
            Next

            If Not (IDs Is Nothing) Then
                For index = 1 To IDs.Count
                    Dim colKey As String = CType(IDs.Item(index), String)
                    oTempInfo = Me.Retrieve(colKey)
                    Me.Save(moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If

                Next
            End If

        End Sub
        Public Sub Save(Optional ByVal moduleID As Integer = 0, Optional ByVal staffID As Integer = 0, Optional ByRef returnVal As String = "")
            Dim oldID As Integer
            Try
                If Not Me.ValidateData() Then
                    Exit Sub
                End If
                oldID = oCalendarInfo.CalendarInfoId
                oCalendarDB.Put(oCalendarInfo, moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If

                If oldID <= 0 Then
                    colCalendar.ChangeKey(oldID, oCalendarInfo.CalendarInfoId)
                End If
                oCalendarInfo.Archive()
                oCalendarInfo.IsDirty = False

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Sub Save(ByRef oCalendar As MUSTER.Info.CalendarInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                If Not Me.ValidateData() Then
                    Exit Sub
                End If

                oCalendarDB.Put(oCalendar, moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If

                oCalendar.Archive()
                oCalendar.IsDirty = False

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colCalendar.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.CalendarId.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colCalendar.Item(nArr.GetValue(colIndex + direction)).CalendarInfoId.ToString
            Else
                Return colCalendar.Item(nArr.GetValue(colIndex)).CalendarInfoId.ToString
            End If
        End Function
        ' ValidateData before Save
        Public Function ValidateData() As Boolean

            Try
                Dim strCalErrMsg As String = ""
                Dim validateSuccess As Boolean = True

                Dim dt As Date
                If Date.Compare(dt, oCalendarInfo.DateDue) = 0 Then
                    strCalErrMsg += vbTab + "Due Date is required" + vbCrLf
                    validateSuccess = False
                End If

                If oCalendarInfo.TaskDescription = String.Empty Then
                    strCalErrMsg += vbTab + "Task Description is required" + vbCrLf
                    validateSuccess = False
                End If

                If oCalendarInfo.ToDo = False And oCalendarInfo.DueToMe = False Then
                    strCalErrMsg += vbTab + "Either To Do or Due To Me must be selected" + vbCrLf
                    validateSuccess = False
                End If

                If oCalendarInfo.GroupId = String.Empty And oCalendarInfo.UserId = String.Empty Then
                    strCalErrMsg += vbTab + "Either UserId or Group must be selected" + vbCrLf
                    validateSuccess = False
                End If

                If validateSuccess Then
                    Return validateSuccess
                Else
                    RaiseEvent evtCalErr(strCalErrMsg)
                    Return validateSuccess
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oCalendarInfo = New MUSTER.Info.CalendarInfo
        End Sub
        Public Sub Reset()
            oCalendarInfo.Reset()
        End Sub
        Public Function DBGetByID(ByVal ID As Integer)
            Try
                oCalendarInfo = oCalendarDB.DBGetByID(ID)
                If Not oCalendarInfo Is Nothing Then
                    Return oCalendarInfo
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#Region "Miscellaneous Operations"
        'Returns a datatable of the Calendares in the collection
        Public Function CalendarTableToDo() As DataTable

            Dim oCalendarInfoLocal As MUSTER.Info.CalendarInfo
            Dim dr As DataRow
            Dim tbCalendarTable As DataTable

            Try

                tbCalendarTable = CreateTable()

                For Each oCalendarInfoLocal In colCalendar.Values
                    If oCalendarInfoLocal.ToDo = True Then
                        dr = tbCalendarTable.NewRow()
                        dr("CALENDAR_INFO_ID") = oCalendarInfoLocal.CalendarInfoId
                        If oCalendarInfoLocal.GroupId = String.Empty Then
                            dr("TARGET") = oCalendarInfoLocal.UserId
                        Else
                            dr("TARGET") = oCalendarInfoLocal.GroupId
                        End If
                        dr("NOTIFICATION_DATE") = oCalendarInfoLocal.NotificationDate.ToShortDateString
                        dr("DATE_DUE") = oCalendarInfoLocal.DateDue.ToShortDateString
                        dr("CURRENT_COLOR_CODE") = oCalendarInfoLocal.CurrentColorCode
                        dr("TASK_DESCRIPTION") = oCalendarInfoLocal.TaskDescription
                        dr("USER_ID") = oCalendarInfoLocal.UserId
                        dr("SOURCE_USER_ID") = oCalendarInfoLocal.SourceUserId
                        dr("GROUP_ID") = oCalendarInfoLocal.GroupId
                        dr("DUE_TO_ME") = oCalendarInfoLocal.DueToMe
                        dr("TO_DO") = oCalendarInfoLocal.ToDo
                        dr("COMPLETED") = oCalendarInfoLocal.Completed
                        dr("DELETED") = oCalendarInfoLocal.Deleted
                        dr("CREATED_BY") = oCalendarInfoLocal.CreatedBy
                        dr("DATE_CREATED") = oCalendarInfoLocal.CreatedOn
                        dr("LAST_EDITED_BY") = oCalendarInfoLocal.ModifiedBy
                        dr("DATE_LAST_EDITED") = oCalendarInfoLocal.ModifiedOn
                        dr("Owning_Entity_Type") = oCalendarInfoLocal.OwningEntityType
                        dr("Owning_Entity_ID") = oCalendarInfoLocal.OwningEntityID
                        tbCalendarTable.Rows.Add(dr)
                    End If
                Next
                Return tbCalendarTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function CalendarTableDueToMe() As DataTable

            Dim oCalendarInfoLocal As MUSTER.Info.CalendarInfo
            Dim dr As DataRow
            Dim tbCalendarTable As DataTable

            Try
                tbCalendarTable = CreateTable()

                For Each oCalendarInfoLocal In colCalendar.Values
                    If oCalendarInfoLocal.DueToMe = True Then
                        dr = tbCalendarTable.NewRow()
                        dr("CALENDAR_INFO_ID") = oCalendarInfoLocal.CalendarInfoId
                        If oCalendarInfoLocal.GroupId = String.Empty Then
                            dr("TARGET") = oCalendarInfoLocal.UserId
                        Else
                            dr("TARGET") = oCalendarInfoLocal.GroupId
                        End If
                        dr("NOTIFICATION_DATE") = oCalendarInfoLocal.NotificationDate.ToShortDateString
                        dr("DATE_DUE") = oCalendarInfoLocal.DateDue.ToShortDateString
                        dr("CURRENT_COLOR_CODE") = oCalendarInfoLocal.CurrentColorCode
                        dr("TASK_DESCRIPTION") = oCalendarInfoLocal.TaskDescription
                        dr("USER_ID") = oCalendarInfoLocal.UserId
                        dr("SOURCE_USER_ID") = oCalendarInfoLocal.SourceUserId
                        dr("GROUP_ID") = oCalendarInfoLocal.GroupId
                        dr("DUE_TO_ME") = oCalendarInfoLocal.DueToMe
                        dr("TO_DO") = oCalendarInfoLocal.ToDo
                        dr("COMPLETED") = oCalendarInfoLocal.Completed
                        dr("DELETED") = oCalendarInfoLocal.Deleted
                        dr("CREATED_BY") = oCalendarInfoLocal.CreatedBy
                        dr("DATE_CREATED") = oCalendarInfoLocal.CreatedOn
                        dr("LAST_EDITED_BY") = oCalendarInfoLocal.ModifiedBy
                        dr("DATE_LAST_EDITED") = oCalendarInfoLocal.ModifiedOn
                        dr("Owning_Entity_Type") = oCalendarInfoLocal.OwningEntityType
                        dr("Owning_Entity_ID") = oCalendarInfoLocal.OwningEntityID
                        tbCalendarTable.Rows.Add(dr)
                    End If

                Next
                Return tbCalendarTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Private Function CreateTable() As DataTable

            Dim dtCal As New DataTable
            dtCal.Columns.Add("CALENDAR_INFO_ID", Type.GetType("System.Int64"))
            dtCal.Columns.Add("DATE_DUE", Type.GetType("System.DateTime"))
            dtCal.Columns.Add("TASK_DESCRIPTION", Type.GetType("System.String"))
            dtCal.Columns.Add("TARGET", Type.GetType("System.String"))
            dtCal.Columns.Add("USER_ID", Type.GetType("System.String"))
            dtCal.Columns.Add("NOTIFICATION_DATE", Type.GetType("System.DateTime"))
            dtCal.Columns.Add("CURRENT_COLOR_CODE", Type.GetType("System.Int64"))
            'dtCal.Columns.Add("TASK_DESCRIPTION", Type.GetType("System.String"))
            dtCal.Columns.Add("SOURCE_USER_ID", Type.GetType("System.String"))
            dtCal.Columns.Add("GROUP_ID", Type.GetType("System.String"))
            dtCal.Columns.Add("DUE_TO_ME", Type.GetType("System.Boolean"))
            dtCal.Columns.Add("TO_DO", Type.GetType("System.Boolean"))
            dtCal.Columns.Add("COMPLETED", Type.GetType("System.Boolean"))
            dtCal.Columns.Add("DELETED", Type.GetType("System.Boolean"))
            dtCal.Columns.Add("CREATED_BY", Type.GetType("System.String"))
            dtCal.Columns.Add("DATE_CREATED", Type.GetType("System.DateTime"))
            dtCal.Columns.Add("LAST_EDITED_BY", Type.GetType("System.String"))
            dtCal.Columns.Add("DATE_LAST_EDITED", Type.GetType("System.DateTime"))
            dtCal.Columns.Add("Owning_Entity_Type", Type.GetType("System.Int64"))
            dtCal.Columns.Add("Owning_Entity_ID", Type.GetType("System.Int64"))
            Return dtCal
        End Function
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub colCalendar_InfoChanged() Handles colCalendar.InfoChanged
            RaiseEvent CalendarChanged(Me.colIsDirty)
        End Sub
        Private Sub Calendar_Changed(ByVal bolState As Boolean) Handles oCalendarInfo.InfoBecameDirty
            RaiseEvent CalendarChanged(Me.colIsDirty)
        End Sub
#End Region

    End Class
End Namespace
