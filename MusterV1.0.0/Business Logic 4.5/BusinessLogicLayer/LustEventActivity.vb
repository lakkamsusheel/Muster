'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.LustEventActivity
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0         AN       3/8/2005    Original class definition
'
' Function          Description
' Retrieve(Name)    Returns the Lust Event Activity requested by the string arg NAME
' Retrieve(ID)      Returns the Lust Event Activity requested by the int arg ID
' GetAll()          Returns an LustEventActivitiyCollection with all Lust Event Activity objects
' Add(ID)           Adds the Lust Event Activity identified by arg ID to the 
'                           internal LustEventActivityCollection
' Add(LustEventInfo)Adds the Lust Event Activity passed as the argument to the internal 
'                           LustEventActivityCollection
' Remove(ID)        Removes the Lust Event Activity identified by arg ID from the internal 
'                           LustEventActivityCollection
' Remove(NAME)      Removes the Lust Event Activity identified by arg NAME from the 
'                           internal LustEventActivityCollection
' Flush()           Saves all objects in the collection
' Clear()           Clears the current object and all objects in the collection
' Reset()           Resets the current object to its original state
' EntityTable()     Returns a datatable containing all columns for the Lust Event 
'                           objects in the internal LustEventsCollection.
'
'
' Attribute          Description
'-------------------------------------------------------------------------------
' ActivityID         The system id for this Lust Event Activity object
' EventID            The id of the lust event that this Activity is associated to
' Closed             The date the item was closed
' Completed          The date the item was completed
' First_GWS_Below    Date
' Second_GWS_Below   Date
' Started            Date item was started
' Type               Activity type id for the item
'
' AgeThreshold       The maximum age the info object can attain before requiring a refresh
' CreatedBy          The ID of the user that created the row
' CreatedOn          The date on which the row was created
' Deleted            Indicates the deleted state of the row
' ModifiedBy         ID of the user that last made changes
' ModifiedOn         The date of the last changes made 
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pLustEventActivity
#Region "Public Events"
        Public Event LustEventErr(ByVal MsgStr As String)
        Public Event LustEventChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private WithEvents oLustEventActivityInfo As New MUSTER.Info.LustActivityInfo
        'Private WithEvents colLustEventActivities As MUSTER.Info.LustActivityCollection
        Private WithEvents oLustEventInfo As MUSTER.Info.LustEventInfo
        Private WithEvents oLustEventDocuments As MUSTER.BusinessLogic.pLustEventDocument
        Private oLustEventActivityDB As New MUSTER.DataAccess.LustEventActivityDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("LustActivity").ID
        Private onFacilityID As Integer = 0
        Private onUserID As Integer = 0
        Private oCalendar As New MUSTER.BusinessLogic.pCalendar
        Private oLetterGen As New MUSTER.BusinessLogic.pLetterGen

#End Region
#Region "Constructors"
        Public Sub New(Optional ByRef ActivityEvent As MUSTER.Info.LustEventInfo = Nothing)
            'oLustEventActivityInfo = New MUSTER.Info.LustActivityInfo
            'colLustEventActivities = New MUSTER.Info.LustActivityCollection            
            If ActivityEvent Is Nothing Then
                oLustEventDocuments = New MUSTER.BusinessLogic.pLustEventDocument(oLustEventActivityInfo)
                oLustEventInfo = New MUSTER.Info.LustEventInfo
            Else
                oLustEventDocuments = New MUSTER.BusinessLogic.pLustEventDocument(oLustEventActivityInfo)
                oLustEventInfo = ActivityEvent
                onFacilityID = oLustEventInfo.FacilityID
                onUserID = oLustEventInfo.UserID
            End If
         End Sub
        '********************************************************
        '
        ' Overloaded NEW which will populate with a single instance
        '   of the LustEvent object with the provided ID.
        '
        '********************************************************
        Public Sub New(ByVal LustEventID As Integer, Optional ByRef ActivityEvent As MUSTER.Info.LustEventInfo = Nothing)
            'oLustEventActivityInfo = New MUSTER.Info.LustActivityInfo
            'colLustEventActivities = New MUSTER.Info.LustActivityCollection            
            If ActivityEvent Is Nothing Then
                oLustEventDocuments = New MUSTER.BusinessLogic.pLustEventDocument(oLustEventActivityInfo)
                oLustEventInfo = New MUSTER.Info.LustEventInfo
            Else
                oLustEventDocuments = New MUSTER.BusinessLogic.pLustEventDocument(oLustEventActivityInfo)
                oLustEventInfo = ActivityEvent
                onFacilityID = oLustEventInfo.FacilityID
                onUserID = oLustEventInfo.UserID
            End If
            Me.Retrieve(LustEventID)
        End Sub
        '********************************************************
        '
        ' Overloaded NEW which will populate with a single instance
        '   of the named LustEvent object.
        '
        '********************************************************
        Public Sub New(ByVal LustEventName As String, Optional ByRef ActivityEvent As MUSTER.Info.LustEventInfo = Nothing)
            'oLustEventActivityInfo = New MUSTER.Info.LustActivityInfo
            'colLustEventActivities = New MUSTER.Info.LustActivityCollection
            If ActivityEvent Is Nothing Then
                oLustEventDocuments = New MUSTER.BusinessLogic.pLustEventDocument(oLustEventActivityInfo)
                oLustEventInfo = New MUSTER.Info.LustEventInfo
            Else
                oLustEventDocuments = New MUSTER.BusinessLogic.pLustEventDocument(oLustEventActivityInfo)
                oLustEventInfo = ActivityEvent
                onFacilityID = oLustEventInfo.FacilityID
                onUserID = oLustEventInfo.UserID
            End If
            Me.Retrieve(LustEventName)
        End Sub
#End Region
#Region "Exposed Attributes"
        ' The system generated ID for the LUST Activity (auto-increment in DB)
        Public Property ActivityID() As Integer
            Get
                Return oLustEventActivityInfo.ActivityID
            End Get
            Set(ByVal Value As Integer)
                oLustEventActivityInfo.ActivityID = Value
            End Set
        End Property
        'Public ReadOnly Property EntityID() As Integer
        '    Get
        '        Return nEntityTypeID
        '    End Get
        'End Property
        Public Property EventID() As Integer
            Get
                Return oLustEventActivityInfo.EventID
            End Get
            Set(ByVal Value As Integer)
                oLustEventActivityInfo.EventID = Value
            End Set
        End Property

        ' The date the LUST Activity was closed
        Public Property Closed() As Date
            Get
                Return oLustEventActivityInfo.Closed
            End Get
            Set(ByVal Value As Date)
                oLustEventActivityInfo.Closed = Value
            End Set
        End Property
        ' The date the LUST Activity was completed
        Public Property Completed() As Date
            Get
                Return oLustEventActivityInfo.Completed
            End Get
            Set(ByVal Value As Date)
                oLustEventActivityInfo.Completed = Value
            End Set
        End Property
        ' The "First GWS Below" for the LUST Activity (associated with REM and GWS activities)
        Public Property First_GWS_Below() As Date
            Get
                Return oLustEventActivityInfo.First_GWS_Below
            End Get
            Set(ByVal Value As Date)
                oLustEventActivityInfo.First_GWS_Below = Value
            End Set
        End Property
        ' (see Frist_GWS_Below)
        Public Property Second_GWS_Below() As Date
            Get
                Return oLustEventActivityInfo.Second_GWS_Below
            End Get
            Set(ByVal Value As Date)
                oLustEventActivityInfo.Second_GWS_Below = Value
            End Set
        End Property
        ' The start date for the LUST Activity
        Public Property Started() As Date
            Get
                Return oLustEventActivityInfo.Started
            End Get
            Set(ByVal Value As Date)
                oLustEventActivityInfo.Started = Value
            End Set
        End Property
        ' The type of the LUST Activity (from tblSYS_PROPERTY_MASTER)
        Public Property Type() As Integer
            Get
                Return oLustEventActivityInfo.Type
            End Get
            Set(ByVal Value As Integer)
                oLustEventActivityInfo.Type = Value
            End Set
        End Property
        ' The deleted flag for the LUST Activity
        Public Property Deleted() As Boolean
            Get
                Return oLustEventActivityInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oLustEventActivityInfo.Deleted = Value
            End Set
        End Property
        Public Property Documents() As MUSTER.Info.LustDocumentCollection
            Get
                Return oLustEventActivityInfo.Documents
            End Get
            Set(ByVal Value As MUSTER.Info.LustDocumentCollection)
                oLustEventActivityInfo.Documents = Value
            End Set
        End Property

        Public Property FacilityID() As Integer
            Get
                Return oLustEventActivityInfo.FacilityID
            End Get
            Set(ByVal Value As Integer)
                oLustEventActivityInfo.FacilityID = Value
            End Set
        End Property

        Public Property UserID() As Integer
            Get
                Return oLustEventActivityInfo.UserID
            End Get
            Set(ByVal Value As Integer)
                oLustEventActivityInfo.UserID = Value
            End Set
        End Property

        Public Property AgeThreshold() As Integer
            Get
                Return oLustEventActivityInfo.AgeThreshold
            End Get
            Set(ByVal Value As Integer)
                oLustEventActivityInfo.AgeThreshold = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oLustEventActivityInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oLustEventActivityInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xLustEventinfo As MUSTER.Info.LustEventInfo
                For Each xLustEventinfo In oLustEventInfo.Activities.Values
                    If xLustEventinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oLustEventActivityInfo.IsDirty = Value
            End Set
        End Property
        Public Property RemSystemID() As Integer
            Get
                Return oLustEventActivityInfo.RemSystemID
            End Get
            Set(ByVal Value As Integer)
                oLustEventActivityInfo.RemSystemID = Value
            End Set
        End Property

        Public Property CreatedBy() As String
            Get
                Return oLustEventActivityInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oLustEventActivityInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oLustEventActivityInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oLustEventActivityInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oLustEventActivityInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oLustEventActivityInfo.ModifiedOn
            End Get
        End Property

#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As MUSTER.Info.LustActivityInfo
            Dim oLustEventActivityInfoLocal As MUSTER.Info.LustActivityInfo
            Try
                If oLustEventInfo.Activities Is Nothing Then
                    oLustEventInfo.Activities = New MUSTER.Info.LustActivityCollection
                End If
                For Each oLustEventActivityInfoLocal In oLustEventInfo.Activities.Values
                    If oLustEventActivityInfoLocal.ActivityID = ID Then
                        If oLustEventActivityInfoLocal.AgeThreshold = 0 Or (oLustEventActivityInfoLocal.IsAgedData = True And oLustEventActivityInfoLocal.IsDirty = False) Then
                            Exit For
                        Else
                            oLustEventActivityInfo = oLustEventActivityInfoLocal
                            oLustEventActivityInfo.FacilityID = onFacilityID
                            oLustEventActivityInfo.UserID = onUserID
                            Return oLustEventActivityInfo
                        End If
                    End If
                Next
                oLustEventActivityInfo = oLustEventActivityDB.DBGetByID(ID)
                oLustEventActivityInfo.FacilityID = onFacilityID
                oLustEventActivityInfo.UserID = onUserID
                If oLustEventActivityInfo.ActivityID = 0 Then
                    'oLustEventActivityInfo.ID = nID
                    nID -= 1
                End If
                oLustEventInfo.Activities.Add(oLustEventActivityInfo)
                Return oLustEventActivityInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Public Function Retrieve(ByVal LustEventName As String) As MUSTER.Info.LustEventInfo
        '    Try
        '        oLustEventActivityInfo = Nothing
        '        If colLustEventActivities.Contains(LustEventName) Then
        '            oLustEventActivityInfo = colLustEventActivities(LustEventName)
        '        Else
        '            If oLustEventActivityInfo Is Nothing Then
        '                oLustEventActivityInfo = New MUSTER.Info.LustEventInfo
        '            End If
        '            oLustEventActivityInfo = oLustEventDB.DBGetByName(LustEventName)
        '            colLustEventActivities.Add(oLustEventActivityInfo)
        '        End If
        '        Return oLustEventActivityInfo
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try

        'End Function
        'Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolUncompletedDoc As Boolean = False)
            Dim strModuleName As String = String.Empty
            Dim bolSubmitForCalendar As Boolean
            Try
                If Me.ValidateData(strModuleName) Then
                    'oLustEventDB.Put(oLustEventActivityInfo)
                    If oLustEventActivityInfo.ActivityID > 0 Then
                        bolSubmitForCalendar = False
                    Else
                        bolSubmitForCalendar = True
                    End If

                    oLustEventActivityDB.Put(oLustEventActivityInfo, moduleID, staffID, returnVal, bolUncompletedDoc)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    oLustEventActivityInfo.Archive()

                    If bolSubmitForCalendar Then
                        CalendarEntries()
                    End If
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Registration") As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = True 'False
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
                    RaiseEvent LustEventErr(errStr)
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
        Function GetAll() As MUSTER.Info.LustActivityCollection
            Try
                oLustEventInfo.Activities.Clear()
                oLustEventInfo.Activities = oLustEventActivityDB.GetAllInfo
                Return oLustEventInfo.Activities
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oLustEventActivityInfo = oLustEventActivityDB.DBGetByID(ID)
                oLustEventActivityInfo.FacilityID = onFacilityID
                oLustEventActivityInfo.UserID = onUserID
                If oLustEventActivityInfo.ActivityID = 0 Then
                    'oLustEventActivityInfo.ID = nID
                    nID -= 1
                End If
                oLustEventInfo.Activities.Add(oLustEventActivityInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oLustEventActivity As MUSTER.Info.LustActivityInfo)
            Try
                oLustEventActivityInfo = oLustEventActivity
                oLustEventActivityInfo.FacilityID = onFacilityID
                oLustEventActivityInfo.UserID = onUserID
                oLustEventInfo.Activities.Add(oLustEventActivityInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oLustEventActivityInfoLocal As MUSTER.Info.LustActivityInfo

            Try
                For Each oLustEventActivityInfoLocal In oLustEventInfo.Activities.Values
                    If oLustEventActivityInfoLocal.ActivityID = ID Then
                        oLustEventInfo.Activities.Remove(oLustEventActivityInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Lust Activity " & ID.ToString & " is not in the collection of Activities.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oLustEventActivity As MUSTER.Info.LustActivityInfo)
            Try
                oLustEventInfo.Activities.Remove(oLustEventActivity)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("LustEventActivity " & oLustEventActivity.ActivityID & " is not in the collection of LustEventActivities.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xLustEventInfo As MUSTER.Info.LustActivityInfo
            For Each xLustEventInfo In oLustEventInfo.Activities.Values
                If xLustEventInfo.IsDirty Then
                    oLustEventActivityInfo = xLustEventInfo
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
                    Me.Save(moduleID, staffID, returnVal)
                End If
            Next
            oLustEventDocuments.Flush(moduleID, staffID, returnVal)
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = oLustEventInfo.Activities.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ActivityID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return oLustEventInfo.Activities.Item(nArr.GetValue(colIndex + direction)).ActivityID.ToString
            Else
                Return oLustEventInfo.Activities.Item(nArr.GetValue(colIndex)).ActivityID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oLustEventActivityInfo = New MUSTER.Info.LustActivityInfo
        End Sub
        Public Sub Reset()
            oLustEventActivityInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oLustEventActivityInfoLocal As New MUSTER.Info.PipeInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("LustEvent ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oLustEventActivityInfoLocal In oLustEventInfo.Activities.Values
                    dr = tbEntityTable.NewRow()
                    dr("Pipe ID") = oLustEventActivityInfoLocal.ID
                    dr("Deleted") = oLustEventActivityInfoLocal.Deleted
                    dr("Created By") = oLustEventActivityInfoLocal.CreatedBy
                    dr("Date Created") = oLustEventActivityInfoLocal.CreatedOn
                    dr("Last Edited By") = oLustEventActivityInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oLustEventActivityInfoLocal.ModifiedOn
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
        Private Sub LustEventInfoChanged() Handles oLustEventActivityInfo.LustActivityInfoChanged
            RaiseEvent LustEventChanged(True)
        End Sub
        'Private Sub TamplateColChanged() Handles colLustEventActivities.LustActivityColChanged
        '    RaiseEvent ColChanged(True)
        'End Sub
#End Region
#Region "Lookup  operations"
        Private Function GetDataTable(ByVal DBViewName As String) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Try
                strSQL = "SELECT * FROM " & DBViewName
                dsReturn = oLustEventActivityDB.DBGetDS(strSQL)
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

        Public Function PopulateLustActivities() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("VTECActivityList")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


        Public Function PopulateLustDocuments(ByVal MGPTFStatus As Integer) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String


            strSQL = "select Document_ID, DocName from dbo.tblTEC_DOCUMENT where Document_ID in (select Document_ID from tblTEC_ACT_DOC_RELATIONSHIP where Activity_ID = " & oLustEventActivityInfo.Type & ") "
            strSQL &= " and NTFE_FLAG = (case " & MGPTFStatus & " when 617 Then 1 when 620 then 1 else NTFE_FLAG end) "
            strSQL &= " and STFS_FLAG = (case " & MGPTFStatus & " when 618 Then 1 when 619 then 1 when 621 then 1 else STFS_FLAG end)"

            Try
                dsReturn = oLustEventActivityDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function

        Public Function GetOpenDocumentCount(ByVal nActivityID As Int64) As Int16
            Dim dsReturn As New DataSet
            Dim nReturn As Int16
            Dim strSQL As String
            Try
                '            strSQL = "SELECT Count(*) as DocCount FROM tblTEC_EVENT_ACTIVITY_DOCUMENT where (Date_Closed is null and Date_Sent_To_Finance is null) and Event_Activity_ID = " & nActivityID
                ' Issue #3179
                ' removed because MM said Date To Financial does not necessarily means the document has been closed.
                strSQL = "SELECT Count(*) as DocCount FROM tblTEC_EVENT_ACTIVITY_DOCUMENT where Date_Closed is null and Event_Activity_ID = " & nActivityID
                dsReturn = oLustEventActivityDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    nReturn = dsReturn.Tables(0).Rows(0)("DocCount")
                Else
                    nReturn = 0
                End If
                Return nReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
   

#End Region
        Public Function CalendarEntries()

            Dim bolAddCalendarEntry As Boolean = False

            Dim dtNotificationDate As Date = Now()
            Dim dtDueDate As Date
            Dim nColorCode
            Dim strTaskDesc As String
            Dim strUserID As String = ""
            Dim strSourceUserID As String = "SYSTEM"
            Dim strGroupID As String = ""
            Dim bolDuetoMe As Boolean = False
            Dim bolToDo As Boolean = False
            Dim bolCompleted As Boolean = False
            Dim bolDeleted As Boolean = False
            Dim oLustEvent As New MUSTER.BusinessLogic.pLustEvent

            'oCalendar.Retrieve(nEntityTypeID, ActivityID, Nothing, Nothing)
            'oCalendar.MarkToDoDeleted(nEntityTypeID, ActivityID)
            'oCalendar.MarkDueToMeDeleted(nEntityTypeID, ActivityID)

            If oLustEventActivityInfo.Type = 698 Then 'TF Checklist
                oLustEvent.Retrieve(oLustEventActivityInfo.EventID)
                ' #1599 do not add cal if event already of type stfs or stfs-direct
                If Not (oLustEvent.MGPTFStatus = 618 Or oLustEvent.MGPTFStatus = 619) Then
                    bolAddCalendarEntry = True
                    strTaskDesc = "TF Paperwork:  Facility:  " & oLustEvent.FacilityID & " - Lust Event:  " & oLustEvent.EVENTSEQUENCE
                    strGroupID = "Financial"
                    strUserID = ""
                    dtDueDate = Started 'DateAdd(DateInterval.Day, 60, Started)
                    bolToDo = True
                End If
            End If




            If bolAddCalendarEntry Then
                'Create a Calendar Info object 
                Dim oCalendarInfo As MUSTER.Info.CalendarInfo
                oCalendarInfo = New MUSTER.Info.CalendarInfo(0, _
                                                dtNotificationDate, _
                                                dtDueDate, _
                                                nColorCode, _
                                                strTaskDesc, _
                                                strUserID, _
                                                strSourceUserID, _
                                                strGroupID, _
                                                bolDuetoMe, _
                                                bolToDo, _
                                                bolCompleted, _
                                                bolDeleted, _
                                                "SYSTEM", _
                                                Now(), _
                                                "SYSTEM", _
                                                Now())

                oCalendarInfo.OwningEntityID = ActivityID
                oCalendarInfo.OwningEntityType = 23
                oCalendarInfo.IsDirty = True
                oCalendar.Add(oCalendarInfo)
                oCalendar.Flush()
            End If


        End Function
    End Class
End Namespace
