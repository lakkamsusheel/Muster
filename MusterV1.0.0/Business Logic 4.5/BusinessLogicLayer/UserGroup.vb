'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.UserGroup
'   Provides the operations required to manipulate a User Group object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         AN      11/29/04    Original class definition.
'   1.1         JC      12/28/04    Updated to add Screens object and added
'                                   events for collection and value changes
'   1.2         JC      01/02/05    Cleaned up for application use
'   1.3         AN      01/03/05    Added Try catch and Exception Handling/Logging
'   1.4         JVC2    01/12/05    Added call to INFO.ARCHIVE in SAVE method.
'   1.5         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   1.6         AB      02/22/05    Added DataAge check to the Retrieve function
'   1.7         AN      06/14/05    Added Active Flag
'   1.8         JVC2    08/08/05    Added DeleteCalendarEntries operation.
'                                       Also modified HasUsers to include tblSYS_PROFILE_INFO.DELETED = 0
'                                       and tblSYS_PROIFLE_INFO.VALUE = Me.Name requirements.
'
' Function          Description
' Get(NAME)   Returns the UserGroup requested by the string arg NAME
' Get(ID)     Returns the UserGroup requested by the int arg ID
' GetAll()    Returns an UserGroupCollection with all UserGroup objects
' Add(ID)           Adds the UserGroup identified by arg ID to the 
'                           internal UserGroupCollection
' Add(Name)         Adds the UserGroup identified by arg NAME to the internal 
'                           UserGroupCollection
' Add(Entity)       Adds the UserGroup passed as the argument to the internal 
'                           UserGroupCollection
' Remove(ID)        Removes the UserGroup identified by arg ID from the internal 
'                           UserGroupCollection
' Remove(NAME)      Removes the UserGroup identified by arg NAME from the 
'                           internal UserGroupCollection
' Events
' Name              Description
' ScreensChanged    Alerts the client that the underlying collection of screens 
'                     for the user group has changed.
' GroupChanged      Alerts the client that the user group object has changed.
' GroupsChanged     Alerts the client that the underlying collection of user groups
'                     has been modified.
'
'TODO ShowDELETED ON USERGROUP
'-------------------------------------------------------------------------------
'
' TODO - Add to app 1/12/2005 - JVC 2
'

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pUserGroup
#Region "Private Member Variables"
        Private WithEvents colUserGroups As Muster.Info.UserGroupCollection
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private WithEvents oUserGroup As Muster.Info.UserGroupInfo
        Private oUserGroupDB As New Muster.DataAccess.UserGroupDB
        'Private WithEvents colGroupScreens As MUSTER.BusinessLogic.pScreens   'The list of forms and their access modes for the user group
        Private bolFlushInProgress As Boolean = False
        Private bolShowDeleted As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Public Events"
        Public Event ScreensChanged(ByVal bolValue As Boolean)
        Public Event GroupChanged(ByVal bolValue As Boolean)
        Public Event GroupsChanged(ByVal bolValue As Boolean)
        Public Event GroupError(ByVal strErr As String, ByVal strSource As String)
#End Region
#Region "Constructors"
        Public Sub New()
            oUserGroup = New Muster.Info.UserGroupInfo
            colUserGroups = New Muster.Info.UserGroupCollection
            'colGroupScreens = New MUSTER.BusinessLogic.pScreens
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int64
            Get
                Return oUserGroup.ID
            End Get

            Set(ByVal value As Int64)
                oUserGroup.ID = Integer.Parse(value)
            End Set
        End Property
        Public Property Name() As String
            Get
                Return oUserGroup.Name
            End Get
            Set(ByVal Value As String)
                If Value = "SYSTEM" Then
                    Throw New Exception("SYSTEM is a reserved User Group Name and may not be used.")
                    Exit Property
                End If
                oUserGroup.Name = Value
            End Set
        End Property
        Public Property Description() As String
            Get
                Return oUserGroup.Description
            End Get
            Set(ByVal Value As String)
                oUserGroup.Description = Value
            End Set
        End Property
        Public Property ShowDeleted() As Boolean
            Get
                Return Me.bolShowDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolShowDeleted = Value
            End Set
        End Property
        'Public Property ScreenPermissions(ByVal ScreenName As String) As String
        '    Get
        '        Try
        '            Return Me.colGroupScreens.Retrieve(Me.Name & "|SCREENS|" & ScreenName & "|NONE").ProfileValue
        '        Catch Ex As Exception
        '            If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '            Throw Ex
        '        End Try
        '    End Get
        '    Set(ByVal Value As String)
        '        Try
        '            Me.colGroupScreens.Retrieve(Me.Name & "|SCREENS|" & ScreenName & "|NONE").ProfileValue = Value
        '        Catch Ex As Exception
        '            If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '            Throw Ex
        '        End Try
        '    End Set
        'End Property
        'Public Property ScreenDeleted(ByVal ScreenName As String) As Boolean
        '    Get
        '        Try
        '            Return Me.colGroupScreens.Retrieve(Me.Name & "|SCREENS|" & ScreenName & "|NONE").Deleted
        '        Catch Ex As Exception
        '            If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '            Throw Ex
        '        End Try
        '    End Get
        '    Set(ByVal Value As Boolean)
        '        Try
        '            Me.colGroupScreens.Retrieve(Me.Name & "|SCREENS|" & ScreenName & "|NONE").Deleted = Value
        '        Catch Ex As Exception
        '            If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '            Throw Ex
        '        End Try
        '    End Set
        'End Property
        'Public ReadOnly Property Screens() As pScreens
        '    Get
        '        Return colGroupScreens
        '    End Get
        'End Property
        Public Property Deleted() As Boolean
            Get
                Return oUserGroup.Deleted
            End Get

            Set(ByVal value As Boolean)
                oUserGroup.Deleted = value
            End Set
        End Property
        Public Property Active() As Boolean
            Get
                Return oUserGroup.Active
            End Get

            Set(ByVal value As Boolean)
                oUserGroup.Active = value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oUserGroup.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oUserGroup.IsDirty = value
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            '
            ' Iterate the collection of groups looking for dirty ones...
            Get
                Dim xGroupInf As MUSTER.Info.UserGroupInfo
                For Each xGroupInf In colUserGroups.Values
                    If xGroupInf.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                '
                ' If none found, then return the state of the screens collection...
                '
                'Return colGroupScreens.colIsDirty
            End Get
            Set(ByVal Value As Boolean)

            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oUserGroup.CreatedBy
            End Get
            Set(ByVal Value As String)
                oUserGroup.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oUserGroup.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oUserGroup.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oUserGroup.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oUserGroup.ModifiedOn
            End Get
        End Property
        Public Property GroupModuleRelationCollection() As MUSTER.Info.GroupModuleRelationsCollection
            Get
                Return oUserGroup.GroupModuleRelationCollection
            End Get
            Set(ByVal Value As MUSTER.Info.GroupModuleRelationsCollection)
                oUserGroup.GroupModuleRelationCollection = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        ' Obtains and returns an entity as called for by name
        Public Function Retrieve(ByVal groupID As Int64) As MUSTER.Info.UserGroupInfo
            Try
                If colUserGroups.Contains(groupID) Then
                    oUserGroup = colUserGroups.Item(groupID)
                    If oUserGroup.IsAgedData = True And oUserGroup.IsDirty = False Then
                        colUserGroups.Add(oUserGroup)
                        Return Retrieve(groupID)
                    Else
                        oUserGroup.GroupModuleRelationCollection = oUserGroupDB.DBGetGroupModuleRel(oUserGroup.ID)
                        Return oUserGroup
                    End If
                Else
                    colUserGroups.Add(oUserGroupDB.DBGetByID(groupID))
                    Return Retrieve(groupID)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal groupName As String) As MUSTER.Info.UserGroupInfo
            Dim oUserGroupInfoLocal As MUSTER.Info.UserGroupInfo
            Dim bolDataAged As Boolean = False
            Try
                For Each oUserGroupInfoLocal In colUserGroups.Values
                    If oUserGroupInfoLocal.Name = groupName Then
                        If oUserGroupInfoLocal.IsAgedData = True And oUserGroupInfoLocal.IsDirty = False Then
                            bolDataAged = True
                            Exit For
                        Else
                            oUserGroup = oUserGroupInfoLocal
                            oUserGroup.GroupModuleRelationCollection = oUserGroupDB.DBGetGroupModuleRel(oUserGroup.ID)
                            Return oUserGroup
                        End If
                    End If
                Next
                If bolDataAged Then
                    colUserGroups.Remove(oUserGroupInfoLocal)
                End If
                oUserGroup = oUserGroupDB.DBGetByName(groupName)
                colUserGroups.Add(oUserGroup)
                oUserGroup.GroupModuleRelationCollection = oUserGroupDB.DBGetGroupModuleRel(oUserGroup.ID)
                'colGroupScreens.Retrieve(oUserGroup.Name & "|SCREENS", bolShowDeleted)
                Return oUserGroup
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        'Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Try
                Dim oldID As Int64 = oUserGroup.ID
                oUserGroupDB.Put(oUserGroup, moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If
                'Me.colGroupScreens.Flush(moduleID, staffID, returnVal, UserID)
                If oldID <> oUserGroup.ID Then
                    Dim oldIDs As New Collection
                    Dim groupModuleRegInfo As MUSTER.Info.GroupModuleRelationInfo
                    For Each groupModuleRegInfo In oUserGroup.GroupModuleRelationCollection.Values
                        oldIDs.Add(groupModuleRegInfo.ID)
                    Next
                    If Not oldIDs Is Nothing Then
                        For index As Integer = 1 To oldIDs.Count
                            Dim colKey As String = CType(oldIDs.Item(index), String)
                            groupModuleRegInfo = oUserGroup.GroupModuleRelationCollection.Item(colKey)
                            groupModuleRegInfo.GroupID = oUserGroup.ID
                            oUserGroup.GroupModuleRelationCollection.ChangeKey(colKey, groupModuleRegInfo.GroupID.ToString + "|" + groupModuleRegInfo.ModuleID.ToString)
                        Next
                    End If
                End If
                'ousergroup.GroupModuleRelationCollection.ChangeKey(
                FlushGroupModuleRel(moduleID, staffID, returnVal, UserID)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If
                oUserGroup.IsDirty = False
                oUserGroup.Archive()
                RaiseEvent GroupChanged(oUserGroup.IsDirty)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#Region "Collection Operations"
        Function GetAll() As MUSTER.Info.UserGroupCollection
            Try
                colUserGroups.Clear()
                colUserGroups = oUserGroupDB.GetAllInfo
                Return colUserGroups
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Int64)
            Try
                oUserGroup = oUserGroupDB.DBGetByID(ID)
                colUserGroups.Add(oUserGroup)
                oUserGroup.GroupModuleRelationCollection = oUserGroupDB.DBGetGroupModuleRel(oUserGroup.ID)
                'colGroupScreens = New MUSTER.BusinessLogic.pScreens
                'colGroupScreens.Retrieve(oUserGroup.Name & "|SCREENS", False)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Adds an entity to the collection as called for by Name
        Public Sub Add(ByVal Name As String)
            Try
                oUserGroup = oUserGroupDB.DBGetByName(Name)
                If oUserGroup.Name <> Name Then
                    oUserGroup.Name = Name
                End If
                colUserGroups.Add(oUserGroup)
                oUserGroup.GroupModuleRelationCollection = oUserGroupDB.DBGetGroupModuleRel(oUserGroup.ID)
                'colGroupScreens = New MUSTER.BusinessLogic.pScreens
                'colGroupScreens.Retrieve(oUserGroup.Name & "|SCREENS", False)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oUserGroupInfo As MUSTER.Info.UserGroupInfo)

            Try
                oUserGroup = oUserGroupInfo
                colUserGroups.Add(oUserGroup)
                oUserGroup.GroupModuleRelationCollection = oUserGroupDB.DBGetGroupModuleRel(oUserGroup.ID)
                'colGroupScreens = New MUSTER.BusinessLogic.pScreens
                'colGroupScreens.Retrieve(oUserGroup.Name & "|SCREENS", False)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Int64)

            Dim myIndex As Int16 = 1
            Dim oUserGroupInfoLocal As MUSTER.Info.UserGroupInfo

            Try
                For Each oUserGroupInfoLocal In colUserGroups.Values
                    If oUserGroupInfoLocal.ID = ID Then
                        colUserGroups.Remove(oUserGroupInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            Throw New Exception("User Group " & ID.ToString & " is not in the collection of user groups.")

        End Sub
        'Removes the entity called for by Name from the collection
        Public Sub Remove(ByVal Name As String)
            Dim myIndex As Int16 = 1

            Try
                colUserGroups.Remove(Name)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            'Throw New Exception("User Group " & Name & " is not in the collection of user groups.")

        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oUserGroupInfoLocal As MUSTER.Info.UserGroupInfo)

            Try
                colUserGroups.Remove(oUserGroupInfoLocal)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            Throw New Exception("User Group " & oUserGroupInfoLocal.Name & " is not in the collection of user groups.")

        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim xGrpInf As MUSTER.Info.UserGroupInfo
            '
            ' Turn off the event handler while the collection is being flushed
            '
            bolFlushInProgress = True
            '
            ' Flush the associated screens
            '
            'colGroupScreens.Flush(moduleID, staffID, returnVal, UserID)
            'If Not returnVal = String.Empty Then
            '    Exit Sub
            'End If
            For Each xGrpInf In colUserGroups.Values
                If xGrpInf.IsDirty Then
                    oUserGroup = xGrpInf
                    If oUserGroup.ID <= 0 Then
                        oUserGroup.CreatedBy = UserID
                    Else
                        oUserGroup.ModifiedBy = UserID
                    End If
                    Me.Save(moduleID, staffID, returnVal, UserID)
                End If
            Next
            '
            ' Turn the event handler on again
            '
            bolFlushInProgress = False
            '
            ' Alert the client that the collection was flushed since we had the event handler turned off...
            '
            RaiseEvent GroupsChanged(colIsDirty)
        End Sub
        Private Sub FlushGroupModuleRel(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            For Each groupModuleRelInfo As MUSTER.Info.GroupModuleRelationInfo In oUserGroup.GroupModuleRelationCollection.Values
                If groupModuleRelInfo.IsDirty Then
                    oUserGroupDB.PutGroupModuleRel(groupModuleRelInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    groupModuleRelInfo.Archive()
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
            Dim strArr() As String = colUserGroups.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colUserGroups.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colUserGroups.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"

        Public Sub Clear()
            oUserGroup = New MUSTER.Info.UserGroupInfo
            'colGroupScreens = New MUSTER.BusinessLogic.pScreens
        End Sub
        Public Sub Reset()
            oUserGroup.Reset()
            'colGroupScreens.ResetCollection()
        End Sub
        Public Sub ResetCollection()
            Dim oTempGrp As MUSTER.Info.UserGroupInfo
            Dim arrGroups As New Collection
            For Each oTempGrp In colUserGroups.Values
                If oTempGrp.IsDirty Then
                    arrGroups.Add(oTempGrp)
                End If
            Next
            If arrGroups.Count > 0 Then
                For Each oTempGrp In arrGroups
                    oTempGrp.Reset()
                    colUserGroups.Add(oTempGrp)
                Next
            End If
            'colGroupScreens.ResetCollection()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        'Returns a datatable of the entities in the collection
        'Public Property UserGroupTable() As DataTable
        '    Get
        '        Dim oUserGroupInfoLocal As Muster.Info.UserGroupInfo
        '        Dim dr As DataRow
        '        Dim dtForms As New DataTable

        '        Try
        '            Return colGroupScreens.ListScreens(oUserGroup.Name)
        '        Catch ex As Exception
        '            Throw ex
        '        End Try
        '    End Get
        '    Set(ByVal dtForms As DataTable)
        '        Dim drRow As DataRow
        '        Dim oPD As Muster.Info.ProfileInfo
        '        For Each drRow In dtForms.Rows
        '            oPD = Me.colGroupScreens.Retrieve(Me.Name & "|SCREENS|" & drRow.Item("Form Name").ToString & "|NONE", True)
        '            oPD.ProfileValue = drRow.Item("ACCESS_MODE")
        '            oPD.Deleted = drRow.Item("INACTIVE")
        '        Next
        '    End Set
        'End Property

        'P1 12/10/04
        'Added by padmaja 
        'Returns a datatable of the UserGroups in the database
        Public Function ListUserGroups() As DataTable
            'TODO Add Back Parameters and if statment
            'Public Function ProfileDataTable(ByVal strValUserID As String, ByVal strValKey As String) As DataTable

            Dim dsSet As DataSet
            Try

                'AN 6/7/2005 - REMOVED - where DELETED <> 1 
                dsSet = oUserGroupDB.DBGetDS("SELECT DISTINCT GROUP_ID,GROUP_DESCRIPTION,GROUP_NAME,DELETED FROM tblSYS_USER_GROUPS where deleted<>1 ORDER BY GROUP_NAME")

                Return dsSet.Tables(0)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Public Function ListScreens() As DataTable
        '    Try
        '        Return colGroupScreens.ListScreens(Me.Name)
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function

        Public Function HasUsers() As Boolean
            Dim dsSet As DataSet
            Try
                'Checks to see if the current usergroup has an users associated to it.
                dsSet = oUserGroupDB.DBGetDS("SELECT * from tblSys_Profile_Info left outer join tblSYS_UST_STAFF_MASTER on tblSYS_UST_STAFF_MASTER.USER_ID = tblSys_Profile_Info.USER_ID  Where tblSYS_UST_STAFF_MASTER.DELETED=0 and Profile_Key='USER GROUPS' and Profile_Modifier_1='" & Me.Name & "' and Profile_Value='" & Me.Name & "'and tblSYS_PROFILE_INFO.DELETED = 0")

                If dsSet.Tables(0).Rows.Count > 0 Then
                    Return True
                Else
                    Return False
                End If

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function HasCalendarEntries() As Boolean
            Dim dsSet As DataSet
            Try
                'Checks to see if the current usergroup has an users associated to it.
                dsSet = oUserGroupDB.DBGetDS("SELECT * from tblSYS_CALENDAR_CALENDAR_INFO WHERE GROUP_ID = '" & Me.Name & "' AND DELETED = 0")
                If dsSet.Tables(0).Rows.Count > 0 Then
                    Return True
                Else
                    Return False
                End If

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub DeleteCalendarEntries()
            Try
                oUserGroupDB.DBGetDS("UPDATE tblSYS_CALENDAR_CALENDAR_INFO set DELETED = 1 WHERE GROUP_ID = '" & Me.Name & "' AND DELETED = 0")
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

#End Region
#End Region
#Region "External Event Handlers"
        'Private Sub GroupScreensChanged(ByVal bolValue As Boolean) Handles colGroupScreens.ScreensChanged
        '    '
        '    ' Alert the client that the underlying collection of screens has changed state!
        '    '
        '    ' Ignore the message if the collection is being flushed...
        '    '
        '    If bolFlushInProgress Then Exit Sub
        '    RaiseEvent ScreensChanged(colGroupScreens.colIsDirty Or Me.IsDirty)
        'End Sub
        Private Sub ThisGroupChanged(ByVal bolValue As Boolean) Handles oUserGroup.UserGroupChanged
            '
            ' Alert the client that the current group data has changed
            '
            RaiseEvent GroupChanged(Me.IsDirty)
        End Sub
        Private Sub GroupColChanged() Handles colUserGroups.UserColChanged
            '
            ' Alert the client that the current group data has changed
            '
            ' Ignore the message if the collection is being flushed...
            '
            If bolFlushInProgress Then Exit Sub
            RaiseEvent GroupsChanged(Me.colIsDirty)
        End Sub
        Private Sub GroupInfoErr(ByVal strErr As String, ByVal strSource As String) Handles oUserGroup.UserGroupErr
            RaiseEvent GroupError(strErr, strSource)
        End Sub
#End Region
    End Class
End Namespace
