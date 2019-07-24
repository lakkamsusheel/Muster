'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.User
'   Provides the operations required to manipulate an USery object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         PN      12/4/04     Original class definition.
'   1.1         PN      12/22/04    Removed default password from new and added in property ID
'   1.2         PN      12/28/04    Removed add(ouserinfo) from all the exposed attributes.
'                                   Commented some code in AddToUserGroups method
'                                   Modified RESET()
'   1.3         JC      12/31/04    Added overloaded NEW() which takes a username as the
'                                   sole argument and loads the user object with the
'                                   appropriate information.
'                                   Changed RETRIEVE() functions so they work properly.
'   1.4         JC      01/02/05    Added UserGroup object and events to handle notifications
'                                   on data changes.
'                                   Exposed colIsDirty as a public operation.
'                                   Changed RESET() to operate on only the exposed User object
'                                   Added RESETCOLLECTION() to the object.
'   1.5         AN      01/04/05    Added Try catch and Exception Handling/Logging
'   1.6         JC      01/12/05    Added function IsAMemberOf
'                                   Modified data changed events to or all IsDirty of all 
'                                   all components.
'   1.7         JVC2    01/14/05    Altered FLUSH method to call save method.  It was 
'                                   previously calling UserDB.PUT which was not saving
'                                   the encapsulated objects (groups and supervised users).
'   1.8         JVC2    01/20/05    Modified AddSupervisedUser and DeleteSupervisedUsers to
'                                   remove the reference to a locally instantiated USER
'                                   object.  Also removed other local instantiations that
'                                   were either unnecessary or erroneous.
'   1.9         PN      01/20/05    Modified AddSupervisedUser and DeleteSupervisedUsers.
'                                   Declared oUserInfoLocal under region Private Variables
'                                   Modified resetCollection. Modified  event handler oUserInfo_UserChanged
'                                   with additional handler oUserInfoLocal.UserChanged(bug 652)(J gave permission to do the
'                                   the changes mentioned)  
'   1.10        MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   1.11        JVC2    02/02/05    Added EntityTypeID to private members and initialize to "User" type.
'                                       Also added EntityType attribute to expose the typeID.
'   1.12        AB      02/22/05    Added DataAge check to the Retrieve function
'   1.13        MR      03/01/05    Added ListAllGroups() to Retrieve all the UserGroups.
'   1.14        MR      03/02/05    Modified Flush() to Call GroupMembership Flush method when the UserObject is Not Dirty.
'   1.15        MR      03/02/05    Modified ListUserNames() to add ORDER BY USER_ID in the Select Query.
'   1.16        AB      03/06/05    Added RetrievePMHead
'
' Function          Description
' GetUser(UserID)   Returns the User requested by the string arg UserID
' GetUSer(ID)     Returns the User requested by the int arg ID
' GetUserAll()    Returns an UserCollection with all User objects
' Add(ID)         Adds the User identified by arg ID to the 
'                           internal UserCollection
' Add(UserID)     Adds the User identified by arg UserID to the internal 
'                           UserCollection
' Add(User)       Adds the User passed as the argument to the internal 
'                           UserCollection
' Remove(ID)      Removes the User identified by arg ID from the internal 
'                           UserCollection
' Remove(UserID)  Removes the User identified by arg UserID from the 
'                           internal UserCollection
' UserTable()     Returns a datatable containing all columns for the User 
'                           objects in the internal UserCollection.
' Reset()         Resets the user collection
'
'---------------
'  SPECIAL NOTE
'---------------
'  THE ID attribute in the UserInfo object returns the logon name.
'  THE USERKEY attribute in the UserInfo object returns the STAFF ID Number
'
'  Since the logon must be unique, using ID is permissible for references.
'
'-------------------------------------------------------------------------------
'
' TODO - Check property and operations lists
'

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pUser
#Region "Public Events"
        Public Event UserExists(ByVal MsgStr As String)
        Public Event UserChanged(ByVal bolValue As Boolean)
        Public Event UsersChanged(ByVal bolValue As Boolean)
        Public Event MembershipsChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private WithEvents colUser As Muster.Info.UserCollection
        Private WithEvents oUserInfo As Muster.Info.UserInfo
        Private oUserDB As New Muster.DataAccess.UserDB
        Private bolValidUser As Boolean = False
        'Private WithEvents oMemberships As MUSTER.BusinessLogic.pUserGroupMemberships
        Private nID As Integer = -1
        Private nUserKey As Integer
        Private colUserAuditPoint As Muster.Info.UserAuditPointCollection
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private WithEvents oUserInfoLocal As Muster.Info.UserInfo
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("User").ID
#End Region
#Region "Constructors"
        Public Sub New()
            oUserInfo = New Muster.Info.UserInfo
            'oMemberships = New MUSTER.BusinessLogic.pUserGroupMemberships
            colUserAuditPoint = New Muster.Info.UserAuditPointCollection
            colUser = New Muster.Info.UserCollection
        End Sub
        Public Sub New(ByVal UserName As String)
            oUserInfo = New Muster.Info.UserInfo
            'oMemberships = New MUSTER.BusinessLogic.pUserGroupMemberships(UserName)
            colUserAuditPoint = New Muster.Info.UserAuditPointCollection
            colUser = New Muster.Info.UserCollection
            Me.Retrieve(UserName)
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property UserKey() As Int16
            Get
                Return oUserInfo.UserKey
            End Get
            Set(ByVal Value As Int16)
                oUserInfo.UserKey = Value
            End Set
        End Property
        Public Property Name() As String
            Get
                Return oUserInfo.Name

            End Get
            Set(ByVal Value As String)
                oUserInfo.Name = Value
            End Set
        End Property
        Public Property ID() As String
            Get
                Return oUserInfo.ID
            End Get
            Set(ByVal Value As String)
                oUserInfo.ID = Value
                oUserInfo.Password = "Password"
            End Set
        End Property
        Public Property ManagerID() As Int16
            Get
                Return oUserInfo.ManagerID

            End Get
            Set(ByVal Value As Int16)
                oUserInfo.ManagerID = Value
            End Set
        End Property
        'Don't need this property,I guess
        Public Property ManagerName() As String
            Get
                Return oUserInfo.ManagerName

            End Get
            Set(ByVal Value As String)
                oUserInfo.ManagerName = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oUserInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oUserInfo.Deleted = Value
            End Set
        End Property
        Public Property Active() As Boolean
            Get
                Return oUserInfo.Active
            End Get
            Set(ByVal Value As Boolean)
                oUserInfo.Active = Value
            End Set
        End Property


        Public Property HEAD_PM() As Boolean
            Get
                Return oUserInfo.HEAD_PM
            End Get
            Set(ByVal Value As Boolean)
                oUserInfo.HEAD_PM = Value
            End Set
        End Property

        Public Property HEAD_CLOSURE() As Boolean
            Get
                Return oUserInfo.HEAD_CLOSURE
            End Get
            Set(ByVal Value As Boolean)
                oUserInfo.HEAD_CLOSURE = Value
            End Set
        End Property

        Public Property HEAD_REGISTRATION() As Boolean
            Get
                Return oUserInfo.HEAD_REGISTRATION
            End Get
            Set(ByVal Value As Boolean)
                oUserInfo.HEAD_REGISTRATION = Value
            End Set
        End Property

        Public Property HEAD_INSPECTION() As Boolean
            Get
                Return oUserInfo.HEAD_INSPECTION
            End Get
            Set(ByVal Value As Boolean)
                oUserInfo.HEAD_INSPECTION = Value
            End Set
        End Property

        Public Property HEAD_CANDE() As Boolean
            Get
                Return oUserInfo.HEAD_CANDE
            End Get
            Set(ByVal Value As Boolean)
                oUserInfo.HEAD_CANDE = Value
            End Set
        End Property

        Public Property HEAD_FEES() As Boolean
            Get
                Return oUserInfo.HEAD_FEES
            End Get
            Set(ByVal Value As Boolean)
                oUserInfo.HEAD_FEES = Value
            End Set
        End Property

        Public Property HEAD_FINANCIAL() As Boolean
            Get
                Return oUserInfo.HEAD_FINANCIAL
            End Get
            Set(ByVal Value As Boolean)
                oUserInfo.HEAD_FINANCIAL = Value
            End Set
        End Property

        Public Property HEAD_ADMIN() As Boolean
            Get
                Return oUserInfo.HEAD_ADMIN
            End Get
            Set(ByVal Value As Boolean)
                oUserInfo.HEAD_ADMIN = Value
            End Set
        End Property

        Public Property EXECUTIVE_DIRECTOR() As Boolean
            Get
                Return oUserInfo.EXECUTIVE_DIRECTOR
            End Get
            Set(ByVal Value As Boolean)
                oUserInfo.EXECUTIVE_DIRECTOR = Value
            End Set
        End Property

        Public Property ShowDeleted() As Boolean
            Get
                Return oUserInfo.ShowDeleted
            End Get
            Set(ByVal Value As Boolean)
                oUserInfo.ShowDeleted = Value
            End Set
        End Property
        Public Property DefaultModule() As Integer
            Get
                Return oUserInfo.DefaultModule
            End Get
            Set(ByVal Value As Integer)
                oUserInfo.DefaultModule = Value
            End Set
        End Property
        Public Property EmailAddress() As String
            Get
                Return oUserInfo.EmailAddress
            End Get
            Set(ByVal Value As String)
                oUserInfo.EmailAddress = Value
            End Set
        End Property
        Public Property PhoneNumber()
            Get
                Return oUserInfo.PhoneNumber
            End Get
            Set(ByVal Value)
                oUserInfo.PhoneNumber = Value
            End Set
        End Property
        Public ReadOnly Property UserModule() As String
            Get
                Dim ovar As MUSTER.Info.UserAuditPoint
                Try
                    ovar = colUserAuditPoint.Item(colUserAuditPoint.Count)
                Catch Ex As Exception
                    If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                    Throw Ex
                End Try
                Return ovar.ModuleName
            End Get
        End Property

        'Public ReadOnly Property ListNonMemberships() As DataTable
        '    Get
        '        Try
        '            Return Me.UserGroups.ListNonMemberships(Me.ID)
        '        Catch Ex As Exception
        '            If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '            Throw Ex
        '        End Try
        '    End Get
        'End Property
        Public ReadOnly Property ListMemberships() As DataTable
            Get
                ' List the groups the user belongs to
                Try
                    Dim ds As DataSet
                    ds = oUserDB.DBGetDS("SELECT G.GROUP_ID, UG.GROUP_NAME AS USER_GROUP, UG.ACTIVE AS INACTIVE FROM tblSYS_SECURITY_USER_GROUP G LEFT OUTER JOIN tblSYS_USER_GROUPS UG ON UG.GROUP_ID = G.GROUP_ID WHERE G.STAFF_ID = " + oUserInfo.UserKey.ToString)
                    'Return Me.UserGroups.ListMemberships(Me.ID)
                    Return ds.Tables(0)
                Catch Ex As Exception
                    If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                    Throw Ex
                End Try
            End Get
        End Property

        'The encrypted password for the user
        '
        Public ReadOnly Property EncryptedPassword() As String
            Get
                Try
                    Return oUserInfo.EncryptedPassword
                Catch Ex As Exception
                    If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                    Throw Ex
                End Try
            End Get
        End Property
        ' The password for the user
        Public WriteOnly Property Password() As String
            Set(ByVal Value As String)
                oUserInfo.Password = Value
            End Set
        End Property
        'Public ReadOnly Property EntityType() As Integer
        '    Get
        '        Return nEntityTypeID
        '    End Get
        'End Property
        Public Property CreatedBy() As String
            Get
                Return oUserInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oUserInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oUserInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oUserInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oUserInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oUserInfo.ModifiedOn
            End Get
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oUserInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oUserInfo.IsDirty = value
            End Set
        End Property
        'Public Property UserGroups() As MUSTER.BusinessLogic.pUserGroupMemberships
        '    Get
        '        Return Me.oMemberships
        '    End Get

        '    Set(ByVal value As MUSTER.BusinessLogic.pUserGroupMemberships)
        '        Me.oMemberships = value
        '    End Set
        'End Property
        Public Property UserGroupRelationCollection() As MUSTER.Info.UserGroupRelationsCollection
            Get
                Return oUserInfo.UserGroupRelationCollection
            End Get
            Set(ByVal Value As MUSTER.Info.UserGroupRelationsCollection)
                oUserInfo.UserGroupRelationCollection = Value
            End Set
        End Property
        Public Property ManagedUsersCollection() As MUSTER.Info.UserCollection
            Get
                Return oUserInfo.ManagedUsersCollection
            End Get
            Set(ByVal Value As MUSTER.Info.UserCollection)
                oUserInfo.ManagedUsersCollection = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        ' Obtains and returns an entity as called for by name
        Public Function Retrieve(ByVal staffID As Int64) As MUSTER.Info.UserInfo
            Dim bolDataAged As Boolean = False
            Try
                oUserInfo = Nothing
                If colUser.Contains(staffID) Then
                    oUserInfo = colUser(staffID)
                    If oUserInfo.IsAgedData = True And oUserInfo.IsDirty = False Then
                        colUser.Remove(oUserInfo)
                        bolDataAged = True
                    End If
                End If
                If oUserInfo Is Nothing Or bolDataAged Then
                    oUserInfo = New MUSTER.Info.UserInfo

                    oUserInfo = oUserDB.DBGetByID(staffID)
                    colUser.Add(oUserInfo)
                    oUserInfo.ManagedUsersCollection = oUserDB.DBGetManagedUsers(oUserInfo.UserKey, False)
                    oUserInfo.UserGroupRelationCollection = oUserDB.DBGetUserGroupRel(oUserInfo.UserKey)
                End If

                'oMemberships.Clear()
                'oMemberships.Populate(oUserInfo.ID)
                Return oUserInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal userName As String) As MUSTER.Info.UserInfo
            Dim oUserInfoLocal As MUSTER.Info.UserInfo
            Dim bolGetFromDB As Boolean = False
            Dim bolDataAged As Boolean = False

            Try
                For Each oUserInfoLocal In colUser.Values
                    If oUserInfoLocal.ID = userName Then
                        If oUserInfoLocal.IsAgedData = True And oUserInfoLocal.IsDirty = False Then
                            bolDataAged = True
                            bolGetFromDB = True
                        Else
                            oUserInfo = oUserInfoLocal
                        End If
                        Exit For
                    End If
                Next

                If bolDataAged Then
                    colUser.Remove(oUserInfoLocal)
                End If

                If oUserInfo Is Nothing Then
                    bolGetFromDB = True
                Else
                    If oUserInfo.UserKey = 0 Then
                        bolGetFromDB = True
                    End If
                End If


                If bolGetFromDB = True Then
                    oUserInfo = oUserDB.DBGetByName(userName)
                    If oUserInfo.Name <> String.Empty Then
                        colUser.Add(oUserInfo)
                        oUserInfo.ManagedUsersCollection = oUserDB.DBGetManagedUsers(oUserInfo.UserKey, False)
                        oUserInfo.UserGroupRelationCollection = oUserDB.DBGetUserGroupRel(oUserInfo.UserKey)
                    End If
                End If
                'oMemberships.Clear()
                'oMemberships.Populate(oUserInfo.ID)
                Return oUserInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function

        'Obtains and returns an entity as called for by ID
        Public Function RetrievePMHead() As MUSTER.Info.UserInfo
            Try
                Dim userInfoLocal As MUSTER.Info.UserInfo
                userInfoLocal = oUserDB.DBGetPMHead()
                Return userInfoLocal
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function RetrieveClosureHead() As MUSTER.Info.UserInfo
            Try
                Dim userInfoLocal As MUSTER.Info.UserInfo
                userInfoLocal = oUserDB.DBGetClosureHead()
                Return userInfoLocal
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function RetrieveExecutiveDirector() As MUSTER.Info.UserInfo
            Try
                Dim userInfoLocal As MUSTER.Info.UserInfo
                userInfoLocal = oUserDB.DBGetExecutiveDirector()
                Return userInfoLocal
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function RetrieveCAEHead() As MUSTER.Info.UserInfo
            Try
                Dim userInfoLocal As MUSTER.Info.UserInfo
                userInfoLocal = oUserDB.DBGetCAEHead()
                Return userInfoLocal
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function RetrieveFEEHead() As MUSTER.Info.UserInfo
            Try
                Dim userInfoLocal As MUSTER.Info.UserInfo
                userInfoLocal = oUserDB.DBGetFEEHead()
                Return userInfoLocal
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        'Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String, Optional ByVal OverrideRights As Boolean = False)
            Dim xUserInf As MUSTER.Info.UserInfo
            Dim bolNewUser As Boolean = False
            Dim drSupervisedUsers As DataRow
            Dim drUnsupervisedUsers As DataRow
            'Dim nUserKey As Integer
            Try
                If oUserInfo.UserKey <= 0 Then
                    oUserInfo.CreatedBy = UserID
                Else
                    oUserInfo.ModifiedBy = UserID
                End If
                Dim oldStaffID As Int16 = oUserInfo.UserKey
                nUserKey = oUserDB.Put(oUserInfo, moduleID, staffID, returnVal, OverrideRights)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If
                oUserInfo.Archive()
                If oldStaffID <> oUserInfo.UserKey Then
                    Dim oldUGIDs As New Collection
                    Dim userGroupRelInfo As MUSTER.Info.UserGroupRelationInfo
                    For Each userGroupRelInfo In oUserInfo.UserGroupRelationCollection.Values
                        If userGroupRelInfo.IsDirty Then
                            oldUGIDs.Add(userGroupRelInfo.ID)
                        End If
                    Next
                    If Not oldUGIDs Is Nothing Then
                        For index As Integer = 1 To oldUGIDs.Count
                            Dim colKey As String = CType(oldUGIDs.Item(index), String)
                            userGroupRelInfo = oUserInfo.UserGroupRelationCollection.Item(colKey)
                            userGroupRelInfo.StaffID = oUserInfo.UserKey
                            oUserInfo.UserGroupRelationCollection.ChangeKey(colKey, userGroupRelInfo.StaffID.ToString + "|" + userGroupRelInfo.GroupID.ToString)
                        Next
                    End If

                    For Each userInfo As MUSTER.Info.UserInfo In oUserInfo.ManagedUsersCollection.Values
                        If userInfo.IsDirty Then
                            userInfo.ManagerID = oUserInfo.UserKey
                        End If
                    Next
                End If
                FlushUserGroupRel(moduleID, staffID, returnVal, UserID)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If
                FlushManagedUsers(moduleID, staffID, returnVal, UserID)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If
                'Save UserGroups
                'If Not oMemberships Is Nothing Then
                '    'TODO - Should we look at overloading save of memberships?
                '    Me.oMemberships.Flush(moduleID, staffID, returnVal, UserID)
                '    If Not returnVal = String.Empty Then
                '        Exit Sub
                '    End If
                'End If
                'Save SupervisedUsers
                'For Each xUserInf In colUser.Values
                '    If xUserInf.IsDirty Then
                '        If xUserInf.ManagerID < 0 Then
                '            xUserInf.ManagerID = nUserKey
                '        End If
                '    End If
                'Next

                'Flush(moduleID, staffID, returnVal, UserID)
                'If Not returnVal = String.Empty Then
                '    Exit Sub
                'End If

                oUserInfo.IsDirty = False
                'colIsDirty = False
                'IsDirty = False
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        Public Sub LogEntry(ByVal ModuleName As String, ByVal GUID As String)

            Dim LogMember As New MUSTER.Info.UserAuditPoint(ModuleName)
            LogMember.GUID = GUID
            Try
                'oUserAuditTrail.Add(LogMember, GUID)
                colUserAuditPoint.Add(LogMember)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        Public Sub LogExit(ByVal GUID As String)

            Dim ovar As MUSTER.Info.UserAuditPoint
            'ovar = oUserAuditTrail.Item(GUID)
            Try
                ovar = colUserAuditPoint.Item(GUID)
                If Not ovar Is Nothing Then
                    ovar.ExitPoint = Now
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        ' The method used to validate the user to the system.
        ' 
        Public Function VerifyPassword() As Boolean
            Dim oPwdStr As String
            Dim bolvaliduser As Boolean

            bolvaliduser = oUserInfo.CheckPwd
            Return bolvaliduser


        End Function
        'Public Function IsAMemberOf(ByVal strValue As String) As Boolean
        '    Return oMemberships.IsAMemberOf(oUserInfo.ID & "|USER GROUPS|" & strValue & "|NONE")
        'End Function

        Public Sub SaveModuleEntityRel(ByVal moduleID As Integer, ByVal entityType As Integer, ByVal deleted As Boolean)
            Try
                oUserDB.PutModuleEntityRel(moduleID, entityType, deleted)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Sub SyncDB(ByVal strSQL As String)
            ' declaring new instance of db as connection string needs to be refreshed when executing this sub
            Dim oUserDBLocal As New MUSTER.DataAccess.UserDB
            oUserDBLocal.SyncDB(strSQL)
        End Sub
#End Region
#Region "Collection Operations"
        Function GetAll() As MUSTER.Info.UserCollection
            Try
                colUser.Clear()
                colUser = oUserDB.GetAllInfo
                Return colUser
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an user to the collection as called for by ID
        Public Sub Add(ByVal ID As Int64)
            Try
                oUserInfo = oUserDB.DBGetByID(ID)
                colUser.Add(oUserInfo)
                oUserInfo.ManagedUsersCollection = oUserDB.DBGetManagedUsers(oUserInfo.UserKey, False)
                oUserInfo.UserGroupRelationCollection = oUserDB.DBGetUserGroupRel(oUserInfo.UserKey)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an user to the collection as called for by Name
        Public Sub Add(ByVal UserID As String)
            Try
                oUserInfo = oUserDB.DBGetByName(UserID)
                colUser.Add(oUserInfo)
                oUserInfo.ManagedUsersCollection = oUserDB.DBGetManagedUsers(oUserInfo.UserKey, False)
                oUserInfo.UserGroupRelationCollection = oUserDB.DBGetUserGroupRel(oUserInfo.UserKey)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oUser As MUSTER.Info.UserInfo)

            Try
                oUserInfo = oUser
                If oUserInfo.UserKey = 0 Then
                    oUserInfo.UserKey = nID
                    nID -= 1
                End If
                colUser.Add(oUserInfo)
                oUserInfo.ManagedUsersCollection = oUserDB.DBGetManagedUsers(oUserInfo.UserKey, False)
                oUserInfo.UserGroupRelationCollection = oUserDB.DBGetUserGroupRel(oUserInfo.UserKey)

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Removes the user called for by ID from the collection
        Public Sub Remove(ByVal ID As Int64)
            Try
                If colUser.Contains(ID) Then
                    colUser.Remove(ID)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'Throw New Exception("Report " & ID.ToString & " is not in the collection of reports.")
        End Sub
        'Removes the user called for by Name from the collection
        Public Sub Remove(ByVal UserID As String)
            Try
                For Each oUserInf As MUSTER.Info.UserInfo In colUser.Values
                    If oUserInf.ID = UserID Then
                        colUser.Remove(oUserInf)
                        Exit Sub
                    End If
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'Throw New Exception("User " & Name & " is not in the collection of Users.")
        End Sub
        'Removes the USer supplied from the collection
        Public Sub Remove(ByVal oUserInf As MUSTER.Info.UserInfo)
            Try
                If colUser.Contains(oUserInf.UserKey) Then
                    colUser.Remove(oUserInf.UserKey)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'Throw New Exception("USer " & oUserInf.Name & " is not in the collection of Users.")
        End Sub
        Public Property colIsDirty() As Boolean
            Get
                Dim xUserInf As MUSTER.Info.UserInfo
                For Each xUserInf In colUser.Values
                    If xUserInf.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)

            End Set
        End Property
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim xUserInf As MUSTER.Info.UserInfo
            For Each xUserInf In colUser.Values
                If xUserInf.IsDirty Then
                    'If xUserInf.ManagerID < 0 Then
                    '    xUserInf.ManagerID = nUserKey
                    'End If
                    oUserInfo = xUserInf
                    Me.Save(moduleID, staffID, returnVal, UserID)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    xUserInf.IsDirty = False
                End If
            Next
            'Adam Nall - Added to change the pProfile isDirty. It changes this on the save 
            '            of the single profileInfo but not the parent class isDirty
            'MR - Added to Save Membership Changes when User Object is Not Dirty.
            'If Not oMemberships Is Nothing Then
            '    Me.oMemberships.Flush(moduleID, staffID, returnVal, UserID)
            '    If Not returnVal = String.Empty Then
            '        Exit Sub
            '    End If
            'End If
            Me.IsDirty = False
        End Sub
        Private Sub FlushUserGroupRel(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            For Each userGroupRelInfo As MUSTER.Info.UserGroupRelationInfo In oUserInfo.UserGroupRelationCollection.Values
                If userGroupRelInfo.IsDirty And Not (userGroupRelInfo.isNew And userGroupRelInfo.Deleted) Then
                    oUserDB.PutUserGroupRel(userGroupRelInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    userGroupRelInfo.Archive()
                End If
            Next
        End Sub
        Private Sub FlushManagedUsers(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            For Each userInfo As MUSTER.Info.UserInfo In oUserInfo.ManagedUsersCollection.Values
                If userInfo.IsDirty Then
                    oUserDB.Put(userInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    userInfo.Archive()
                End If
            Next
        End Sub
        'Public Sub AddSupervisedUsers(ByVal UserID As String)
        '    'Dim oUserInfoLocal As Muster.Info.UserInfo
        '    Try
        '        If colUser.Contains(UserID) Then
        '            oUserInfoLocal = colUser.Item(UserID)
        '            'P1 12/19/04 start
        '            'oUserInfoLocal.ManagerID = oUserInfo.UserKey
        '            If oUserInfo.UserKey <= 0 Then
        '                oUserInfoLocal.ManagerID = nID
        '            Else
        '                oUserInfoLocal.ManagerID = oUserInfo.UserKey
        '            End If
        '        Else
        '            oUserInfoLocal = oUserDB.DBGetByName(UserID)
        '            oUserInfoLocal.ManagerID = oUserInfo.UserKey
        '            If oUserInfo.UserKey <= 0 Then
        '                oUserInfoLocal.ManagerID = nID
        '            Else
        '                oUserInfoLocal.ManagerID = oUserInfo.UserKey
        '            End If
        '            colUser.Add(oUserInfoLocal)
        '        End If
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Sub
        ''Sets the manager Id to 0 for the record that is deleted in the collection 
        'Public Sub DeleteSupervisedUsers(ByVal UserID As String)
        '    'Dim myIndex As Int16 = 1
        '    'Dim oUserInfoLocal As Muster.Info.UserInfo

        '    Try
        '        If colUser.Contains(UserID) Then
        '            oUserInfoLocal = colUser.Item(UserID)
        '            oUserInfoLocal.ManagerID = 0
        '        End If
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try

        '    'Throw New Exception("User " & Name & " is not in the collection of Users.")


        'End Sub
        'Public Sub AddToUserGroup(ByVal strGroupName As String)
        '    Try
        '        oMemberships.AddMembership = strGroupName
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Sub
        ' Remove the user from a user group
        'Public Sub RemoveFromUserGroup(ByVal strGroupName As String)
        '    Try
        '        oMemberships.RemoveMembership = strGroupName
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Sub

        Public Function Items() As MUSTER.Info.UserCollection
            Return colUser
        End Function
        Public Function Values() As ICollection
            Return colUser.Values
        End Function
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colUser.GetKeys()
            'Dim nArr(strArr.GetUpperBound(0)) As Integer
            'Dim y As String
            'For Each y In strArr
            '    nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            'Next
            'nArr.Sort(nArr)
            colIndex = Array.BinarySearch(strArr, Me.UserKey.ToString)
            If colIndex + direction > -1 And _
                colIndex + direction <= strArr.GetUpperBound(0) Then
                Return colUser.Item(strArr.GetValue(colIndex + direction)).UserKey.ToString
            Else
                Return colUser.Item(strArr.GetValue(colIndex)).UserKey.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Function ListModulesUserHasAccessTo(ByVal staffID As Integer) As DataTable
            Dim ds As DataSet
            Dim strSQL As String = ""
            Try
                strSQL = "SELECT * FROM vMODULESUSERHASACCESSTO WHERE STAFF_ID = " + staffID.ToString
                'strSQL = "SELECT PROPERTY_ID, PROPERTY_NAME FROM tblSYS_PROPERTY_MASTER WHERE PROPERTY_ID IN (" + _
                '           "SELECT DISTINCT MODULE_ID FROM tblSYS_SECURITY_GROUP_MODULE WHERE DELETED = 0 AND (WRITE_ACCESS = 1 OR READ_ACCESS = 1) AND GROUP_ID IN (" + _
                '              "SELECT GROUP_ID FROM tblSYS_SECURITY_USER_GROUP WHERE DELETED = 0 AND STAFF_ID = " + staffID.ToString + ")" + _
                '              ")"

                ds = oUserDB.DBGetDS(strSQL)
                Return ds.Tables(0)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function ListModulesUserCanSearch(ByVal staffID As Integer) As DataTable
            Dim ds As DataSet
            Dim strSQL As String = ""
            Try
                strSQL = "SELECT * FROM vSEARCHABLEMODULES WHERE PROPERTY_ID IN (" + _
                        "SELECT PROPERTY_ID FROM vMODULESUSERHASACCESSTO WHERE STAFF_ID = " + staffID.ToString + ")"
                ds = oUserDB.DBGetDS(strSQL)
                Return ds.Tables(0)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function ListUsersToSend() As DataTable
            Dim ds As DataSet
            Dim strSQL As String = ""
            Try
                strSQL = "spGetUsersToSendTicklers"

                ds = oUserDB.DBGetDS(String.Format("exec {0}", strSQL))
                Return ds.Tables(0)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function ListModuleEntityRel() As DataSet
            Dim ds As DataSet
            Dim dsRel As DataRelation
            Try
                ds = oUserDB.DBGetDS("SELECT DISTINCT MODULE_ID AS MODULE FROM tblSYS_SECURITY_MODULE_ENTITY; SELECT MODULE_ID AS MODULE, ENTITY_TYPE AS ENTITY FROM tblSYS_SECURITY_MODULE_ENTITY WHERE DELETED = 0")
                dsRel = New DataRelation("ModuleToEntity", ds.Tables(0).Columns("MODULE"), ds.Tables(1).Columns("MODULE"), False)
                ds.Relations.Add(dsRel)
                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function ListEntityTypes() As DataTable
            Dim ds As DataSet
            Try
                ds = oUserDB.DBGetDS("SELECT ENTITY_ID, ENTITY_NAME FROM tblSYS_ENTITY ORDER BY ENTITY_NAME")
                Return ds.Tables(0)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function ListUserNames(Optional ByVal bolDeleted As Boolean = False) As DataTable
            Dim dtUserNames As DataTable
            'colUser = Me.GetAll()
            'dtUserNames = UserCombo()
            Dim strSQL As String
            Dim dsset As New DataSet

            strSQL = "SELECT STAFF_ID, USER_ID FROM tblSYS_UST_STAFF_MASTER WHERE ISNULL(DELETED,0) = " + IIf(bolDeleted, "1", "0")
            strSQL += " ORDER BY USER_ID"

            Try
                dsset = oUserDB.DBGetDS(strSQL)
                Return dsset.Tables(0)
                'If dsset.Tables(0).Rows.Count > 0 Then
                '    dtUserNames = dsset.Tables(0)
                'Else
                '    dtUserNames = Nothing
                'End If
                'Return dtUserNames
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function ListPrimaryModules() As DataTable
            Dim ds As DataSet
            Try
                ds = oUserDB.DBGetDS("SELECT PROPERTY_ID, PROPERTY_NAME FROM TBLSYS_PROPERTY_MASTER WHERE PROPERTY_TYPE_ID = 89 ORDER BY PROPERTY_NAME")
                Return ds.Tables(0)
                'If ds.Tables(0).Rows.Count > 0 Then
                '    Return ds.Tables(0)
                'Else
                '    Return Nothing
                'End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function ListSupervisedUsers() As DataTable
            'Dim dtSupervisedUsersList As New DataTable
            'Dim oUserInfoLocal As MUSTER.Info.UserInfo
            Dim strSQL As String
            Dim dsSet As DataSet
            'Dim drRow As DataRow
            strSQL = "SELECT USER_ID, USER_NAME AS USERNAME FROM tblSYS_UST_STAFF_MASTER WHERE (MANAGER_ID=" + Me.UserKey.ToString & ")"
            dsSet = oUserDB.DBGetDS(strSQL)

            Return dsSet.Tables(0)
            'If dsSet.Tables(0).Rows.Count > 0 Then
            '    dtSupervisedUsersList = dsSet.Tables(0)
            '    'For Each drRow In dsSet.Tables(0).Rows
            '    '    Add(drRow.Item("USER_ID"))
            '    'Next
            'Else
            '    dtSupervisedUsersList = Nothing
            'End If
            'Try
            '    If Not colUser Is Nothing Then
            '        colUser.Clear()
            '        colUser = Me.GetAll()
            '    End If
            '    dtSupervisedUsersList.Columns.Add("USER_ID")
            '    dtSupervisedUsersList.Columns.Add("USERNAME")

            '    For Each oUserInfoLocal In colUser.Values
            '        If oUserInfoLocal.ManagerID = Me.UserKey Then
            '            drRow = dtSupervisedUsersList.NewRow()
            '            drRow("USER_ID") = oUserInfoLocal.ID
            '            drRow("USERNAME") = oUserInfoLocal.Name
            '            dtSupervisedUsersList.Rows.Add(drRow)
            '        End If
            '    Next
            '    Return dtSupervisedUsersList
            '    Return dtSupervisedUsersList
            'Catch Ex As Exception
            '    If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
            '    Throw Ex
            'End Try
        End Function

        Public Function ListAllUsers(Optional ByVal bolShowActiveOnly As Boolean = True, Optional ByVal bolShowDeleted As Boolean = False) As DataTable
            Dim strSQL As String
            Dim dsSet As DataSet
            strSQL = "SELECT USER_ID, USER_NAME AS USERNAME FROM tblSYS_UST_STAFF_MASTER WHERE ACTIVE = " + IIf(bolShowActiveOnly, "0", "ACTIVE") + " AND DELETED = " + IIf(bolShowDeleted, "DELETED", "0")
            dsSet = oUserDB.DBGetDS(strSQL)
            Return dsSet.Tables(0)
        End Function

        'Public Function ListUnSupervisedUsers() As DataTable
        '    Dim dtUnSupervisedUsers As New DataTable
        '    'Dim oUserInfoLocal As MUSTER.Info.UserInfo
        '    'Dim drRow As DataRow
        '    Dim strSQL As String
        '    Dim dsSet As DataSet
        '    strSQL = "SELECT USER_ID, USER_NAME AS USERNAME FROM tblSYS_UST_STAFF_MASTER WHERE (MANAGER_ID=0 and staff_id <>" & Me.ManagerID.ToString & ")"
        '    dsSet = oUserDB.DBGetDS(strSQL)

        '    If dsSet.Tables(0).Rows.Count > 0 Then
        '        dtUnSupervisedUsers = dsSet.Tables(0)
        '    Else
        '        dtUnSupervisedUsers = Nothing
        '    End If
        '    'Try
        '    '    If Not colUser Is Nothing Then
        '    '        colUser.Clear()
        '    '        colUser = Me.GetAll()
        '    '    End If
        '    '    dtUnSupervisedUsersList.Columns.Add("USER_ID")
        '    '    dtUnSupervisedUsersList.Columns.Add("USERNAME")
        '    '    dtUnSupervisedUsersList.Columns.Add("STAFF_ID")
        '    '    For Each oUserInfoLocal In colUser.Values
        '    '        If (oUserInfoLocal.ManagerID = 0) And (oUserInfoLocal.UserKey <> Me.ManagerID) Then
        '    '            drRow = dtUnSupervisedUsersList.NewRow()
        '    '            drRow("USER_ID") = oUserInfoLocal.ID
        '    '            drRow("USERNAME") = oUserInfoLocal.Name
        '    '            drRow("staff_ID") = oUserInfoLocal.UserKey
        '    '            dtUnSupervisedUsersList.Rows.Add(drRow)
        '    '        End If
        '    '    Next
        '    '    Return dtUnSupervisedUsersList
        '    'Catch Ex As Exception
        '    '    If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '    '    Throw Ex
        '    'End Try
        'End Function
        Public Function ListAllGroups() As DataTable
            Dim strSQL As String = "SELECT DISTINCT GROUP_NAME AS USER_GROUP, ACTIVE AS INACTIVE FROM tblSYS_USER_GROUPS WHERE DELETED <>1 ORDER BY GROUP_NAME"
            Dim ds As DataSet
            Try
                ds = oUserDB.DBGetDS(strSQL)
                Return ds.Tables(0)
                'If ds.Tables(0).Rows.Count > 0 Then
                '    Return ds.Tables(0)
                'Else
                '    Return Nothing
                'End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function RunSQLQuery(ByVal strSQL As String) As DataSet
            Try
                Return oUserDB.DBGetDS(strSQL)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Sub Clear()
            oUserInfo = New MUSTER.Info.UserInfo
            'TODO-COMPLETE USERPARAMS
            'oReportParams = New MUSTER.BusinessLogic.pReportParams
        End Sub
        Public Sub Reset()
            If Not oUserInfo Is Nothing Then
                oUserInfo.Reset()
            End If
            'oMemberships.Reset()
        End Sub

        Public Sub ResetCollection()
            Dim xUserInfo As MUSTER.Info.UserInfo
            Dim arrUsers As New Collection
            For Each xUserInfo In colUser.Values
                If xUserInfo.IsDirty Then
                    arrUsers.Add(xUserInfo)
                End If
            Next
            If arrUsers.Count > 0 Then
                For Each xUserInfo In arrUsers
                    xUserInfo.Reset()
                    colUser.Add(xUserInfo)
                Next
            End If
            'P1 12/28/04 start
            'If Not oUserInfo Is Nothing Then
            '    oUserInfo = colUser.Item(oUserInfo.ID)
            'End If
            'P1 12/28/04 end
            'oMemberships.ResetCollection()
            'oMemberships.colIsDirty = False
        End Sub
#End Region
#Region "Miscellaneous Operations"
        'Returns a datatable of the users in the collection
        'Public Function UserTable() As DataTable

        '    Dim oUserInfoLocal As MUSTER.Info.UserInfo
        '    Dim dr As DataRow
        '    Dim tbUserTable As New DataTable

        '    Try
        '        tbUserTable.Columns.Add("UserID")
        '        tbUserTable.Columns.Add("UserName")
        '        tbUserTable.Columns.Add("EmailAddress")
        '        tbUserTable.Columns.Add("PhoneNumber")
        '        tbUserTable.Columns.Add("DefaultModule")
        '        tbUserTable.Columns.Add("ManagerID")
        '        'tbUserTable.Columns.Add("Password")
        '        tbUserTable.Columns.Add("Created By")
        '        tbUserTable.Columns.Add("Created On")
        '        tbUserTable.Columns.Add("Modified By")
        '        tbUserTable.Columns.Add("Modified On")

        '        For Each oUserInfoLocal In colUser
        '            dr = tbUserTable.NewRow()
        '            dr("USERID") = oUserInfoLocal.ID
        '            dr("USERNAME") = oUserInfoLocal.Name
        '            dr("EMAILADDRESS") = oUserInfoLocal.EmailAddress
        '            dr("PHONENUMBER") = oUserInfoLocal.PhoneNumber
        '            dr("DEFAULTMODULE") = oUserInfoLocal.DefaultModule
        '            dr("MANAGERID") = oUserInfoLocal.ManagerID
        '            'dr("PASSWORD") = oUserInfoLocal.Password
        '            dr("CREATED BY") = oUserInfoLocal.CreatedBy
        '            dr("Created On") = oUserInfoLocal.CreatedOn
        '            dr("Modified By") = oUserInfoLocal.ModifiedBy
        '            dr("Modified On") = oUserInfoLocal.ModifiedOn
        '            tbUserTable.Rows.Add(dr)
        '        Next
        '        Return tbUserTable
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try

        'End Function
        Public Function ModuleHeadCheck(ByVal HeadField As String, ByVal UserID As String) As Boolean
            Try
                Return oUserDB.ModuleHeadCheck(HeadField, UserID)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function HasAccess(ByVal moduleID As Integer, ByVal staffID As Integer, ByVal EntityID As Integer) As Boolean
            Return oUserDB.HasAccess(moduleID, staffID, EntityID)
        End Function
#End Region
#End Region
        '#Region "Private Operations"
        '#Region "Miscellaneous Operations"
        '        'Returns a two-column datatable of the User in the collection column names USer ID and User Name
        '        Private Function UserCombo() As DataTable

        '            Dim oUserInfoLocal As MUSTER.Info.UserInfo
        '            Dim dr As DataRow
        '            Dim tbUserTable As New DataTable

        '            Try
        '                tbUserTable.Columns.Add("USER_ID")
        '                tbUserTable.Columns.Add("USERNAME")

        '                For Each oUserInfoLocal In colUser.Values
        '                    dr = tbUserTable.NewRow()
        '                    dr("USER_ID") = oUserInfoLocal.ID
        '                    dr("USERNAME") = oUserInfoLocal.Name
        '                    tbUserTable.Rows.Add(dr)
        '                Next
        '                Return tbUserTable
        '            Catch Ex As Exception
        '                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '                Throw Ex
        '            End Try

        '        End Function
        '#End Region
        '#End Region
#Region "External Event Handlers"
        Private Sub oUserInfo_UserChanged(ByVal bolValue As Boolean) Handles oUserInfo.UserChanged, oUserInfoLocal.UserChanged
            RaiseEvent UserChanged(bolValue Or Me.colIsDirty)
            colUser.Add(oUserInfo)
        End Sub

        Private Sub colUser_UserColChanged() Handles colUser.UserColChanged
            RaiseEvent UsersChanged(Me.colIsDirty Or Me.IsDirty)
        End Sub

        'Private Sub oMemberships_MembershipsChanged(ByVal IsDirtyState As Boolean) Handles oMemberships.MembershipsChanged
        '    RaiseEvent MembershipsChanged(IsDirtyState Or Me.IsDirty Or Me.colIsDirty)
        'End Sub
#End Region
    End Class
End Namespace
