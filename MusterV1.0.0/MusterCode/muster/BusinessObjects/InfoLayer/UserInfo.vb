'-------------------------------------------------------------------------------
' MUSTER.Info.UserInfo
'   Provides the container to persist MUSTER User state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        PN      12/3/04    Original class definition.
'  1.1        AN      12/30/04    Added Try catch and Exception Handling/Logging
'  1.2        JC      01/02/05    Added notifications for value changes.
'  1.3        JC      01/12/05    Added call to raise data change event on reset.
'  1.4        JC      01/31/2005  Added function CheckPwd which compares the current
'                                  encrypted password to the original encrypted password.
'  1.5        AB      02/22/05    Added AgeThreshold and IsAgedData Attributes
'  1.6        AN      06/03/2005  Added HEAD Flags
'  1.7        JVC     06/12/2005  Added Active flag
'
' Function          Description
' New()             Instantiates an empty UserInfo object.
' New(ID, Name, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
'                   Instantiates a populated UserInfo object.
' New(dr)           Instantiates a populated UserInfo object taking member state
'                       from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                       last saved to the repository.
'
'
'Attribute          Description
' ID                The unique identifier associated with the User in the repository.
' Name              The name of the User.
' ManagerID
' ManagerName
' UserKey
' ShowDeleted
' DefaultModule
' EmailAddress
' PhoneNumber
' Password
' Deleted
' IsDirty           Indicates if the user state has been altered since it was
'                       last loaded from or saved to the repository.
'-------------------------------------------------------------------------------
'
' TODO - Add to app 1/3/2005 - JVC 2
' TODO - check properties and operations against list.
'

Namespace MUSTER.Info

    <Serializable()> _
Public Class UserInfo

#Region "Private member variables"
        Private nStaffID As Integer             'Some MDEQ-UST ID Number
        Private strUserName As String           'The user's name (human, not system)
        Private strUserID As String             'The user's system name (login ID)
        Private nParentStaffID As Integer       'The user's manager's MDEQ-UST ID Number
        Private strParentUserName As String     'The user's manager's system identification
        Private nEntityID As Integer            'The Entity ID associated with a user entity
        Private strCurrentModule As String      'The user's current module
        ' Private oUserAuditTrail As Collection   'The audit trail associated with the user
        Private strPassword As String          'The user's unencrypted password
        Private strPassKey As String            'The user's encrypted password
        Private bolValidUser As Boolean = False 'Indicator that the user has passed authentication
        Private bolActive As Boolean = True     'Indicator that user is current and active in system
        Private strEmail As String              'The user's e-mail address
        Private strPhoneNumber As String        'The user's MDEQ Phone Number
        'Private strDefaultModule As String      'The user's default module
        Private nDefaultModule As Integer       'The user's default module
        Private strEncryptedPassword As String  'The user's encrypted password
        'Private arrSupervisedUsers As ArrayList 'The list of users supervised by this user

        'Accessories
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime = DateTime.Now.ToShortDateString
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime = DateTime.Now.ToShortDateString

        Private bolDeleted As Boolean            'The user's Deleted flag setting (active/inactive keys on this)
        ' Base state vars
        Private onStaffID As Integer            'Some MDEQ-UST ID Number
        Private ostrUserName As String          'The user's system identification
        Private ostrUserID As String            ' The user's system name
        Private onParentStaffID As Integer      'The user's manager's MDEQ-UST ID Number
        Private ostrParentUserName As String    'The user's manager's system identification
        'Private nUserGroup As Integer           
        'Private strUserGroup As String
        Private onEntityID As Integer           'The Entity ID associated with a user entity
        Private ostrCurrentModule As String     'The user's current module
        'Private ooUserAuditTrail As Collection  'The audit trail associated with the user
        Private ostrPassword As String          'The user's unencrypted password
        Private ostrPassKey As String           'The user's encrypted password
        Private obolValidUser As Boolean = False 'Indicator that the user has passed authentication
        Private obolActive As Boolean = True     'Indicator that user is current and active in system
        Private ostrEmail As String             'The user's e-mail address
        Private ostrPhoneNumber As String       'The user's MDEQ Phone Number
        'Private ostrDefaultModule As String     'The user's default module
        Private onDefaultModule As Integer
        Private ostrEncryptedPassword As String
        'Private oarrSupervisedUsers As ArrayList
        Private obolDeleted As Boolean           'The user's Deleted flag setting (active/inactive keys on this)
        Private bolShowDeleted As Boolean = False
        Private bolIsDirty As Boolean = False
        'Private colUserGroups As UserGroups     'The User Groups that the user is associated to...
        'Private colSupervisedUsers As SupervisedUsers
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private MusterException As New MUSTER.Exceptions.MusterExceptions

        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As DateTime = DateTime.Now.ToShortDateString
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As DateTime = DateTime.Now.ToShortDateString

        Private bolHEAD_PM As Boolean
        Private bolHEAD_CLOSURE As Boolean
        Private bolHEAD_REGISTRATION As Boolean
        Private bolHEAD_INSPECTION As Boolean
        Private bolHEAD_CANDE As Boolean
        Private bolHEAD_FEES As Boolean
        Private bolHEAD_FINANCIAL As Boolean
        Private bolHEAD_ADMIN As Boolean
        Private bolEXECUTIVE_DIRECTOR As Boolean

        Private obolHEAD_PM As Boolean
        Private obolHEAD_CLOSURE As Boolean
        Private obolHEAD_REGISTRATION As Boolean
        Private obolHEAD_INSPECTION As Boolean
        Private obolHEAD_CANDE As Boolean
        Private obolHEAD_FEES As Boolean
        Private obolHEAD_FINANCIAL As Boolean
        Private obolHEAD_ADMIN As Boolean
        Private obolEXECUTIVE_DIRECTOR As Boolean

        Private colUserGroupRelation As MUSTER.Info.UserGroupRelationsCollection
        Private colManagedUsers As MUSTER.Info.UserCollection
#End Region
#Region "Public Events"
        Public Event UserChanged(ByVal bolValue As Boolean)
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Init()
            InitCollection()
            dtDataAge = Now()
            'arrSupervisedUsers = New ArrayList
            'oarrSupervisedUsers = New ArrayList
        End Sub
        Sub New(ByVal StaffID As Integer, _
                ByVal UserID As String, _
                ByVal UserName As String, _
                ByVal EmailAddress As String, _
                ByVal PhoneNumber As String, _
                ByVal DefaultModule As Integer, _
                ByVal ManagerID As Integer, _
                ByVal Password As String, _
                ByVal Deleted As Boolean, _
                ByVal HEADPM As Boolean, _
                ByVal HEADCLOSURE As Boolean, _
                ByVal HEADREGISTRATION As Boolean, _
                ByVal HEADINSPECTION As Boolean, _
                ByVal HEADCANDE As Boolean, _
                ByVal HEADFEES As Boolean, _
                ByVal HEADFINANCIAL As Boolean, _
                ByVal HEADADMIN As Boolean, _
                ByVal ACTIVE As Boolean, _
                ByVal CreatedBy As String, _
                ByVal CreatedOn As Date, _
                ByVal ModifiedBy As String, _
                ByVal LastEdited As Date, _
                ByVal EXECDIRECTOR As Boolean)

            onStaffID = StaffID
            ostrUserID = UserID
            ostrUserName = UserName
            onParentStaffID = ManagerID
            'ostrParentUserName = strParentUserName
            onEntityID = 14
            'ostrCurrentModule = strCurrentModule
            ostrPassword = String.Empty
            ostrPassKey = Password
            ostrEncryptedPassword = Password
            obolValidUser = bolValidUser
            ostrEmail = EmailAddress
            ostrPhoneNumber = PhoneNumber
            onDefaultModule = DefaultModule
            obolDeleted = Deleted
            obolActive = ACTIVE

            obolHEAD_PM = HEADPM
            obolHEAD_CLOSURE = HEADCLOSURE
            obolHEAD_REGISTRATION = HEADREGISTRATION
            obolHEAD_INSPECTION = HEADINSPECTION
            obolHEAD_CANDE = HEADCANDE
            obolHEAD_FEES = HEADFEES
            obolHEAD_FINANCIAL = HEADFINANCIAL
            obolHEAD_ADMIN = HEADADMIN
            obolEXECUTIVE_DIRECTOR = EXECDIRECTOR
            obolActive = ACTIVE

            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = LastEdited
            dtDataAge = Now()
            InitCollection()
            Me.Reset()
            'arrSupervisedUsers = New ArrayList
            'oarrSupervisedUsers = New ArrayList
        End Sub
        Sub New(ByVal drUser As DataRow)
            Try
                onStaffID = drUser.Item("STAFF_ID")
                ostrUserID = drUser.Item("USER_ID")
                ostrUserName = drUser.Item("USER_NAME")
                onParentStaffID = drUser.Item("MANAGER_ID")
                'ostrParentUserName = String.Empty ' This should be a lookup
                onEntityID = 14
                'ostrCurrentModule = String.Empty
                ostrPassword = String.Empty
                ostrPassKey = drUser.Item("PASSWORD")
                'obolValidUser = False
                ostrEncryptedPassword = ostrPassKey
                ostrEmail = drUser.Item("EMAIL_ADDRESS")
                ostrPhoneNumber = drUser.Item("PHONE_NUMBER")
                onDefaultModule = drUser.Item("DEFAULT_MODULE")
                obolDeleted = drUser.Item("DELETED")
                obolActive = drUser.Item("ACTIVE")

                obolHEAD_PM = drUser.Item("HEAD_PM")
                obolHEAD_CLOSURE = drUser.Item("HEAD_CLOSURE")
                obolHEAD_REGISTRATION = drUser.Item("HEAD_REGISTRATION")
                obolHEAD_INSPECTION = drUser.Item("HEAD_INSPECTION")
                obolHEAD_CANDE = drUser.Item("HEAD_CANDE")
                obolHEAD_FEES = drUser.Item("HEAD_FEES")
                obolHEAD_FINANCIAL = drUser.Item("HEAD_FINANCIAL")
                obolHEAD_ADMIN = drUser.Item("HEAD_ADMIN")
                obolEXECUTIVE_DIRECTOR = drUser.Item("EXECUTIVE_DIRECTOR")

                ostrCreatedBy = drUser.Item("CREATED_BY")
                odtCreatedOn = drUser.Item("DATE_CREATED")
                ostrModifiedBy = drUser.Item("LAST_EDITED_BY")
                odtModifiedOn = drUser.Item("DATE_LAST_EDITED")
                InitCollection()
                dtDataAge = Now()
                'arrSupervisedUsers = New ArrayList
                'oarrSupervisedUsers = New ArrayList
                Me.Reset()
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Function CheckPwd() As Boolean
            Return strEncryptedPassword = ostrEncryptedPassword
        End Function
        Public Sub Reset()

            nStaffID = onStaffID
            strUserID = ostrUserID
            strUserName = ostrUserName
            nParentStaffID = onParentStaffID
            'strParentUserName = ostrParentUserName
            nEntityID = onEntityID
            'strCurrentModule = ostrCurrentModule
            strPassword = ostrPassword
            strPassKey = ostrPassKey
            strEncryptedPassword = strPassKey
            bolValidUser = obolValidUser
            strEmail = ostrEmail
            strPhoneNumber = ostrPhoneNumber
            nDefaultModule = onDefaultModule
            bolDeleted = obolDeleted

            bolHEAD_PM = obolHEAD_PM
            bolHEAD_CLOSURE = obolHEAD_CLOSURE
            bolHEAD_REGISTRATION = obolHEAD_REGISTRATION
            bolHEAD_INSPECTION = obolHEAD_INSPECTION
            bolHEAD_CANDE = obolHEAD_CANDE
            bolHEAD_FEES = obolHEAD_FEES
            bolHEAD_FINANCIAL = obolHEAD_FINANCIAL
            bolHEAD_ADMIN = obolHEAD_ADMIN
            bolEXECUTIVE_DIRECTOR = obolEXECUTIVE_DIRECTOR

            bolActive = obolActive

            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn

            bolIsDirty = False
            ResetUserGroupRelationCollection()
            ResetManagedUsersCollection()
            RaiseEvent UserChanged(bolIsDirty)

        End Sub
        Public Sub Archive()

            onStaffID = nStaffID
            ostrUserID = strUserID
            ostrUserName = strUserName
            onParentStaffID = nParentStaffID
            'ostrParentUserName = strParentUserName
            onEntityID = nEntityID
            'ostrCurrentModule = strCurrentModule
            ostrPassword = strPassword
            ostrPassKey = strPassKey
            ostrEncryptedPassword = ostrPassKey
            obolValidUser = bolValidUser
            ostrEmail = strEmail
            ostrPhoneNumber = strPhoneNumber
            onDefaultModule = nDefaultModule
            obolDeleted = bolDeleted

            obolHEAD_PM = bolHEAD_PM
            obolHEAD_CLOSURE = bolHEAD_CLOSURE
            obolHEAD_REGISTRATION = bolHEAD_REGISTRATION
            obolHEAD_INSPECTION = bolHEAD_INSPECTION
            obolHEAD_CANDE = bolHEAD_CANDE
            obolHEAD_FEES = bolHEAD_FEES
            obolHEAD_FINANCIAL = bolHEAD_FINANCIAL
            obolHEAD_ADMIN = bolHEAD_ADMIN
            obolEXECUTIVE_DIRECTOR = bolEXECUTIVE_DIRECTOR

            obolActive = bolActive

            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn

            bolIsDirty = False
        End Sub
        Public Sub ResetUserGroupRelationCollection()
            For Each userGroupRelInfo As MUSTER.Info.UserGroupRelationInfo In colUserGroupRelation.Values
                userGroupRelInfo.Reset()
            Next
        End Sub
        Public Sub ResetManagedUsersCollection()
            For Each userInfo As MUSTER.Info.UserInfo In colManagedUsers.Values
                userInfo.Reset()
            Next
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty

            bolIsDirty = (nStaffID <> onStaffID) Or _
                         (strUserID <> ostrUserID) Or _
                         (strUserName <> ostrUserName) Or _
                         (nParentStaffID <> onParentStaffID) Or _
                         (strParentUserName <> ostrParentUserName) Or _
                         (nEntityID <> onEntityID) Or _
                         (strCurrentModule <> ostrCurrentModule) Or _
                         (strPassword <> ostrPassword) Or _
                         (strPassKey <> ostrPassKey) Or _
                         (strEmail <> ostrEmail) Or _
                         (strPhoneNumber <> ostrPhoneNumber) Or _
                         (nDefaultModule <> onDefaultModule) Or _
                         (bolDeleted <> obolDeleted) Or _
                         (bolHEAD_PM <> obolHEAD_PM) Or _
                         (bolHEAD_CLOSURE <> obolHEAD_CLOSURE) Or _
                         (bolHEAD_REGISTRATION <> obolHEAD_REGISTRATION) Or _
                         (bolHEAD_INSPECTION <> obolHEAD_INSPECTION) Or _
                         (bolHEAD_CANDE <> obolHEAD_CANDE) Or _
                         (bolHEAD_FEES <> obolHEAD_FEES) Or _
                         (bolHEAD_FINANCIAL <> obolHEAD_FINANCIAL) Or _
                         (bolHEAD_ADMIN <> obolHEAD_ADMIN) Or _
                         (bolEXECUTIVE_DIRECTOR <> obolEXECUTIVE_DIRECTOR) Or _
                         (bolActive <> obolActive)

            If Not bolIsDirty Then
                For Each userGroupRelInfo As MUSTER.Info.UserGroupRelationInfo In colUserGroupRelation.Values
                    If userGroupRelInfo.IsDirty Then
                        bolIsDirty = True
                        Exit For
                    End If
                Next
            End If

            If Not bolIsDirty Then
                For Each userInfo As MUSTER.Info.UserInfo In colManagedUsers.Values
                    If userInfo.IsDirty Then
                        bolIsDirty = True
                        Exit For
                    End If
                Next
            End If

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent UserChanged(bolIsDirty)
            End If

        End Sub
        Private Sub Init()

            onStaffID = 0
            ostrUserID = String.Empty
            ostrUserName = String.Empty
            onParentStaffID = 0
            'ostrParentUserName = String.Empty
            onEntityID = 0
            'ostrCurrentModule = String.Empty
            ostrPassword = String.Empty
            ostrPassKey = String.Empty
            ostrEncryptedPassword = String.Empty
            obolValidUser = False
            ostrEmail = String.Empty
            ostrPhoneNumber = String.Empty
            onDefaultModule = 0
            obolDeleted = False
            obolHEAD_PM = False
            obolHEAD_CLOSURE = False
            obolHEAD_REGISTRATION = False
            obolHEAD_INSPECTION = False
            obolHEAD_CANDE = False
            obolHEAD_FEES = False
            obolHEAD_FINANCIAL = False
            obolHEAD_ADMIN = False
            obolEXECUTIVE_DIRECTOR = False
            obolActive = False
            ostrCreatedBy = String.Empty
            odtCreatedOn = DateTime.Now.ToShortDateString
            ostrModifiedBy = String.Empty
            odtModifiedOn = DateTime.Now.ToShortDateString
            InitCollection()
            Me.Reset()
        End Sub
        Private Sub InitCollection()
            colUserGroupRelation = New MUSTER.Info.UserGroupRelationsCollection
            colManagedUsers = New MUSTER.Info.UserCollection
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property UserKey() As Int16
            Get
                Return nStaffID
            End Get
            Set(ByVal Value As Int16)
                nStaffID = Value
            End Set
        End Property
        Public Property Name() As String
            Get
                Return strUserName

            End Get
            Set(ByVal Value As String)
                strUserName = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ID() As String
            Get
                Return strUserID
            End Get
            Set(ByVal Value As String)
                strUserID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ManagerID() As Int16
            Get
                Return nParentStaffID

            End Get
            Set(ByVal Value As Int16)
                nParentStaffID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ManagerName() As String
            Get
                Return strParentUserName

            End Get
            Set(ByVal Value As String)
                strParentUserName = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Active() As Boolean
            Get
                Return bolActive
            End Get
            Set(ByVal Value As Boolean)
                bolActive = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolDeleted = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ShowDeleted() As Boolean
            Get
                Return bolShowDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolShowDeleted = Value
            End Set
        End Property
        Public Property DefaultModule() As Integer
            Get
                Return nDefaultModule
            End Get
            Set(ByVal Value As Integer)
                nDefaultModule = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property EmailAddress() As String
            Get
                Return strEmail
            End Get
            Set(ByVal Value As String)
                strEmail = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PhoneNumber()
            Get
                Return strPhoneNumber
            End Get
            Set(ByVal Value)
                strPhoneNumber = Value
                Me.CheckDirty()
            End Set
        End Property
        Public ReadOnly Property UserModule() As String
            Get
                'TODO -Complete UserAuditPoint 
                'Dim ovar As UserAuditPoint
                'ovar = oUserAuditTrail.Item(oUserAuditTrail.Count)
                'Return ovar.ModuleName
            End Get
        End Property

        Public Property HEAD_PM() As Boolean
            Get
                Return bolHEAD_PM
            End Get
            Set(ByVal Value As Boolean)
                bolHEAD_PM = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property HEAD_CLOSURE() As Boolean
            Get
                Return bolHEAD_CLOSURE
            End Get
            Set(ByVal Value As Boolean)
                bolHEAD_CLOSURE = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property HEAD_REGISTRATION() As Boolean
            Get
                Return bolHEAD_REGISTRATION
            End Get
            Set(ByVal Value As Boolean)
                bolHEAD_REGISTRATION = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property HEAD_INSPECTION() As Boolean
            Get
                Return bolHEAD_INSPECTION
            End Get
            Set(ByVal Value As Boolean)
                bolHEAD_INSPECTION = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property HEAD_CANDE() As Boolean
            Get
                Return bolHEAD_CANDE
            End Get
            Set(ByVal Value As Boolean)
                bolHEAD_CANDE = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property HEAD_FEES() As Boolean
            Get
                Return bolHEAD_FEES
            End Get
            Set(ByVal Value As Boolean)
                bolHEAD_FEES = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property HEAD_FINANCIAL() As Boolean
            Get
                Return bolHEAD_FINANCIAL
            End Get
            Set(ByVal Value As Boolean)
                bolHEAD_FINANCIAL = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property HEAD_ADMIN() As Boolean
            Get
                Return bolHEAD_ADMIN
            End Get
            Set(ByVal Value As Boolean)
                bolHEAD_ADMIN = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property EXECUTIVE_DIRECTOR() As Boolean
            Get
                Return bolEXECUTIVE_DIRECTOR
            End Get
            Set(ByVal Value As Boolean)
                bolEXECUTIVE_DIRECTOR = Value
                Me.CheckDirty()
            End Set
        End Property

        ' The password for the user
        '
        Public WriteOnly Property Password() As String
            'Get
            '    Return strPassword
            'End Get
            Set(ByVal Value As String)
                Dim oEncrypter As New CipherBlock
                Dim oPwdStr As String

                strPassword = Value
                strPassKey = strPassword

                bolValidUser = False
                Try
                    oEncrypter.IV = "MusterApplicationIsFunctional"
                    oEncrypter.Password = strPassword
                    oEncrypter.Provider = CipherBlock.CryptoProviders.DES
                    strEncryptedPassword = oEncrypter.Encrypt(strPassword)
                    strPassKey = strEncryptedPassword
                Catch Ex As Exception
                    MusterException.Publish(Ex, Nothing, Nothing)
                    Throw Ex
                End Try
                Me.CheckDirty()
            End Set
        End Property
        Public ReadOnly Property PasswordKey() As String
            Get
                Return Me.strPassKey
            End Get
        End Property
        'The encrypted password for the user
        '
        Public ReadOnly Property EncryptedPassword() As String
            Get
                Return strEncryptedPassword
            End Get
        End Property
        'Confirm with jay
        Public ReadOnly Property PassKey() As String
            Get
                Return strPassKey
            End Get
        End Property
        Public Property IsDirty() As Boolean
            Get
                If bolIsDirty Then
                    Return bolIsDirty
                Else
                    For Each userGroupRegInfo As MUSTER.Info.UserGroupRelationInfo In colUserGroupRelation.Values
                        If userGroupRegInfo.IsDirty Then
                            Return True
                        End If
                    Next
                    For Each userInfo As MUSTER.Info.UserInfo In colManagedUsers.Values
                        If userInfo.IsDirty Then
                            Return True
                        End If
                    Next
                End If
                Return False
            End Get

            Set(ByVal value As Boolean)
                bolIsDirty = value
            End Set
        End Property
        Public Property AgeThreshold() As Int16
            Get
                Return nAgeThreshold
            End Get

            Set(ByVal value As Int16)
                nAgeThreshold = Int16.Parse(value)
            End Set
        End Property
        Public ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get
        End Property

        Public Property UserGroupRelationCollection() As MUSTER.Info.UserGroupRelationsCollection
            Get
                Return colUserGroupRelation
            End Get
            Set(ByVal Value As MUSTER.Info.UserGroupRelationsCollection)
                colUserGroupRelation = Value
            End Set
        End Property
        Public Property ManagedUsersCollection() As MUSTER.Info.UserCollection
            Get
                Return colManagedUsers
            End Get
            Set(ByVal Value As MUSTER.Info.UserCollection)
                colManagedUsers = Value
            End Set
        End Property

#Region "IAccessors"
        Public Property CreatedBy() As String
            Get
                Return strCreatedBy
            End Get
            Set(ByVal Value As String)
                strCreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return dtCreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
        End Property
#End Region
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region

    End Class
End Namespace
