'-------------------------------------------------------------------------------
' MUSTER.Info.ProfileInfo
'   Provides the container to persist MUSTER Profile data
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC2      11/29/04    Original class definition.
'  1.1        JC        12/28/04    Added event for data update notification.
'                                   Added firing of event in CHECKDIRTY() if
'                                     dirty state changed.
'  1.2        AN        12/30/04    Added Try catch and Exception Handling/Logging
'  1.3        JVC2      01/20/05    Added call to raise event InfoBecameDirty in RESET
'  1.4        AB        02/22/05    Added AgeThreshold and IsAgedData Attributes
'
' Operations
' Function          Description
' New()             Instantiates an empty ProfileInfo object.
' New(User, Key, Mod1, Mod2, Value, Deleted, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
'                   Instantiates a populated ProfileInfo object.
' New(dr)           Instantiates a populated ProfileInfo object taking member state
'                       from the datarow provided.
' Reset()           Sets the object state to the original state when loaded from or
'                       last saved to the repository.
'
' Read-Write Attributes
'Attribute          Description
' ID                The primary key associated with the ProfileInfo in the repository.
'                     Composite of User, Key, Mod1, Mod2 delimited with | characters.
' ProfileKey        The ProfileKey of the ProfileInfo object.
' ProfileMod1       The first modifier for the ProfileKey of the ProfileInfo object.
' ProfileMod2       The second modifier for the ProfileKey of the ProfileInfo object.
' User              The User ID associated with the ProfileKey of the ProfileInfo object.
' ProfileValue      The value associated with the ProfileKey of the ProfileInfo object.
' Deleted           The deleted state of the ProfileInfo object (True = deleted).
' IsDirty           Indicates if the ProfileInfo state has been altered since it was
'                       last loaded from or saved to the repository.
'
' Read-Only Attributes
' CreatedBy         The name of the user that created the ProfileInfo object.
' CreatedOn         The date that the ProfileInfo object was created.
' ModifiedBy        The name of the user that last modified the ProfileInfo object.
' ModifiedOn        The date that the ProfileInfo object was last modified.
'-------------------------------------------------------------------------------
'
' TODO - 12/29 - Check list of Opertaions and Attributes.
'
Namespace MUSTER.Info

    <Serializable()> _
      Public Class ProfileInfo

#Region "Private member variables"

        Private strUserID As String                  'The internal ID number associated with the user group
        Private strKey As String                   'The name of the user group
        Private strModifier1 As String                 'The module the report is associated with
        Private strModifier2 As String            'The description of the report
        Private strProfileValue As String              'The location that the report resides in
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime = DateTime.Now.ToShortDateString
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime = DateTime.Now.ToShortDateString
        Private bolDeleted As Boolean
        'Private colAssociatedGroups As UserGroups   'The list of user groups associated with the form
        Private dtReportNames As DataTable         'Lists the report names for the client
        Private bolFavoriteReport As Boolean = False

        Private ostrUserID As String                  'The internal ID number associated with the user group
        Private ostrKey As String                   'The name of the user group
        Private ostrModifier1 As String                 'The module the report is associated with
        Private ostrModifier2 As String            'The description of the report
        Private ostrProfileValue As String              'The location that the report resides in
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As DateTime = DateTime.Now.ToShortDateString
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As DateTime = DateTime.Now.ToShortDateString
        Private bolShowDeleted As Boolean = False

        Private bolIsDirty As Boolean = False
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Public Events"
        Public Event InfoBecameDirty(ByVal DirtyState As Boolean)
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            dtDataAge = Now()
        End Sub
        Sub New(ByVal UserID As String, _
            ByVal ProfileKey As String, _
            ByVal Modifier1 As String, _
            ByVal Modifier2 As String, _
            ByVal Value As String, _
            ByVal Deleted As Boolean, _
            ByVal CreatedBy As String, _
            ByVal CreatedOn As Date, _
            ByVal ModifiedBy As String, _
            ByVal LastEdited As Date)
            ostrUserID = UserID                  'The internal ID number associated with the user group
            ostrKey = ProfileKey                   'The name of the user group
            ostrModifier1 = Modifier1                'The module the report is associated with
            ostrModifier2 = Modifier2            'The description of the report
            ostrProfileValue = Value             'The location that the report resides in
            obolDeleted = Deleted
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = LastEdited
            dtDataAge = Now()
            Me.Reset()
        End Sub
        Sub New(ByVal drReport As DataRow)
            Try
                ostrUserID = drReport.Item("USER_ID")
                ostrKey = drReport.Item("PROFILE_KEY")
                ostrModifier1 = drReport.Item("PROFILE_MODIFIER_1")
                ostrModifier2 = drReport.Item("PROFILE_MODIFIER_2")
                ostrProfileValue = drReport.Item("PROFILE_VALUE")
                obolDeleted = drReport.Item("DELETED")
                ostrCreatedBy = drReport.Item("CREATED_BY")
                odtCreatedOn = drReport.Item("DATE_CREATED")
                ostrModifiedBy = drReport.Item("LAST_EDITED_BY")
                odtModifiedOn = drReport.Item("DATE_LAST_EDITED")
                dtDataAge = Now()
                Me.Reset()
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"

        Public Sub Archive()

            ostrUserID = strUserID
            ostrKey = strKey
            ostrModifier1 = strModifier1
            ostrModifier2 = strModifier2
            ostrProfileValue = strProfileValue
            obolDeleted = bolDeleted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            bolIsDirty = False
        End Sub
        Public Sub Reset()

            strUserID = ostrUserID
            strKey = ostrKey
            strModifier1 = ostrModifier1
            strModifier2 = ostrModifier2
            strProfileValue = ostrProfileValue
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn

            bolIsDirty = False
            RaiseEvent InfoBecameDirty(bolIsDirty)
        End Sub
#End Region
#Region "Private Operations"

        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty

            bolIsDirty = (strUserID <> ostrUserID) Or _
                         (strModifier1 <> ostrModifier1) Or _
                         (strModifier2 <> ostrModifier2) Or _
                         (strProfileValue <> ostrProfileValue) Or _
                         (strKey <> ostrKey) Or _
                         (bolDeleted <> obolDeleted)

            If bolOldState <> bolIsDirty Then
                RaiseEvent InfoBecameDirty(bolIsDirty)
            End If
        End Sub
        Private Sub Init()

            ostrUserID = String.Empty
            ostrKey = String.Empty
            ostrModifier1 = String.Empty
            ostrModifier2 = String.Empty
            ostrProfileValue = String.Empty
            obolDeleted = False
            dtCreatedOn = DateTime.Now.ToShortDateString
            dtModifiedOn = DateTime.Now.ToShortDateString
            strCreatedBy = String.Empty
            strModifiedBy = String.Empty
            Me.Reset()

        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As String
            Get
                Return strUserID & "|" & strKey & "|" & strModifier1 & "|" & strModifier2
            End Get

            Set(ByVal value As String)
                Try
                    Dim arrVals() As String
                    arrVals = value.Split("|")
                    strUserID = arrVals(0)
                    strKey = arrVals(1)
                    strModifier1 = arrVals(2)
                    strModifier2 = arrVals(3)
                Catch Ex As Exception
                    MusterException.Publish(Ex, Nothing, Nothing)
                    Throw Ex
                End Try
                Me.CheckDirty()
            End Set
        End Property
        Public Property ProfileKey() As String
            Get
                Return strKey
            End Get
            Set(ByVal Value As String)
                strKey = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property ProfileMod1() As String
            Get
                Return strModifier1
            End Get
            Set(ByVal Value As String)
                strModifier1 = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property ProfileMod2() As String
            Get
                Return strModifier2
            End Get
            Set(ByVal Value As String)
                strModifier2 = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property User() As String
            Get
                Return strUserID
            End Get

            Set(ByVal value As String)
                strUserID = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property ProfileValue() As String
            Get
                Return strProfileValue
            End Get

            Set(ByVal value As String)
                strProfileValue = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get

            Set(ByVal value As Boolean)
                bolDeleted = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
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


#Region "iAccessors"
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

