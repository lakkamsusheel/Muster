'-------------------------------------------------------------------------------
' MUSTER.Info.UserGroupInfo
'   Provides the container to persist MUSTER UserGroup state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        AN      11/29/04    Original class definition.
'  1.1        JC      12/28/04    Updated Reset to set bolIsDirty to False
'                                 Added event for notification of data changes.
'                                 Added firing of event in CHECKDIRTY() if
'                                   dirty state changed.
'  1.2        AN      12/30/04    Added Try catch and Exception Handling/Logging
'  1.3        JC      01/12/05    Added UserGroupErr event and code to fire when NAME
'                                   attribute is blank.
'                                 Added missing ARCHIVE method.
'  1.4        JVC2    01/17/05    Fixed bug in SET method of NAME attribute.
'  1.5        AN      01/21/05    Added RaiseEvent on reset
'  1.6        AB      02/22/05    Added AgeThreshold and IsAgedData Attributes
'  1.7        AN      06/14/05    Added Active Flag
'
' Function          Description
' New()             Instantiates an empty UserGroupInfo object.
' New(IsNewItem, ID , Name , Description , ShowDeleted , Deleted , CreatedBy , CreatedOn , ModifiedBy , ModifiedOn
'                   Instantiates a populated UserGroupInfo object.
' New(dr)           Instantiates a populated UserGroupInfo object taking member state
'                       from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                       last saved to the repository.
'
'
'Attribute          Description
' ID                The unique identifier associated with the User Group in the repository.
' Name              The name of the User Group.
'TODO Add Attributes
' Deleted
' IsDirty           Indicates if the Entity state has been altered since it was
'                       last loaded from or saved to the repository.
'-------------------------------------------------------------------------------
'
' TODO - 01/12/05 - Check lists of operations and attributes in header.
'
Namespace MUSTER.Info
    <Serializable()> _
      Public Class UserGroupInfo
#Region "Private member variables"

        Private nGroupID As Int64                   'The internal ID number associated with the user group
        Private strName As String                   'The name of the user group
        Private strDescription As String = String.Empty  'The description of the user group
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime = DateTime.Now.ToShortDateString
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime = DateTime.Now.ToShortDateString
        'Private colAssociatedForms As ProfileData   'The list of forms and their access modes for the user group
        Private bolIsNewItem As Boolean

        Private onGroupID As Int64
        Private ostrName As String
        Private ostrDescription As String = String.Empty
        Private bolShowDeleted As Boolean = False
        Private bolDeleted As Boolean
        Private obolDeleted As Boolean
        Private bolActive As Boolean = False
        Private obolActive As Boolean = False
        Private obolIsNewItem As Boolean

        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As DateTime = DateTime.Now.ToShortDateString
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As DateTime = DateTime.Now.ToShortDateString

        Private bolIsDirty As Boolean = False
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private MusterException As New MUSTER.Exceptions.MusterExceptions

        Private WithEvents colGroupModuleRel As MUSTER.Info.GroupModuleRelationsCollection
#End Region
#Region "Public Events"
        Public Event UserGroupChanged(ByVal bolValue As Boolean)
        Public Event UserGroupErr(ByVal strErr As String, ByVal strSource As String)
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            InitCollection()
            dtDataAge = Now()
        End Sub
        Sub New(ByVal IsNewItem As Boolean, _
            ByVal ID As Integer, _
            ByVal Name As String, _
            ByVal Description As String, _
            ByVal ShowDeleted As Boolean, _
            ByVal Deleted As Boolean, _
            ByVal Active As Boolean, _
            ByVal CreatedBy As String, _
            ByVal CreatedOn As Date, _
            ByVal ModifiedBy As String, _
            ByVal ModifiedOn As Date)

            obolIsNewItem = IsNewItem
            onGroupID = ID
            ostrName = Name
            ostrDescription = Description
            bolShowDeleted = ShowDeleted
            obolIsNewItem = IsNewItem
            obolDeleted = Deleted
            obolActive = Active
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = ModifiedOn
            dtDataAge = Now()
            InitCollection()
            Me.Reset()
        End Sub
        'Sub New(ByVal drReport As DataRow)
        '    Try
        '        bolIsNewItem = drReport.Item("REPORT_ID")
        '        onGroupID = drReport.Item("REPORT_ID")
        '        ostrName = drReport.Item("REPORT_ID")
        '        ostrDescription = drReport.Item("REPORT_ID")
        '        bolShowDeleted = drReport.Item("REPORT_ID")
        '        obolDeleted = drReport.Item("REPORT_ID")
        '        obolIsNewItem = drReport.Item("REPORT_ID")
        '        bolDeleted = drReport.Item("REPORT_ID")
        '        obolActive = drReport.Item("ACTIVE")
        '        strCreatedBy = drReport.Item("REPORT_ID")
        '        dtCreatedOn = drReport.Item("REPORT_ID")
        '        strModifiedBy = drReport.Item("REPORT_ID")
        '        dtModifiedOn = drReport.Item("REPORT_ID")
        '        dtDataAge = Now()
        '        InitCollection()
        '    Catch Ex As Exception
        '        MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            strName = ostrName
            strDescription = ostrDescription
            nGroupID = onGroupID
            bolDeleted = obolDeleted
            bolActive = obolActive
            bolIsNewItem = obolIsNewItem
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
            ResetGroupModuleRelationCollection()
            RaiseEvent UserGroupChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            ostrName = strName
            ostrDescription = strDescription
            onGroupID = nGroupID
            obolDeleted = bolDeleted
            obolActive = bolActive
            obolIsNewItem = bolIsNewItem
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn

        End Sub
        Public Sub ResetGroupModuleRelationCollection()
            For Each groupModuleInfo As MUSTER.Info.GroupModuleRelationInfo In colGroupModuleRel.Values
                groupModuleInfo.Reset()
            Next
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()

            Dim bolOldValue As Boolean = bolIsDirty

            bolIsDirty = (strDescription <> ostrDescription) Or _
               (strName <> ostrName) Or _
               (bolDeleted <> obolDeleted) Or _
               (bolActive <> obolActive) Or _
               (strDescription <> ostrDescription) Or _
               (bolIsNewItem <> obolIsNewItem)

            If Not bolIsDirty Then
                For Each groupModuleRel As MUSTER.Info.GroupModuleRelationInfo In colGroupModuleRel.Values
                    If groupModuleRel.IsDirty Then
                        bolIsDirty = True
                        Exit For
                    End If
                Next
            End If

            If bolOldValue <> bolIsDirty Then
                RaiseEvent UserGroupChanged(bolIsDirty)
            End If

        End Sub
        Private Sub Init()

            ostrName = String.Empty
            ostrDescription = String.Empty
            onGroupID = 0
            obolIsNewItem = False
            obolActive = True
            odtCreatedOn = DateTime.Now.ToShortDateString
            odtModifiedOn = DateTime.Now.ToShortDateString
            ostrCreatedBy = String.Empty
            ostrModifiedBy = String.Empty
            InitCollection()
            Me.Reset()

        End Sub
        Private Sub InitCollection()
            colGroupModuleRel = New MUSTER.Info.GroupModuleRelationsCollection
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int64
            Get
                Return nGroupID
            End Get

            Set(ByVal value As Int64)
                nGroupID = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property
        Public Property Name() As String
            Get
                Return strName
            End Get
            Set(ByVal Value As String)
                If Value <> String.Empty Then
                    strName = Value
                    Me.CheckDirty()
                Else
                    RaiseEvent UserGroupErr("User Group name cannot be empty!", "UserGroupInfo")
                End If
            End Set
        End Property
        Public Property Description() As String
            Get
                Return strDescription
            End Get
            Set(ByVal Value As String)
                strDescription = Value
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
        Public Property Active() As Boolean
            Get
                Return bolActive
            End Get

            Set(ByVal value As Boolean)
                bolActive = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                If bolIsDirty Then
                    Return bolIsDirty
                Else
                    For Each groupModuleRegInfo As MUSTER.Info.GroupModuleRelationInfo In colGroupModuleRel.Values
                        If groupModuleRegInfo.IsDirty Then
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
        Public Property GroupModuleRelationCollection() As MUSTER.Info.GroupModuleRelationsCollection
            Get
                Return colGroupModuleRel
            End Get
            Set(ByVal Value As MUSTER.Info.GroupModuleRelationsCollection)
                colGroupModuleRel = Value
            End Set
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
#Region "External Events"
        Private Sub GroupModuleRelationColChanged() Handles colGroupModuleRel.GroupModuleRelationColChanged
            CheckDirty()
        End Sub
#End Region
    End Class
End Namespace

