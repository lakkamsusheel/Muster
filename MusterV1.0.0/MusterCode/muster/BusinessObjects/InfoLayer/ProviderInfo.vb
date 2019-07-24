'-------------------------------------------------------------------------------
' MUSTER.Info.ProviderInfo
'   Provides the container to persist MUSTER Flag state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MR       5/22/05    Original class definition.
'
' Function          Description
' New()             Instantiates an empty ProviderInfo object
'                   Instantiates a populated ProviderInfo object
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
' AgeThreshold      Indicates the number of minutes old data can be before it should be 
'                        refreshed from the DB.  Data should only be refreshed when Retrieved
'                        and when IsDirty is false
' IsAgedData        Will return true if the data has been held longer than the AgeThreshold
'
'-------------------------------------------------------------------------------


Namespace MUSTER.Info
    <Serializable()> _
    Public Class ProviderInfo
#Region "Private Member Variables"

        'Old Variables
        Private onProviderID As Integer = 0
        Private obolActive As Boolean
        Private ostrProviderName As String
        Private ostrAbbrev As String = String.Empty
        Private ostrDepartment As String = String.Empty
        Private ostrWebsite As String = String.Empty
        Private obolDeleted As Boolean

        'Current Variables
        Private nProviderID As Integer = 0
        Private bolActive As Boolean
        Private strProviderName As String
        Private strAbbrev As String = String.Empty
        Private strDepartment As String = String.Empty
        Private strWebsite As String = String.Empty
        Private bolDeleted As Boolean

        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime = DateTime.Now.ToShortDateString
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime = DateTime.Now.ToShortDateString
        Private ostrCreatedBy As String
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String
        Private odtModifiedOn As DateTime

        Private dtDataAge As DateTime

        Private nAgeThreshold As Int16 = 5

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.new()
        End Sub
        Sub New(ByVal ProviderID As Integer, _
        ByVal Active As Boolean, _
        ByVal ProviderName As String, _
        ByVal Abbrev As String, _
        ByVal Department As String, _
        ByVal Website As String, _
        ByVal Deleted As Boolean, _
        ByVal CreatedBy As String, _
        ByVal dateCreated As Date, _
        ByVal LAST_EDITED_BY As String, _
        ByVal DATE_LAST_EDITED As Date)

            onProviderID = ProviderID
            obolActive = Active
            ostrProviderName = ProviderName
            ostrAbbrev = Abbrev
            ostrDepartment = Department
            ostrWebsite = Website
            obolDeleted = Deleted

            ostrCreatedBy = CreatedBy
            odtCreatedOn = dateCreated
            ostrModifiedBy = LAST_EDITED_BY
            odtModifiedOn = DATE_LAST_EDITED
            dtDataAge = Now()
            Me.Reset()
        End Sub
#End Region


#Region "Exposed Operations"

        Public Sub Reset()
            nProviderID = onProviderID
            bolActive = obolActive
            strProviderName = ostrProviderName
            strAbbrev = ostrAbbrev
            strDepartment = ostrDepartment
            strWebsite = ostrWebsite
            bolDeleted = obolDeleted

            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn

        End Sub
        Public Sub Archive()
            onProviderID = nProviderID
            obolActive = bolActive
            ostrProviderName = strProviderName
            ostrAbbrev = strAbbrev
            ostrDepartment = strDepartment
            ostrWebsite = strWebsite
            obolDeleted = bolDeleted
        End Sub
#End Region

#Region "Private Operations"
        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty
            bolIsDirty = (nProviderID <> onProviderID Or _
                        bolActive <> obolActive Or _
                        strProviderName <> ostrProviderName Or _
                        strAbbrev <> ostrAbbrev Or _
                        strDepartment <> ostrDepartment Or _
                        strWebsite <> ostrWebsite Or _
                        bolDeleted <> obolDeleted)
        End Sub
        Private Sub Init()
            onProviderID = 0
            obolActive = False
            ostrProviderName = String.Empty
            ostrAbbrev = String.Empty
            ostrDepartment = String.Empty
            ostrWebsite = String.Empty
            obolDeleted = False

            ostrCreatedBy = String.Empty
            odtCreatedOn = DateTime.Now.ToShortDateString
            ostrModifiedBy = String.Empty
            odtModifiedOn = DateTime.Now.ToShortDateString

            Me.Reset()
        End Sub
#End Region

#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return nProviderID
            End Get

            Set(ByVal value As Integer)
                nProviderID = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property
        Public Property Active() As Boolean
            Get
                Return bolActive
            End Get

            Set(ByVal value As Boolean)
                bolActive = value
            End Set
        End Property
        Public Property ProviderName() As String
            Get
                Return strProviderName
            End Get
            Set(ByVal Value As String)
                strProviderName = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Abbrev() As String
            Get
                Return strAbbrev
            End Get
            Set(ByVal Value As String)
                strAbbrev = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Department() As String
            Get
                Return strDepartment
            End Get

            Set(ByVal value As String)
                strDepartment = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Website() As String
            Get
                Return strWebsite
            End Get

            Set(ByVal value As String)
                strWebsite = value
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
