'-------------------------------------------------------------------------------
' MUSTER.Info.CourseInfo
'   Provides the container to persist MUSTER Flag state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MR       5/22/05    Original class definition.
'
' Function          Description
' New()             Instantiates an empty CourseInfo object
'                   Instantiates a populated CourseInfo object
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
    Public Class CourseInfo

#Region "Private Member Variables"

        'Old Variables
        Private onCourseID As Integer = 0
        Private obolActive As Boolean
        Private onProviderID As Integer
        Private ostrCourseTitle As String = String.Empty
        'Private ostrCourseDates As String = String.Empty
        'Private ostrLocation As String = String.Empty
        'Private onCourseTypeID As Integer
        Private ostrProviderName As String = String.Empty
        'Private ostrCreditHours As String = String.Empty
        Private obolDeleted As Boolean


        'Current Variables
        Private nCourseID As Integer = 0
        Private bolActive As Boolean
        Private nProviderID As Integer
        Private strCourseTitle As String = String.Empty
        'Private strCourseDates As String = String.Empty
        'Private strLocation As String = String.Empty
        'Private nCourseTypeID As Integer
        Private strProviderName As String = String.Empty
        'Private strCreditHours As String = String.Empty
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
        Sub New(ByVal CourseID As Integer, _
        ByVal Active As Boolean, _
        ByVal ProviderID As Integer, _
        ByVal CourseTitle As String, _
        ByVal ProviderName As String, _
        ByVal Deleted As Boolean, _
        ByVal CreatedBy As String, _
        ByVal dateCreated As Date, _
        ByVal LAST_EDITED_BY As String, _
        ByVal DATE_LAST_EDITED As Date)

            onCourseID = CourseID
            obolActive = Active
            onProviderID = ProviderID
            ostrCourseTitle = CourseTitle
            'ostrCourseDates = CourseDates
            'ostrLocation = Location
            'onCourseTypeID = CourseTypeID
            ostrProviderName = ProviderName
            'ostrCreditHours = CreditHours
            obolDeleted = Deleted

            ostrCreatedBy = CreatedBy
            odtCreatedOn = dateCreated
            ostrModifiedBy = LAST_EDITED_BY
            odtModifiedOn = DATE_LAST_EDITED
            dtDataAge = Now()
            Me.Reset()
            'ByVal CourseDates As String, _
            '        ByVal Location As String, _
            '        ByVal CourseTypeID As Integer, _
            'ByVal CreditHours As String, _
        End Sub
#End Region


#Region "Exposed Operations"

        Public Sub Reset()
            nCourseID = onCourseID
            bolActive = obolActive
            nProviderID = onProviderID
            strCourseTitle = ostrCourseTitle
            'strCourseDates = ostrCourseDates
            'strLocation = ostrLocation
            'nCourseTypeID = onCourseTypeID
            strProviderName = ostrProviderName
            'strCreditHours = ostrCreditHours
            bolDeleted = obolDeleted

            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn

        End Sub
        Public Sub Archive()
            onCourseID = nCourseID
            obolActive = bolActive
            onProviderID = nProviderID
            ostrCourseTitle = strCourseTitle
            'ostrCourseDates = strCourseDates
            'ostrLocation = strLocation
            'onCourseTypeID = nCourseTypeID
            ostrProviderName = strProviderName
            'ostrCreditHours = strCreditHours
            obolDeleted = bolDeleted
        End Sub

#End Region

#Region "Private Operations"
        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty
            bolIsDirty = (nCourseID <> onCourseID Or _
                        nProviderID <> onProviderID Or _
                        strCourseTitle <> ostrCourseTitle Or _
                        strProviderName <> ostrProviderName Or _
                        bolDeleted <> obolDeleted)
            'strCourseDates <> ostrCourseDates Or _
            '           strLocation <> ostrLocation Or _
            '           nCourseTypeID <> onCourseTypeID Or _
            'strCreditHours <> ostrCreditHours
        End Sub
        Private Sub Init()
            nCourseID = 0
            onProviderID = 0
            ostrCourseTitle = String.Empty
            'ostrCourseDates = 0
            'ostrLocation = String.Empty
            'onCourseTypeID = 0
            ostrProviderName = String.Empty
            'ostrCreditHours = String.Empty
            obolDeleted = False

            strCreatedBy = String.Empty
            dtCreatedOn = DateTime.Now.ToShortDateString
            strModifiedBy = String.Empty
            dtModifiedOn = System.DateTime.Now.ToShortDateString

            Me.Reset()
        End Sub

#End Region

#Region "Exposed Attributes"

        Public Property ID() As Integer
            Get
                Return nCourseID
            End Get

            Set(ByVal value As Integer)
                nCourseID = Integer.Parse(value)
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
        Public Property ProviderID() As Integer
            Get
                Return nProviderID
            End Get
            Set(ByVal Value As Integer)
                nProviderID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CourseTitle() As String
            Get
                Return strCourseTitle
            End Get
            Set(ByVal Value As String)
                strCourseTitle = Value
                Me.CheckDirty()
            End Set
        End Property
        'Public Property CourseDates() As String
        '    Get
        '        Return strCourseDates
        '    End Get

        '    Set(ByVal value As String)
        '        strCourseDates = value
        '        Me.CheckDirty()
        '    End Set
        'End Property
        'Public Property Location() As String
        '    Get
        '        Return strLocation
        '    End Get

        '    Set(ByVal value As String)
        '        strLocation = value
        '        Me.CheckDirty()
        '    End Set
        'End Property
        'Public Property CourseTypeID() As Integer
        '    Get
        '        Return nCourseTypeID
        '    End Get

        '    Set(ByVal value As Integer)
        '        nCourseTypeID = Integer.Parse(value)
        '        Me.CheckDirty()
        '    End Set
        'End Property
        Public Property ProviderName() As String
            Get
                Return strProviderName
            End Get

            Set(ByVal value As String)
                strProviderName = value
                Me.CheckDirty()
            End Set
        End Property
        'Public Property CreditHours() As String
        '    Get
        '        Return strCreditHours
        '    End Get

        '    Set(ByVal value As String)
        '        strCreditHours = value
        '        Me.CheckDirty()
        '    End Set
        'End Property
        
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get

            Set(ByVal value As Boolean)
                bolDeleted = value
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
