'-------------------------------------------------------------------------------
' MUSTER.Info.CompanyLicenseeInfo
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      MKK/RAF     05/26/2005  Original class definition
'  1.1        MR        06/04/2005  Added Functions and New Attributes
'
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class LicenseeCourseTestInfo
#Region "Public Events"
        Public Event LicenseeCourseTestInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"

        Private nID As Int64
        Private nLicenseeID As Integer
        Private nCourseTypeID As Integer
        Private dtTestDate As DateTime
        Private strStartTime As String
        Private nTestScore As Integer
        Private bolDeleted As Boolean
        Private strCreatedBy As String
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String
        Private dtModifiedOn As DateTime

        Private onID As Int64
        Private onLicenseeID As Integer
        Private onCourseTypeID As Integer
        Private odtTestDate As DateTime
        Private ostrStartTime As String
        Private onTestScore As Integer
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String
        Private odtModifiedOn As DateTime

        Private bolShowDeleted As Boolean = False
        Private nAgeThreshold As Int16 = 5
        Private dtDataAge As DateTime
        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub

        Sub New(ByVal ID As Integer, _
            ByVal LicenseeID As Integer, _
            ByVal CourseTypeID As Integer, _
            ByVal TestDate As DateTime, _
            ByVal StartTime As String, _
            ByVal TestScore As String, _
            ByVal Deleted As Boolean, _
            ByVal CreatedBy As String, _
            ByVal dateCreated As Date, _
            ByVal LAST_EDITED_BY As String, _
            ByVal DATE_LAST_EDITED As Date)

            onID = ID
            onLicenseeID = LicenseeID
            onCourseTypeID = CourseTypeID
            odtTestDate = TestDate
            ostrStartTime = StartTime
            onTestScore = TestScore
            obolDeleted = Deleted
            ostrCreatedBy = CreatedBy
            odtCreatedOn = dateCreated
            ostrModifiedBy = LAST_EDITED_BY
            odtModifiedOn = DATE_LAST_EDITED
            Me.Reset()
            dtDataAge = Now()
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nID >= 0 Then
                nID = onID
            End If
            nLicenseeID = onLicenseeID
            nCourseTypeID = onCourseTypeID
            dtTestDate = odtTestDate
            strStartTime = ostrStartTime
            nTestScore = onTestScore
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
            RaiseEvent LicenseeCourseTestInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onID = nID
            onLicenseeID = nLicenseeID
            onCourseTypeID = nCourseTypeID
            odtTestDate = dtTestDate
            ostrStartTime = strStartTime
            onTestScore = nTestScore
            obolDeleted = bolDeleted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            'nID <> onID) Or _
            '(nLicenseeID <> onLicenseeID) Or _
            Dim obolIsDirty As Boolean = bolIsDirty

            bolIsDirty = (nCourseTypeID <> onCourseTypeID) Or _
            (dtTestDate <> odtTestDate) Or _
            (strStartTime <> ostrStartTime) Or _
            (nTestScore <> onTestScore) Or _
            (bolDeleted <> obolDeleted) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent LicenseeCourseTestInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            nID = 0
            nLicenseeID = 0
            nCourseTypeID = 0
            dtTestDate = System.DateTime.Now
            strStartTime = String.Empty
            nTestScore = 0
            bolDeleted = False
            strCreatedBy = String.Empty
            dtCreatedOn = System.DateTime.Now
            strModifiedBy = String.Empty
            dtModifiedOn = System.DateTime.Now
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return nID
            End Get
            Set(ByVal Value As Integer)
                nID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LicenseeID() As Integer
            Get
                Return nLicenseeID
            End Get
            Set(ByVal Value As Integer)
                nLicenseeID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CourseTypeID() As Integer
            Get
                Return nCourseTypeID
            End Get
            Set(ByVal Value As Integer)
                nCourseTypeID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TestDate() As DateTime
            Get
                Return dtTestDate
            End Get
            Set(ByVal Value As DateTime)
                dtTestDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property StartTime() As String
            Get
                Return strStartTime
            End Get
            Set(ByVal Value As String)
                strStartTime = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TestScore() As Integer
            Get
                Return nTestScore
            End Get
            Set(ByVal Value As Integer)
                nTestScore = Value
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
        Public Property CreatedBy() As String
            Get
                Return strCreatedBy
            End Get
            Set(ByVal Value As String)
                strCreatedBy = Value
            End Set
        End Property
        Public Property CreatedOn() As DateTime
            Get
                Return dtCreatedOn
            End Get
            Set(ByVal Value As DateTime)
                dtCreatedOn = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
            End Set
        End Property
        Public Property ModifiedOn() As DateTime
            Get
                Return dtModifiedOn
            End Get
            Set(ByVal Value As DateTime)
                dtModifiedOn = Value
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
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace
