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
    Public Class LicenseeCourseInfo
#Region "Public Events"
        Public Event LicCourseInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nID As Int64
        Private nLicenseeID As Integer
        Private nProviderID As Integer
        Private nCourseTypeID As Integer
        Private dtCourseDate As DateTime
        Private bolDeleted As Boolean
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime = DateTime.Now.ToShortDateString
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime = DateTime.Now.ToShortDateString

        Private onID As Int64
        Private onCourseID As Integer
        Private onLicenseeID As Integer
        Private onProviderID As Integer
        Private onCourseTypeID As Integer
        Private odtCourseDate As DateTime
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
        ByVal ProviderID As Integer, _
        ByVal CourseTypeID As Integer, _
        ByVal CourseDate As DateTime, _
        ByVal Deleted As Boolean, _
        ByVal CreatedBy As String, _
        ByVal dateCreated As Date, _
        ByVal LAST_EDITED_BY As String, _
        ByVal DATE_LAST_EDITED As Date)
            onID = ID
            onLicenseeID = LicenseeID
            onProviderID = ProviderID
            onCourseTypeID = CourseTypeID
            odtCourseDate = CourseDate
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
            nProviderID = onProviderID
            nCourseTypeID = onCourseTypeID
            dtCourseDate = odtCourseDate
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
            RaiseEvent LicCourseInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onID = nID
            onLicenseeID = nLicenseeID
            onProviderID = nProviderID
            onCourseTypeID = nCourseTypeID
            odtCourseDate = dtCourseDate
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            '(nID <> onID) Or _
            '(nLicenseeID <> onLicenseeID) Or _
            Dim obolIsDirty As Boolean = bolIsDirty

            bolIsDirty = (nProviderID <> onProviderID) Or _
            (nCourseTypeID <> onCourseTypeID) Or _
            (dtCourseDate <> odtCourseDate) Or _
            (bolDeleted <> obolDeleted) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent LicCourseInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            nID = 0
            nLicenseeID = 0
            nProviderID = 0
            nCourseTypeID = 0
            dtCourseDate = System.DateTime.Now
            bolDeleted = False
            strCreatedBy = String.Empty
            dtCreatedOn = DateTime.Now.ToShortDateString
            strModifiedBy = String.Empty
            dtModifiedOn = System.DateTime.Now.ToShortDateString
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int64
            Get
                Return nID
            End Get
            Set(ByVal Value As Int64)
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
        Public Property ProviderID() As Integer
            Get
                Return nProviderID
            End Get
            Set(ByVal Value As Integer)
                nProviderID = Value
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
        Public Property CourseDate() As DateTime
            Get
                Return dtCourseDate
            End Get
            Set(ByVal Value As DateTime)
                dtCourseDate = Value
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
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace
