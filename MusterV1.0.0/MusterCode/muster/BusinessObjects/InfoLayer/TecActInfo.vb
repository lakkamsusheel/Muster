'-------------------------------------------------------------------------------
' MUSTER.Info.TecActInfo
'   Provides the container to persist MUSTER Technical Activity state
'
' Copyright (C) 2004, 2005 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC       05/31/05    Original class definition.
'
' Function          Description
'-------------------------------------------------------------------------------
'

Namespace MUSTER.Info
    Public Class TecActInfo

#Region "ENUMS"
        Public Enum ActivityCostModeEnum

            NotFundable = 0
            PerActivity = 1
            PerMonth = 2
            PerYear = 3
            PerQuarter = 4

        End Enum
#End Region

#Region "Private Member Variables"
        Private bolActive As Boolean
        Private bolDeleted As Boolean
        Private strName As String
        Private nActDays As Int16
        Private nWarnDays As Int16
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private nEntityID As Integer
        Private obolActive As Boolean
        Private obolDeleted As Boolean
        Private onActID As Int64
        Private ostrName As String
        Private onActDays As Int16
        Private onWarnDays As Int16
        Private nActID As Int64
        Private bolIsDirty As Boolean
        Private strCreatedBy As String
        Private strModifiedBy As String
        Private acmCostMode As ActivityCostModeEnum
        Private oacmCostMode As ActivityCostModeEnum

        Private mCost As Double
        Private omCost As Double

        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private colDocs As MUSTER.Info.TecDocCollection

#End Region
#Region "Public Events"
        Public Delegate Sub InfoChangedEventHandler()
        Public Event InfoChanged As InfoChangedEventHandler
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            colDocs = New MUSTER.Info.TecDocCollection
            dtDataAge = Now()
        End Sub
        Public Sub New(ByVal Id As Long, _
            ByVal sName As String, _
            ByVal nAction As Int16, _
            ByVal nWarn As Int16, _
            ByVal CreatedBy As String, _
            ByVal CreateDate As Date, _
            ByVal LastEditedBy As String, _
            ByVal LastEditDate As Date, _
            ByVal bActive As Boolean, _
            ByVal bDeleted As Boolean, Optional ByVal acmCostMode As ActivityCostModeEnum = ActivityCostModeEnum.NotFundable, _
                Optional ByVal mCostAmount As Double = 0)

            onActID = Id
            ostrName = sName
            onActDays = nAction
            onWarnDays = nWarn
            strCreatedBy = CreatedBy
            strModifiedBy = LastEditedBy
            dtModifiedOn = LastEditDate
            dtCreatedOn = CreateDate
            obolActive = bActive
            obolDeleted = bDeleted

            dtDataAge = Now()
            oacmCostMode = acmCostMode
            omCost = mCostAmount

            colDocs = New MUSTER.Info.TecDocCollection

            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"

        ' the Active/Inactive flag for the TEC_DOC
        Public Property CostMode() As ActivityCostModeEnum
            Get
                Return acmCostMode
            End Get
            Set(ByVal Value As ActivityCostModeEnum)
                acmCostMode = Value
            End Set
        End Property


        Public Property Cost() As Double
            Get
                Return mCost
            End Get
            Set(ByVal Value As Double)
                mCost = Value
            End Set
        End Property

        Public Property Active() As Boolean
            Get
                Return bolActive
            End Get
            Set(ByVal Value As Boolean)
                bolActive = Value
                CheckDirty()
            End Set
        End Property
        ' The maximum age the info object can attain before requiring a refresh
        Public Property AgeThreshold() As Date
            Get
                Return dtDataAge
            End Get
            Set(ByVal Value As Date)
                dtDataAge = Value
            End Set
        End Property

        ' The deleted flag for the TEC_ACT
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolDeleted = Value
                CheckDirty()
            End Set
        End Property
        ' The entity ID associated with a technical activity.
        Public ReadOnly Property EntityID() As Integer
            Get
                Return nEntityID
            End Get
        End Property
        ' The system ID for this Technical Activity
        Public Property ID() As Long
            Get
                Return nActID
            End Get
            Set(ByVal Value As Long)
                nActID = Value
            End Set
        End Property
        ' The Action Days threshold for this Technical Activity
        Public Property ActDays() As Long
            Get
                Return nActDays
            End Get
            Set(ByVal Value As Long)
                nActDays = Value
                CheckDirty()
            End Set
        End Property
        ' The Warn Days threshold for this Technical Activity
        Public Property WarnDays() As Long
            Get
                Return nWarnDays
            End Get
            Set(ByVal Value As Long)
                nWarnDays = Value
                CheckDirty()
            End Set
        End Property
        ' Returns a boolean indicating if the data has aged beyond its preset limit
        Protected ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get
        End Property
        ' Raised when any of the TechnicalEventInfo attributes are modified
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get
            Set(ByVal Value As Boolean)
                bolIsDirty = Value
            End Set
        End Property

        ' The ID of the user that created the row
        Public Property CreatedBy() As String
            Get
                Return strCreatedBy
            End Get
            Set(ByVal Value As String)
                strCreatedBy = Value
            End Set
        End Property
        ' The date on which the row was created
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
        ' The name of the technical activity to be displayed throughout the system
        Public Property Name() As String
            Get
                Return strName
            End Get
            Set(ByVal Value As String)
                strName = Value
                CheckDirty()
            End Set
        End Property
        Public Property DocumentsCollection() As MUSTER.Info.TecDocCollection
            Get
                Return colDocs
            End Get
            Set(ByVal Value As MUSTER.Info.TecDocCollection)
                colDocs = Value
            End Set
        End Property
#End Region
#Region "Exposed Methods"
        Public Sub Archive()
            obolActive = bolActive
            obolDeleted = bolDeleted
            onActID = nActID
            ostrName = strName
            onActDays = nActDays
            onWarnDays = nWarnDays
            ostrName = strName
            omCost = mCost
            oacmCostMode = acmCostMode
            bolIsDirty = False

        End Sub

        Public Sub Reset()
            bolActive = obolActive
            bolDeleted = obolDeleted
            nActID = onActID
            strName = ostrName
            nActDays = onActDays
            nWarnDays = onWarnDays
            strName = ostrName
            bolIsDirty = False
            mCost = omCost
            acmCostMode = oacmCostMode

        End Sub
#End Region
#Region "Private Methods"
        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty

            bolIsDirty = (obolActive <> bolActive) Or _
                        (obolDeleted <> bolDeleted) Or _
                        (onActID <> nActID) Or _
                        (ostrName <> strName) Or _
                        (onActDays <> nActDays) Or _
                        (onWarnDays <> nWarnDays) Or _
                        (ostrName <> strName) Or _
                        (omCost <> mCost) Or _
                        (oacmCostMode <> acmCostMode)


        End Sub


        Sub Init()

            bolActive = True
            bolDeleted = False
            nActID = -1
            strName = String.Empty
            nActDays = 0
            nWarnDays = 0
            dtDataAge = System.DateTime.Now
            nAgeThreshold = 5
            nEntityID = 0
            obolActive = True
            obolDeleted = False
            onActID = -1
            ostrName = String.Empty
            onActDays = 0
            onWarnDays = 0
            omCost = 0
            oacmCostMode = ActivityCostModeEnum.NotFundable
            mCost = 0
            acmCostMode = ActivityCostModeEnum.NotFundable


            bolIsDirty = False
            strCreatedBy = String.Empty
            strModifiedBy = String.Empty
            dtModifiedOn = DateTime.Now.ToShortDateString
            dtCreatedOn = DateTime.Now.ToShortDateString

        End Sub
#End Region
#Region "Protected Methods"
        Protected Overrides Sub Finalize()
        End Sub
#End Region
    End Class
End Namespace
