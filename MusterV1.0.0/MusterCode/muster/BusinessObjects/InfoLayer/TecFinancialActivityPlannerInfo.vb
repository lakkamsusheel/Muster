'-------------------------------------------------------------------------------
' MUSTER.Info.TecFinancialActivityPlannerInfo
'   Provides the container to persist MUSTER TecFinancialActivityPlannerInfo state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0  Thomas Franey   05/06/09    Original class definition.
'
' Function          Description
'-------------------------------------------------------------------------------
'
'-------------------------------------------------------------------------------
Namespace MUSTER.Info
    Public Class TecFinancialActivityPlannerInfo

#Region "Private Member Variables"
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private bolIsDirty As Boolean = False

        Private nActivityTypeID As Int64
        Private mCost As Double
        Private nEventID As Int64
        Private nDuration As Int64

        Private onActivityTypeID As Int64
        Private omCost As Double
        Private onEventID As Int64
        Private onDuration As Int64


#End Region
#Region "Public Events"
        Public Delegate Sub InfoChangedEventHandler()
        Public Event InfoChanged As InfoChangedEventHandler
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            dtDataAge = Now()
        End Sub
        Public Sub New(ByVal EventID As Int64, _
            ByVal ActivityTypeID As Int64, _
            Optional ByVal Duration As Int64 = 0, _
            Optional ByVal Cost As Double = 0)

            onEventID = EventID
            onActivityTypeID = ActivityTypeID
            omCost = Cost
            onDuration = Duration

            dtDataAge = Now()

            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"

        ' The system ID for this Technical Activity
        Public Property ActivityTypeID() As Int64
            Get
                Return nActivityTypeID
            End Get
            Set(ByVal Value As Int64)
                nActivityTypeID = Value
            End Set
        End Property


        Public Property EventID() As Int64
            Get
                Return nEventID
            End Get
            Set(ByVal Value As Int64)
                nEventID = Value
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

        Public Property Duration() As Int64
            Get
                Return nDuration
            End Get
            Set(ByVal Value As Int64)
                nDuration = Value
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

        ' Returns a boolean indicating if the data has aged beyond its preset limit
        Protected ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get
        End Property
        ' Raised when any of the TechnicalEventInfo attributes are modified
        Public Property IsDirty() As Boolean
            Get

                CheckDirty()
                Return bolIsDirty
            End Get
            Set(ByVal Value As Boolean)
                bolIsDirty = Value
            End Set
        End Property


#End Region
#Region "Exposed Methods"
        Public Sub Archive()

            omCost = mCost
            onDuration = nDuration
            onEventID = nEventID
            onActivityTypeID = nActivityTypeID

            bolIsDirty = False

        End Sub
        Public Sub Reset()

            mCost = omCost
            nDuration = onDuration
            nEventID = onEventID
            nActivityTypeID = onActivityTypeID
            bolIsDirty = False

        End Sub
#End Region

#Region "Private Methods"
        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty

            bolIsDirty = (omCost <> mCost) Or _
            (onDuration <> nDuration) Or _
            (onEventID <> nEventID) Or _
            (onActivityTypeID <> nActivityTypeID)

        End Sub


        Sub Init()

            onEventID = 0
            onActivityTypeID = 0
            omCost = 0
            onDuration = 0

            nEventID = 0
            nActivityTypeID = 0
            mCost = 0
            nDuration = 0

        End Sub
#End Region
#Region "Protected Methods"
        Protected Overrides Sub Finalize()
        End Sub
#End Region


    End Class
End Namespace
