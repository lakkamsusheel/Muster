
Namespace MUSTER.Info
    ' -------------------------------------------------------------------------------
    '    MUSTER.Info.FeeBasisInfo
    '          Provides the container to persist MUSTER Fee Basis state
    ' 
    '    Copyright (C) 2004, 2005 CIBER, Inc.
    '    All rights reserved.
    ' 
    '    Release   Initials    Date        Description
    '       1.0        JVC       06/14/05    Original class definition.
    ' 
    '    Function          Description
    ' -------------------------------------------------------------------------------
    '
    Public Class FeeBasisInfo
#Region "Private member variables"

        Private bolDeleted As Boolean
        Private bolIsDirty As Boolean = False
        Private nID As Integer
        Private curBaseFee As Decimal
        Private curLateFee As Decimal
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private dtDataAge As DateTime
        Private dtEarlyGrace As Date
        Private dtLateGrace As Date
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private dtPeriodEnd As Date
        Private dtPeriodStart As Date
        Private nAgeThreshold As Int16 = 5
        Private nBaseType As Long
        Private nEntityID As Integer
        Private nFiscalYear As Integer
        Private nLatePeriod As Long
        Private nLateType As Long
        Private obolDeleted As Boolean
        Private ocurBaseFee As Decimal
        Private ocurLateFee As Decimal
        Private odtCompleted As Date
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private odtEarlyGrace As Date
        Private odtLateGrace As Date
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private odtPeriodEnd As Date
        Private odtPeriodStart As Date
        Private onBaseType As Long
        Private onFiscalYear As Integer
        Private onLatePeriod As Long
        Private onLateType As Long
        Private ostrCreatedBy As String = String.Empty
        Private ostrModifiedBy As String = String.Empty
        Private strCreatedBy As String = String.Empty
        Private strDescription As String
        Private ostrDescription As String
        Private strModifiedBy As String = String.Empty

        Private odtGenerateDate As Date
        Private dtGenerateDate As Date
        Private odtApprovedDate As Date
        Private dtApprovedDate As Date

        Private odtGenerateTime As Date
        Private dtGenerateTime As Date
        Private odtApprovedTime As Date
        Private dtApprovedTime As Date
        Private bolGenerated As Boolean
        Private obolGenerated As Boolean

#End Region
#Region "Constructors"
        Public Sub New()
            MyBase.new()
            dtDataAge = Now()
            Me.Init()
        End Sub
        Public Sub New(ByVal ID As Integer, _
                        ByVal CREATED_BY As String, _
                        ByVal CREATE_DATE As String, _
                        ByVal LAST_EDITED_BY As String, _
                        ByVal DATE_LAST_EDITED As Date, _
                        ByVal DELETED As Integer, _
                        ByVal FISCAL_YEAR As Long, _
                        ByVal PERIOD_START As Date, _
                        ByVal PERIOD_END As Date, _
                        ByVal EARLY_GRACE As Date, _
                        ByVal LATE_GRACE As Date, _
                        ByVal BASE_FEE As Decimal, _
                        ByVal BASE_UNIT As Long, _
                        ByVal LATE_FEE As Decimal, _
                        ByVal LATE_TYPE As Long, _
                        ByVal LATE_PERIOD As Long, _
                        ByVal GENERATION_DATE As Date, _
                        ByVal ApprovedDate As Date, _
                        ByVal GENERATION_TIME As Date, _
                        ByVal ApprovedTime As Date, _
                        ByVal bGENERATED As Boolean, _
                        ByVal DESCRIPTION As String)

            odtGenerateTime = GENERATION_TIME
            odtApprovedTime = ApprovedTime
            odtGenerateDate = GENERATION_DATE
            odtApprovedDate = ApprovedDate
            obolDeleted = DELETED
            ocurBaseFee = BASE_FEE
            ocurLateFee = LATE_FEE
            odtEarlyGrace = EARLY_GRACE
            odtLateGrace = LATE_GRACE
            odtPeriodEnd = PERIOD_END
            odtPeriodStart = PERIOD_START
            onBaseType = BASE_UNIT
            onFiscalYear = FISCAL_YEAR
            onLatePeriod = LATE_PERIOD
            onLateType = LATE_TYPE
            ostrDescription = DESCRIPTION
            obolGenerated = bGENERATED

            ostrCreatedBy = CREATED_BY
            odtCreatedOn = CREATE_DATE
            ostrModifiedBy = LAST_EDITED_BY
            odtModifiedOn = DATE_LAST_EDITED

            nID = ID
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Archive()
            obolDeleted = bolDeleted
            'bolIsDirty = False
            ocurBaseFee = curBaseFee
            ocurLateFee = curLateFee
            'dtDataAge = odtDataAge
            odtEarlyGrace = dtEarlyGrace
            odtLateGrace = dtLateGrace
            odtPeriodEnd = dtPeriodEnd
            odtPeriodStart = dtPeriodStart
            'nAgeThreshold = onAgeThreshold
            onBaseType = nBaseType
            'nEntityID = onEntityID
            onFiscalYear = nFiscalYear
            onLatePeriod = nLatePeriod
            onLateType = nLateType
            ostrDescription = strDescription
            odtGenerateDate = dtGenerateDate
            odtApprovedDate = dtApprovedDate
            odtGenerateTime = dtGenerateTime
            odtApprovedTime = dtApprovedTime
            obolGenerated = bolGenerated

            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn

        End Sub
        Public Sub Reset()
            bolDeleted = obolDeleted
            bolIsDirty = False
            curBaseFee = ocurBaseFee
            curLateFee = ocurLateFee
            'dtDataAge = odtDataAge
            dtEarlyGrace = odtEarlyGrace
            dtLateGrace = odtLateGrace
            dtPeriodEnd = odtPeriodEnd
            dtPeriodStart = odtPeriodStart
            'nAgeThreshold = onAgeThreshold
            nBaseType = onBaseType
            'nEntityID = onEntityID
            nFiscalYear = onFiscalYear
            nLatePeriod = onLatePeriod
            nLateType = onLateType
            strDescription = ostrDescription
            dtGenerateDate = odtGenerateDate
            dtApprovedDate = odtApprovedDate
            dtGenerateTime = odtGenerateTime
            dtApprovedTime = odtApprovedTime
            bolGenerated = obolGenerated
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn

        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            bolIsDirty = (bolDeleted <> obolDeleted) Or _
                            (curBaseFee <> ocurBaseFee) Or _
                            (curLateFee <> ocurLateFee) Or _
                            (dtCreatedOn <> odtCreatedOn) Or _
                            (dtEarlyGrace <> odtEarlyGrace) Or _
                            (dtLateGrace <> odtLateGrace) Or _
                            (dtModifiedOn <> odtModifiedOn) Or _
                            (dtPeriodEnd <> odtPeriodEnd) Or _
                            (dtPeriodStart <> odtPeriodStart) Or _
                            (nBaseType <> onBaseType) Or _
                            (nFiscalYear <> onFiscalYear) Or _
                            (nLatePeriod <> onLatePeriod) Or _
                            (nLateType <> onLateType) Or _
                            (strDescription <> ostrDescription) Or _
                            (dtGenerateDate <> odtGenerateDate) Or _
                            (dtGenerateTime <> odtGenerateTime) Or _
                            (dtApprovedDate <> odtApprovedDate) Or _
                            (obolGenerated <> bolGenerated) Or _
                            (dtApprovedTime <> odtApprovedTime)
        End Sub
        Private Sub Init()
            nID = 0
            obolDeleted = False
            bolIsDirty = False
            ocurBaseFee = 0
            ocurLateFee = 0
            odtCreatedOn = System.DateTime.Now
            'dtDataAge = odtDataAge
            odtEarlyGrace = System.DateTime.Now
            odtLateGrace = System.DateTime.Now
            odtPeriodEnd = System.DateTime.Now
            odtPeriodStart = System.DateTime.Now
            'nAgeThreshold = onAgeThreshold
            onBaseType = 0
            'nEntityID = onEntityID
            onFiscalYear = 0
            onLatePeriod = 0
            onLateType = 0
            bolGenerated = False
            ostrDescription = String.Empty
            odtModifiedOn = DateTime.Now.ToShortDateString
            odtCreatedOn = DateTime.Now.ToShortDateString
            ostrCreatedBy = String.Empty
            ostrModifiedBy = String.Empty
            obolGenerated = False

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
        Public Property AgeThreshold() As Integer
            Get
                Return nAgeThreshold
            End Get
            Set(ByVal Value As Integer)
                nAgeThreshold = Value
            End Set
        End Property

        ' The base fee for the billing period
        Public Property BaseFee() As Decimal
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{FF7FADBF-919A-4CBD-9843-3081D461C046}
                Return curBaseFee
                ' #End Region ' XDEOperation End Template Expansion{FF7FADBF-919A-4CBD-9843-3081D461C046}
            End Get
            Set(ByVal Value As Decimal)
                ' #Region "XDEOperation" ' Begin Template Expansion{C0D15056-8A07-4C87-B44E-D54CEF94E898}
                curBaseFee = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{C0D15056-8A07-4C87-B44E-D54CEF94E898}
            End Set
        End Property
        ' The billing unit for the base fee (from tblSYS_PROPERTY)
        Public Property BaseUnit() As Long
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{15B8CEBE-3CD1-43AE-9D58-5786A4F1F1A0}
                Return nBaseType
                ' #End Region ' XDEOperation End Template Expansion{15B8CEBE-3CD1-43AE-9D58-5786A4F1F1A0}
            End Get
            Set(ByVal Value As Long)
                ' #Region "XDEOperation" ' Begin Template Expansion{D1272688-AF77-43AE-9922-2DE86A9025BF}
                nBaseType = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{D1272688-AF77-43AE-9922-2DE86A9025BF}
            End Set
        End Property
        ' The deleted flag for the LUST Activity
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolDeleted = Value
                Me.CheckDirty()
            End Set
        End Property
        ' The description associated with the Fee Basis.
        Public Property Description() As String
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{C74BAFAB-0772-415F-BCAE-086A627BC4AC}
                Return strDescription
                ' #End Region ' XDEOperation End Template Expansion{C74BAFAB-0772-415F-BCAE-086A627BC4AC}
            End Get
            Set(ByVal Value As String)
                ' #Region "XDEOperation" ' Begin Template Expansion{CC493E99-194F-400F-91A8-99FE99EE3035}
                strDescription = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{CC493E99-194F-400F-91A8-99FE99EE3035}
            End Set
        End Property
        ' The early grace date for the billing period
        Public Property EarlyGrace() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{D96312D0-9E26-4B91-BAC9-139C4BDF0411}
                Return dtEarlyGrace
                ' #End Region ' XDEOperation End Template Expansion{D96312D0-9E26-4B91-BAC9-139C4BDF0411}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{4FE52D00-EB97-48A7-AF7B-76D4043874F7}
                dtEarlyGrace = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{4FE52D00-EB97-48A7-AF7B-76D4043874F7}
            End Set
        End Property

        ' The fiscal year for the billing period
        Public Property FiscalYear() As Integer
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{BCAA7BC2-0BB4-4519-93A3-3E4B9EA7D025}
                Return nFiscalYear
                ' #End Region ' XDEOperation End Template Expansion{BCAA7BC2-0BB4-4519-93A3-3E4B9EA7D025}
            End Get
            Set(ByVal Value As Integer)
                ' #Region "XDEOperation" ' Begin Template Expansion{A4D2F744-0722-4100-8864-82A126F6A672}
                nFiscalYear = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{A4D2F744-0722-4100-8864-82A126F6A672}
            End Set
        End Property
        Public ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get
            Set(ByVal Value As Boolean)
                bolIsDirty = Value
            End Set
        End Property
        ' The base late fee for the billing period
        Public Property LateFee() As Decimal
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{2931EDF2-45F5-43F5-A976-52D27405B250}
                Return curLateFee
                ' #End Region ' XDEOperation End Template Expansion{2931EDF2-45F5-43F5-A976-52D27405B250}
            End Get
            Set(ByVal Value As Decimal)
                ' #Region "XDEOperation" ' Begin Template Expansion{00343597-7F77-4897-A47E-0853BB66EA59}
                curLateFee = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{00343597-7F77-4897-A47E-0853BB66EA59}
            End Set
        End Property
        ' The late grace date for the billing period
        Public Property LateGrace() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{211079AE-1910-4331-AD23-6B375FBBED21}
                Return dtLateGrace
                ' #End Region ' XDEOperation End Template Expansion{211079AE-1910-4331-AD23-6B375FBBED21}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{B8E949F6-1546-4A1C-95AE-519CC267E88F}
                dtLateGrace = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{B8E949F6-1546-4A1C-95AE-519CC267E88F}
            End Set
        End Property
        ' ??? Always a period of time (from tblSYS_PROPERTY_MASTER)
        Public Property LatePeriod() As Long
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{C2B0A621-58E1-41C4-AB57-B0D6F0FBCF22}
                Return nLatePeriod
                ' #End Region ' XDEOperation End Template Expansion{C2B0A621-58E1-41C4-AB57-B0D6F0FBCF22}
            End Get
            Set(ByVal Value As Long)
                ' #Region "XDEOperation" ' Begin Template Expansion{522D93BB-D321-4FB7-B00C-BBA718634EEA}
                nLatePeriod = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{522D93BB-D321-4FB7-B00C-BBA718634EEA}
            End Set
        End Property
        ' The type of late fee calculation involving the LateFee for the billing period (from tblSYS_PROPERTY_MASTER)
        Public Property LateType() As Long
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{599013B2-E512-4E4D-B4D2-29D9734DED67}
                Return nLateType
                ' #End Region ' XDEOperation End Template Expansion{599013B2-E512-4E4D-B4D2-29D9734DED67}
            End Get
            Set(ByVal Value As Long)
                ' #Region "XDEOperation" ' Begin Template Expansion{F2EE745C-7387-4389-B2A4-DA894EAEDB30}
                nLateType = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{F2EE745C-7387-4389-B2A4-DA894EAEDB30}
            End Set
        End Property
        ' The last date for the billing period
        Public Property PeriodEnd() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{FFDAFCDE-46FA-4684-BDD8-2F72C266C3EA}
                Return dtPeriodEnd
                ' #End Region ' XDEOperation End Template Expansion{FFDAFCDE-46FA-4684-BDD8-2F72C266C3EA}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{DE1BC0A4-8723-4D56-9324-7F25316BB788}
                dtPeriodEnd = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{DE1BC0A4-8723-4D56-9324-7F25316BB788}
            End Set
        End Property
        ' The start date for the billing period
        Public Property PeriodStart() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{AAD5AF15-174C-4D19-9D10-E75C5F9554EF}
                Return dtPeriodStart
                ' #End Region ' XDEOperation End Template Expansion{AAD5AF15-174C-4D19-9D10-E75C5F9554EF}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{D9621ED1-17A3-4274-ABE1-1C7A7A45806F}
                dtPeriodStart = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{D9621ED1-17A3-4274-ABE1-1C7A7A45806F}
            End Set
        End Property
        Public Property Generated() As Boolean
            Get
                Return bolGenerated
            End Get
            Set(ByVal Value As Boolean)
                bolGenerated = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property GenerateDate() As Date
            Get
                Return dtGenerateDate
            End Get
            Set(ByVal Value As Date)
                dtGenerateDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property GenerateTime() As Date
            Get
                Return dtGenerateTime
            End Get
            Set(ByVal Value As Date)
                dtGenerateTime = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ApprovedTime() As Date
            Get
                Return dtApprovedTime
            End Get
            Set(ByVal Value As Date)
                dtApprovedTime = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ApprovedDate() As Date
            Get
                Return dtApprovedDate
            End Get
            Set(ByVal Value As Date)
                dtApprovedDate = Value
                Me.CheckDirty()
            End Set
        End Property

        ' The Entity ID associated with a LUST Activity (from tblSYS_ENTITY)
        Protected ReadOnly Property EntityID() As Integer
            Get
                Return nEntityID
            End Get
        End Property

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
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
        End Sub
#End Region

        Public Delegate Sub FeeBasisInfoChangedEventHandler()


        ' Fired when CheckDirty determines that an attribute of the activity has been modified
        Public Event FeeBasisInfoChanged As FeeBasisInfoChangedEventHandler

    End Class
End Namespace
