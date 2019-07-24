'-------------------------------------------------------------------------------
' MUSTER.Info.FinancialActivityInfo
'   Provides the container to persist MUSTER FinancialActivityInfo state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        AB        06/23/05    Original class definition.
'  2.0  Thomas Franey   03/06/09    Added template doc Fields
'
' Function          Description
'-------------------------------------------------------------------------------
'
'-------------------------------------------------------------------------------
Namespace MUSTER.Info
    Public Class FinancialActivityInfo

#Region "Private Member Variables"
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private bolIsDirty As Boolean

        Private nActivityID As Int64
        Private strActivityDesc As String
        Private ostrActivityDesc As String
        Private strActivityDescShort As String
        Private ostrActivityDescShort As String
        Private nTimeAndMaterials As Int64
        Private onTimeAndMaterials As Int64
        Private nCostPlus As Int64
        Private onCostPlus As Int64
        Private nFixedPrice As Int64
        Private onFixedPrice As Int64
        Private strDueDateStatement As String
        Private ostrDueDateStatement As String
        Private nReimbursementCondition As Int64
        Private onReimbursementCondition As Int64

        Private strTimeAndMaterialsDesc As String
        Private strCostPlusDesc As String
        Private strFixedPriceDesc As String
        Private strReimburseConditionDesc As String

        Private strCoverTemplateDoc As String
        Private strNoticeTemplateDoc As String

        Private oStrCoverTemplateDoc As String
        Private oStrNoticeTemplateDoc As String




        Private bolActive As Boolean
        Private obolActive As Boolean
        Private bolDeleted As Boolean
        Private obolDeleted As Boolean

        Private strCreatedBy As String = String.Empty
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString

        Private ostrCreatedBy As String = String.Empty
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString

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
        Public Sub New(ByVal ActivityID As Int64, _
            ByVal ActivityDesc As String, _
            ByVal ActivityDescShort As String, _
            ByVal TimeAndMaterials As Int64, _
            ByVal CostPlus As Int64, _
            ByVal FixedPrice As Int64, _
            ByVal DueDateStatement As String, _
            ByVal ReimbursementCondition As Int64, _
            ByVal ReimbursementConditionDesc As String, _
            ByVal CreatedBy As String, _
            ByVal CreateDate As Date, _
            ByVal LastEditedBy As String, _
            ByVal LastEditDate As Date, _
            ByVal bActive As Boolean, _
            ByVal bDeleted As Boolean, _
            ByVal TimeAndMaterialsDesc As String, _
            ByVal CostPlusDesc As String, _
            ByVal FixedPriceDesc As String, _
            Optional ByVal CoverTemplateDoc As String = "", _
            Optional ByVal NoticeTemplateDoc As String = "", _
            Optional ByVal tecActivityType As Integer = 0)

            nActivityID = ActivityID
            ostrActivityDesc = ActivityDesc
            ostrActivityDescShort = ActivityDescShort
            onTimeAndMaterials = TimeAndMaterials
            onCostPlus = CostPlus
            onFixedPrice = FixedPrice
            ostrDueDateStatement = DueDateStatement
            onReimbursementCondition = ReimbursementCondition
            strReimburseConditionDesc = ReimbursementConditionDesc
            strTimeAndMaterialsDesc = TimeAndMaterialsDesc
            strCostPlusDesc = CostPlusDesc
            strFixedPriceDesc = FixedPriceDesc

            strCoverTemplateDoc = CoverTemplateDoc
            strNoticeTemplateDoc = NoticeTemplateDoc

            oStrCoverTemplateDoc = CoverTemplateDoc
            oStrNoticeTemplateDoc = NoticeTemplateDoc

            ostrCreatedBy = CreatedBy
            ostrModifiedBy = LastEditedBy
            odtModifiedOn = LastEditDate
            odtCreatedOn = CreateDate
            obolActive = bActive
            obolDeleted = bDeleted


            dtDataAge = Now()

            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"

        ' The system ID for this Technical Activity
        Public Property ActivityID() As Long
            Get
                Return nActivityID
            End Get
            Set(ByVal Value As Long)
                nActivityID = Value
            End Set
        End Property

        Public Property ActivityDesc() As String
            Get
                Return strActivityDesc
            End Get
            Set(ByVal Value As String)
                strActivityDesc = Value
                CheckDirty()
            End Set
        End Property
        Public Property ActivityDescShort() As String
            Get
                Return strActivityDescShort
            End Get
            Set(ByVal Value As String)
                strActivityDescShort = Value
                CheckDirty()
            End Set
        End Property

        Public Property NoticeTemplateDoc() As String
            Get
                Return strNoticeTemplateDoc
            End Get
            Set(ByVal Value As String)
                strNoticeTemplateDoc = Value
            End Set
        End Property

        Public Property CoverTemplateDoc() As String
            Get
                Return strCoverTemplateDoc
            End Get
            Set(ByVal Value As String)
                strCoverTemplateDoc = Value
            End Set
        End Property


        Public Property TimeAndMaterials() As Integer
            Get
                Return nTimeAndMaterials
            End Get
            Set(ByVal Value As Integer)
                nTimeAndMaterials = Value
                CheckDirty()
            End Set
        End Property
        Public Property CostPlus() As Integer
            Get
                Return nCostPlus
            End Get
            Set(ByVal Value As Integer)
                nCostPlus = Value
                CheckDirty()
            End Set
        End Property
        Public Property FixedPrice() As Integer
            Get
                Return nFixedPrice
            End Get
            Set(ByVal Value As Integer)
                nFixedPrice = Value
                CheckDirty()
            End Set
        End Property
        Public Property DueDateStatement() As String
            Get
                Return strDueDateStatement
            End Get
            Set(ByVal Value As String)
                strDueDateStatement = Value
                CheckDirty()
            End Set
        End Property
        Public Property ReimbursementCondition() As Integer
            Get
                Return nReimbursementCondition
            End Get
            Set(ByVal Value As Integer)
                nReimbursementCondition = Value
                CheckDirty()
            End Set
        End Property



        Public ReadOnly Property ReimbursementConditionDesc() As String
            Get
                Return strReimburseConditionDesc
            End Get
        End Property
        Public Property TimeAndMaterialsDesc() As String
            Get
                Return strTimeAndMaterialsDesc
            End Get
            Set(ByVal Value As String)
                strTimeAndMaterialsDesc = Value
            End Set
        End Property

        Public Property FixedPriceDesc() As String
            Get
                Return strFixedPriceDesc
            End Get
            Set(ByVal Value As String)
                strFixedPriceDesc = Value
            End Set
        End Property

        Public Property CostPlusDesc() As String
            Get
                Return strCostPlusDesc
            End Get
            Set(ByVal Value As String)
                strCostPlusDesc = Value
            End Set
        End Property


        ' the Active/Inactive flag for the TEC_DOC
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
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolDeleted = Value
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
#Region "Exposed Methods"
        Public Sub Archive()
            obolActive = bolActive
            obolDeleted = bolDeleted
            ostrActivityDesc = strActivityDesc
            ostrActivityDescShort = strActivityDescShort
            onTimeAndMaterials = nTimeAndMaterials
            onCostPlus = nCostPlus
            onFixedPrice = nFixedPrice
            ostrDueDateStatement = strDueDateStatement
            onReimbursementCondition = nReimbursementCondition
            oStrCoverTemplateDoc = strCoverTemplateDoc
            oStrNoticeTemplateDoc = strNoticeTemplateDoc

            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn

            bolIsDirty = False

        End Sub
        Public Sub Reset()
            bolActive = obolActive
            bolDeleted = obolDeleted
            strActivityDesc = ostrActivityDesc
            strActivityDescShort = ostrActivityDescShort
            nTimeAndMaterials = onTimeAndMaterials
            nCostPlus = onCostPlus
            nFixedPrice = onFixedPrice
            strDueDateStatement = ostrDueDateStatement
            nReimbursementCondition = onReimbursementCondition

            strCoverTemplateDoc = oStrCoverTemplateDoc
            strNoticeTemplateDoc = oStrNoticeTemplateDoc

            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn


            bolIsDirty = False
        End Sub
#End Region
#Region "Private Methods"
        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty

            bolIsDirty = (obolActive <> bolActive) Or _
                        (obolDeleted <> bolDeleted) Or _
                        (ostrActivityDesc <> strActivityDesc) Or _
                        (ostrActivityDescShort <> strActivityDescShort) Or _
                        (onTimeAndMaterials <> nTimeAndMaterials) Or _
                        (onCostPlus <> nCostPlus) Or _
                        (onFixedPrice <> nFixedPrice) Or _
                        (ostrDueDateStatement <> strDueDateStatement) Or _
                        (onReimbursementCondition <> nReimbursementCondition) Or _
                        (oStrCoverTemplateDoc <> strCoverTemplateDoc) Or _
                        (oStrNoticeTemplateDoc <> strNoticeTemplateDoc)


        End Sub


        Sub Init()

            bolActive = True
            bolDeleted = False
            nActivityID = 0
            ostrActivityDesc = String.Empty
            ostrActivityDescShort = String.Empty
            onTimeAndMaterials = 0
            onCostPlus = 0
            onFixedPrice = 0
            ostrDueDateStatement = String.Empty
            onReimbursementCondition = 0

            bolIsDirty = False
            strCreatedBy = String.Empty
            strModifiedBy = String.Empty
            dtModifiedOn = DateTime.Now.ToShortDateString
            dtCreatedOn = DateTime.Now.ToShortDateString

            strCoverTemplateDoc = String.Empty
            strNoticeTemplateDoc = String.Empty


        End Sub
#End Region
#Region "Protected Methods"
        Protected Overrides Sub Finalize()
        End Sub
#End Region


    End Class
End Namespace
