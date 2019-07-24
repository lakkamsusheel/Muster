' -------------------------------------------------------------------------------
' MUSTER.Info.FinancialCommitmentInfo
' Provides the container to persist MUSTER FinancialCommitmentInfo state
' 
' Copyright (C) 2004, 2005 CIBER, Inc.
' All rights reserved.
' 
' Release   Initials    Date        Description
' 1.0        AB       06/24/05    Original class definition.
' 
' Function          Description
' ---

Namespace MUSTER.Info

    Public Class FinancialCommitAdjustmentInfo
        ' Delegate for event to indicate to parent that the info object has been modified in some manner
        Public Delegate Sub FinancialCommitAdjustmentChangedEventHandler()

#Region "Private Member Variables"

        Private nCommitAdjustmentID As Int64

        Private nCommitmentID As Int64
        Private dtAdjustDate As Date
        Private nAdjustType As Int64
        Private sAdjustMoney As Double
        Private bolDirectorApprovalReq As Boolean
        Private bolFinancialApprovalReq As Boolean
        Private bolApproved As Boolean
        Private strComments As String

        Private onCommitmentID As Int64
        Private odtAdjustDate As Date
        Private onAdjustType As Int64
        Private osAdjustMoney As Double
        Private obolDirectorApprovalReq As Boolean
        Private obolFinancialApprovalReq As Boolean
        Private obolApproved As Boolean
        Private ostrComments As String

        Private bolIsDirty As Boolean
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private nEntityID As Integer
        Private bolDeleted As Boolean
        Private obolDeleted As Boolean

        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString

        Private ostrCreatedBy As String = String.Empty
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString

        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region

#Region "Public Events"
        Public Event FinancialCommitAdjustmentInfoChanged As FinancialCommitAdjustmentChangedEventHandler
#End Region

#Region "Constructors"
        Public Sub New()
            MyBase.New()
            Me.Init()
            dtDataAge = Now()
        End Sub
        Public Sub New(ByVal Id As Long, _
                        ByVal Commitment_ID As Int64, _
                        ByVal Adjust_Date As Date, _
                        ByVal Adjust_Type As Int64, _
                        ByVal Adjust_Amount As Double, _
                        ByVal Director_App_Req As Boolean, _
                        ByVal Fin_App_Req As Boolean, _
                        ByVal Approved As Boolean, _
                        ByVal Comments As String, _
                        ByVal CreatedBy As String, _
                        ByVal CreateDate As Date, _
                        ByVal LastEditedBy As String, _
                        ByVal LastEditDate As Date, _
                        ByVal bDeleted As Boolean)


            nCommitAdjustmentID = Id
            onCommitmentID = Commitment_ID
            odtAdjustDate = Adjust_Date
            onAdjustType = Adjust_Type
            osAdjustMoney = Adjust_Amount
            obolDirectorApprovalReq = Director_App_Req
            obolFinancialApprovalReq = Fin_App_Req
            obolApproved = Approved
            ostrComments = Comments

            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreateDate
            ostrModifiedBy = LastEditedBy
            odtModifiedOn = LastEditDate


            obolDeleted = bDeleted

            dtDataAge = Now()
            Me.Reset()

        End Sub
#End Region

#Region "Exposed Methods"
        ' Add other attributes as necessitated by design
        Public Sub Archive()

            onCommitmentID = nCommitmentID
            odtAdjustDate = dtAdjustDate
            onAdjustType = nAdjustType
            osAdjustMoney = sAdjustMoney
            obolDirectorApprovalReq = bolDirectorApprovalReq
            obolFinancialApprovalReq = bolFinancialApprovalReq
            obolApproved = bolApproved
            ostrComments = strComments

            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn

            obolDeleted = bolDeleted
            IsDirty = False
        End Sub

        Public Sub Reset()


            nCommitmentID = onCommitmentID
            dtAdjustDate = odtAdjustDate
            nAdjustType = onAdjustType
            sAdjustMoney = osAdjustMoney
            bolDirectorApprovalReq = obolDirectorApprovalReq
            bolFinancialApprovalReq = obolFinancialApprovalReq
            bolApproved = obolApproved
            strComments = ostrComments
            bolDeleted = obolDeleted

            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn

            IsDirty = False
        End Sub

#End Region

#Region "Private Methods"

        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty

            bolIsDirty = (nCommitmentID <> onCommitmentID) Or _
                        (dtAdjustDate <> odtAdjustDate) Or _
                        (nAdjustType <> onAdjustType) Or _
                        (sAdjustMoney <> osAdjustMoney) Or _
                        (bolDirectorApprovalReq <> obolDirectorApprovalReq) Or _
                        (bolFinancialApprovalReq <> obolFinancialApprovalReq) Or _
                        (bolApproved <> obolApproved) Or _
                        (strComments <> ostrComments) Or _
                        (obolDeleted <> bolDeleted)

        End Sub

        Public Sub Init()
            Dim tmpDate As Date

            nCommitAdjustmentID = 0

            onCommitmentID = 0
            odtAdjustDate = "01/01/0001"
            onAdjustType = 0
            osAdjustMoney = 0
            obolDirectorApprovalReq = False
            obolFinancialApprovalReq = False
            obolApproved = False
            ostrComments = String.Empty

            strCreatedBy = String.Empty
            dtCreatedOn = tmpDate
            strModifiedBy = String.Empty
            dtModifiedOn = tmpDate
            obolDeleted = False
        End Sub
#End Region

#Region "Protected Methods"
        Protected Overrides Sub Finalize()
        End Sub
#End Region

#Region "Exposed Attributes"
        ' the uniqueIdetifier for the _ProtoInfo
        Public Property CommitAdjustmentID() As Int64
            Get
                Return nCommitAdjustmentID
            End Get
            Set(ByVal Value As Int64)
                nCommitAdjustmentID = Value
            End Set
        End Property
        Public Property CommitmentID() As Int64
            Get
                Return nCommitmentID
            End Get
            Set(ByVal Value As Int64)
                nCommitmentID = Value
            End Set
        End Property
        Public Property AdjustDate() As Date
            Get
                Return dtAdjustDate
            End Get
            Set(ByVal Value As Date)
                dtAdjustDate = Value
            End Set
        End Property
        Public Property AdjustType() As Int64
            Get
                Return nAdjustType
            End Get
            Set(ByVal Value As Int64)
                nAdjustType = Value
            End Set
        End Property
        Public Property AdjustMoney() As Double
            Get
                Return sAdjustMoney
            End Get
            Set(ByVal Value As Double)
                sAdjustMoney = Value
            End Set
        End Property

        Public Property DirectorApprovalReq() As Boolean
            Get
                Return bolDirectorApprovalReq
            End Get
            Set(ByVal Value As Boolean)
                bolDirectorApprovalReq = Value
            End Set
        End Property
        Public Property FinancialApprovalReq() As Boolean
            Get
                Return bolFinancialApprovalReq
            End Get
            Set(ByVal Value As Boolean)
                bolFinancialApprovalReq = Value
            End Set
        End Property
        Public Property Approved() As Boolean
            Get
                Return bolApproved
            End Get
            Set(ByVal Value As Boolean)
                bolApproved = Value
            End Set
        End Property
        Public Property Comments() As String
            Get
                Return strComments
            End Get
            Set(ByVal Value As String)
                strComments = Value
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
        ' 
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

        ' The entity ID associated.
        Public ReadOnly Property EntityID() As Integer
            Get
            End Get
        End Property


        ' Raised when any of the _ProtoInfo attributes are modified
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get
            Set(ByVal Value As Boolean)
                bolIsDirty = False
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
        ' Returns a boolean indicating if the data has aged beyond its preset limit
        Protected ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get
        End Property
        Public ReadOnly Property ApprovedOriginal() As Boolean
            Get
                Dim bolTempApproved As Boolean = False
                bolTempApproved = bolApproved
                If Not obolApproved = bolApproved Then
                    bolApproved = obolApproved
                    CheckDirty()
                    If bolIsDirty Then
                        Return False
                    Else
                        bolApproved = bolTempApproved
                        Return True
                    End If
                Else
                    Return False
                End If
            End Get
        End Property
#End Region

    End Class

End Namespace
