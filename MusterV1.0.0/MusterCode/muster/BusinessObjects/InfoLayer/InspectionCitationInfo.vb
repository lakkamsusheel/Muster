'-------------------------------------------------------------------------------
' MUSTER.Info.InspectionCitationInfo
'   Provides the container to persist MUSTER InspectionCitation state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MNR        06/15/05    Original class definition
'
' Function          Description
' New()             Instantiates an empty InspectionCitationInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated InspectionCitationInfo object
' New(dr)           Instantiates a populated InspectionCitationInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
' NOTE: This file to be used as InspectionCitation to build other objects.
'       Replace keyword "InspectionCitation" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class InspectionCitationInfo
#Region "Public Events"
        Public Event evtInspectionCitationInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nInsCITID As Int64
        Private nInspectionID As Int64
        Private nQuestionID As Int64
        Private nFacilityID As Int64
        Private nFCEID As Int64
        Private nCitationID As Int64
        Private strCCAT As String
        Private bolRescinded As Boolean
        Private dtCitationDueDate As Date
        Private dtCitationReceivedDate As Date
        Private dtNFADate As Date
        Private bolDeleted As Boolean
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime
        Private nOCEID As Int64

        Private onInsCITID As Int64
        Private onInspectionID As Int64
        Private onQuestionID As Int64
        Private onFacilityID As Int64
        Private onFCEID As Int64
        Private onCitationID As Int64
        Private ostrCCAT As String
        Private obolRescinded As Boolean
        Private odtCitationDueDate As Date
        Private odtCitationReceivedDate As Date
        Private odtNFADate As Date
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As DateTime
        Private onOCEID As Int64

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub
        Sub New(ByVal id As Int64, _
        ByVal inspectionID As Int64, _
        ByVal questionID As Int64, _
        ByVal facilityID As Int64, _
        ByVal fceid As Int64, _
        ByVal oceid As Int64, _
        ByVal citationID As Int64, _
        ByVal ccat As String, _
        ByVal rescinded As Boolean, _
        ByVal citationDueDate As Date, _
        ByVal citationReceivedDate As Date, _
        ByVal nfadate As Date, _
        ByVal deleted As Boolean, _
        ByVal createdBy As String, _
        ByVal createdOn As Date, _
        ByVal modifiedBy As String, _
        ByVal modifiedOn As Date)
            onInsCITID = id
            onInspectionID = inspectionID
            onQuestionID = questionID
            onFacilityID = facilityID
            onFCEID = fceid
            onCitationID = citationID
            ostrCCAT = ccat
            obolRescinded = rescinded
            odtCitationDueDate = citationDueDate
            odtCitationReceivedDate = citationReceivedDate
            odtNFADate = nfadate
            obolDeleted = deleted
            ostrCreatedBy = createdBy
            odtCreatedOn = createdOn
            ostrModifiedBy = modifiedBy
            odtModifiedOn = modifiedOn
            Me.Reset()
        End Sub
        Sub New(ByVal drInspectionCitation As DataRow)
            Try
                onInsCITID = drInspectionCitation.Item("INS_CIT_ID")
                onInspectionID = drInspectionCitation.Item("INSPECTION_ID")
                onQuestionID = IIf(drInspectionCitation.Item("QUESTION_ID") Is DBNull.Value, 0, drInspectionCitation.Item("QUESTION_ID"))
                onFacilityID = IIf(drInspectionCitation.Item("FACILITY_ID") Is DBNull.Value, 0, drInspectionCitation.Item("FACILITY_ID"))
                onFCEID = IIf(drInspectionCitation.Item("FCE_ID") Is DBNull.Value, 0, drInspectionCitation.Item("FCE_ID"))
                onOCEID = IIf(drInspectionCitation.Item("OCE_ID") Is DBNull.Value, 0, drInspectionCitation.Item("OCE_ID"))
                onCitationID = IIf(drInspectionCitation.Item("CITATION_ID") Is DBNull.Value, 0, drInspectionCitation.Item("CITATION_ID"))
                ostrCCAT = IIf(drInspectionCitation.Item("CCAT") Is DBNull.Value, String.Empty, drInspectionCitation.Item("CCAT"))
                obolRescinded = IIf(drInspectionCitation.Item("RESCINDED") Is DBNull.Value, False, drInspectionCitation.Item("RESCINDED"))
                odtCitationDueDate = IIf(drInspectionCitation.Item("CITATION_DUE_DATE") Is DBNull.Value, CDate("01/01/0001"), drInspectionCitation.Item("CITATION_DUE_DATE"))
                odtCitationReceivedDate = IIf(drInspectionCitation.Item("CITATION_RECEIVED_DATE") Is DBNull.Value, CDate("01/01/0001"), drInspectionCitation.Item("CITATION_RECEIVED_DATE"))
                odtNFADate = IIf(drInspectionCitation.Item("NFA_DATE") Is DBNull.Value, CDate("01/01/0001"), drInspectionCitation.Item("NFA_DATE"))
                obolDeleted = IIf(drInspectionCitation.Item("DELETED") Is DBNull.Value, False, drInspectionCitation.Item("DELETED"))
                ostrCreatedBy = IIf(drInspectionCitation.Item("CREATED_BY") Is DBNull.Value, String.Empty, drInspectionCitation.Item("CREATED_BY"))
                odtCreatedOn = IIf(drInspectionCitation.Item("DATE_CREATED") Is DBNull.Value, CDate("01/01/0001"), drInspectionCitation.Item("DATE_CREATED"))
                ostrModifiedBy = IIf(drInspectionCitation.Item("LAST_EDITED_BY") Is DBNull.Value, String.Empty, drInspectionCitation.Item("LAST_EDITED_BY"))
                odtModifiedOn = IIf(drInspectionCitation.Item("DATE_LAST_EDITED") Is DBNull.Value, CDate("01/01/0001"), drInspectionCitation.Item("DATE_LAST_EDITED"))
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nInsCITID >= 0 Then
                nInsCITID = onInsCITID
            End If
            nInspectionID = onInspectionID
            nQuestionID = onQuestionID
            nFacilityID = onFacilityID
            nFCEID = onFCEID
            nOCEID = onOCEID
            nCitationID = onCitationID
            strCCAT = ostrCCAT
            bolRescinded = obolRescinded
            dtCitationDueDate = odtCitationDueDate
            dtCitationReceivedDate = odtCitationReceivedDate
            dtNFADate = odtNFADate
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
            RaiseEvent evtInspectionCitationInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onInsCITID = nInsCITID
            onInspectionID = nInspectionID
            onQuestionID = nQuestionID
            onFacilityID = nFacilityID
            onFCEID = nFCEID
            onOCEID = nOCEID
            onCitationID = nCitationID
            ostrCCAT = strCCAT
            obolRescinded = bolRescinded
            odtCitationDueDate = dtCitationDueDate
            odtCitationReceivedDate = dtCitationReceivedDate
            odtNFADate = dtNFADate
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
            Dim obolIsDirty As Boolean = bolIsDirty

            bolIsDirty = (nInspectionID <> onInspectionID) Or _
            (nQuestionID <> onQuestionID) Or _
            (nFacilityID <> onFacilityID) Or _
            (nFCEID <> onFCEID) Or _
            (nOCEID <> onOCEID) Or _
            (nCitationID <> onCitationID) Or _
            (strCCAT <> ostrCCAT) Or _
            (bolRescinded <> obolRescinded) Or _
            (dtCitationDueDate <> odtCitationDueDate) Or _
            (dtCitationReceivedDate <> odtCitationReceivedDate) Or _
            (dtNFADate <> odtNFADate) Or _
            (bolDeleted <> obolDeleted) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent evtInspectionCitationInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onInsCITID = 0
            onInspectionID = 0
            onQuestionID = 0
            onFacilityID = 0
            onFCEID = 0
            onCitationID = 0
            ostrCCAT = String.Empty
            obolRescinded = False
            odtCitationDueDate = CDate("01/01/0001")
            odtCitationReceivedDate = CDate("01/01/0001")
            odtNFADate = CDate("01/01/0001")
            obolDeleted = False
            ostrCreatedBy = String.Empty
            odtCreatedOn = CDate("01/01/0001")
            ostrModifiedBy = String.Empty
            odtModifiedOn = CDate("01/01/0001")
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property OCEID() As Int64
            Get
                Return nOCEID
            End Get
            Set(ByVal Value As Int64)
                nOCEID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ID() As Int64
            Get
                Return nInsCITID
            End Get
            Set(ByVal Value As Int64)
                nInsCITID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property InspectionID() As Int64
            Get
                Return nInspectionID
            End Get
            Set(ByVal Value As Int64)
                nInspectionID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property QuestionID() As Int64
            Get
                Return nQuestionID
            End Get
            Set(ByVal Value As Int64)
                nQuestionID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FacilityID() As Int64
            Get
                Return nFacilityID
            End Get
            Set(ByVal Value As Int64)
                nFacilityID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FCEID() As Int64
            Get
                Return onFCEID
            End Get
            Set(ByVal Value As Int64)
                onFCEID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CitationID() As Int64
            Get
                Return nCitationID
            End Get
            Set(ByVal Value As Int64)
                nCitationID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CCAT() As String
            Get
                Return strCCAT
            End Get
            Set(ByVal Value As String)
                strCCAT = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Rescinded() As Boolean
            Get
                Return bolRescinded
            End Get
            Set(ByVal Value As Boolean)
                bolRescinded = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CitationDueDate() As Date
            Get
                Return dtCitationDueDate
            End Get
            Set(ByVal Value As Date)
                dtCitationDueDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CitationReceivedDate() As Date
            Get
                Return dtCitationReceivedDate
            End Get
            Set(ByVal Value As Date)
                dtCitationReceivedDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property NFADate() As Date
            Get
                Return dtNFADate
            End Get
            Set(ByVal Value As Date)
                dtNFADate = Value
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
                RaiseEvent evtInspectionCitationInfoChanged(bolIsDirty)
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
        Public Property CreatedOn() As Date
            Get
                Return dtCreatedOn
            End Get
            Set(ByVal Value As Date)
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
        Public Property ModifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
            Set(ByVal Value As Date)
                dtModifiedOn = Value
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
