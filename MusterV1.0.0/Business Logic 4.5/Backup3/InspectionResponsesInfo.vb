'-------------------------------------------------------------------------------
' MUSTER.Info.InspectionResponsesInfo
'   Provides the container to persist MUSTER InspectionResponses state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MNR        06/14/05    Original class definition
'
' Function          Description
' New()             Instantiates an empty InspectionResponsesInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated InspectionResponsesInfo object
' New(dr)           Instantiates a populated InspectionResponsesInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
' NOTE: This file to be used as InspectionResponses to build other objects.
'       Replace keyword "InspectionResponses" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class InspectionResponsesInfo
#Region "Public Events"
        Public Event evtInspectionResponsesInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nInsRespID As Int64
        Private nInspectionID As Int64
        Private nQuestionID As Int64
        Private bolSOC As Boolean
        Private nResponse As Int64
        Private bolDeleted As Boolean
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime

        Private onInsRespID As Int64
        Private onInspectionID As Int64
        Private onQuestionID As Int64
        Private obolSOC As Boolean
        Private onResponse As Int64
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As DateTime

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub
        Sub New(ByVal InsRespID As Int64, _
        ByVal inspectionID As Int64, _
        ByVal questionID As Int64, _
        ByVal soc As Boolean, _
        ByVal response As Int64, _
        ByVal deleted As Boolean, _
        ByVal createdby As String, _
        ByVal createdon As Date, _
        ByVal modifiedby As String, _
        ByVal modifiedon As Date)
            onInsRespID = InsRespID
            onInspectionID = inspectionID
            onQuestionID = questionID
            obolSOC = soc
            onResponse = response
            obolDeleted = deleted
            ostrCreatedBy = createdby
            odtCreatedOn = createdon
            ostrModifiedBy = modifiedby
            odtModifiedOn = modifiedon
            Me.Reset()
        End Sub
        Sub New(ByVal drInspectionResponses As DataRow)
            Try
                onInsRespID = drInspectionResponses.Item("INS_RESP_ID")
                onInspectionID = IIf(drInspectionResponses.Item("INSPECTION_ID") Is DBNull.Value, 0, drInspectionResponses.Item("INSPECTION_ID"))
                onQuestionID = IIf(drInspectionResponses.Item("QUESTION_ID") Is DBNull.Value, 0, drInspectionResponses.Item("QUESTION_ID"))
                obolSOC = IIf(drInspectionResponses.Item("SOC") Is DBNull.Value, False, drInspectionResponses.Item("SOC"))
                onResponse = IIf(drInspectionResponses.Item("RESPONSE") Is DBNull.Value, -1, drInspectionResponses.Item("RESPONSE"))
                obolDeleted = IIf(drInspectionResponses.Item("DELETED") Is DBNull.Value, False, drInspectionResponses.Item("DELETED"))
                ostrCreatedBy = IIf(drInspectionResponses.Item("CREATED_BY") Is DBNull.Value, String.Empty, drInspectionResponses.Item("CREATED_BY"))
                odtCreatedOn = IIf(drInspectionResponses.Item("DATE_CREATED") Is DBNull.Value, CDate("01/01/0001"), drInspectionResponses.Item("DATE_CREATED"))
                ostrModifiedBy = IIf(drInspectionResponses.Item("LAST_EDITED_BY") Is DBNull.Value, String.Empty, drInspectionResponses.Item("LAST_EDITED_BY"))
                odtModifiedOn = IIf(drInspectionResponses.Item("DATE_LAST_EDITED") Is DBNull.Value, CDate("01/01/0001"), drInspectionResponses.Item("DATE_LAST_EDITED"))
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If onInsRespID >= 0 Then
                nInsRespID = onInsRespID
            End If
            nInspectionID = onInspectionID
            nQuestionID = onQuestionID
            bolSOC = obolSOC
            nResponse = onResponse
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
            RaiseEvent evtInspectionResponsesInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onInspectionID = nInspectionID
            onQuestionID = nQuestionID
            obolSOC = bolSOC
            onResponse = nResponse
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
            (bolSOC <> obolSOC) Or _
            (nResponse <> onResponse) Or _
            (bolDeleted <> obolDeleted) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent evtInspectionResponsesInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            ' Response -> -1 = Nothing, 0 = False/NO, 1 = True/YES
            onInsRespID = 0
            onInspectionID = 0
            onQuestionID = 0
            obolSOC = False
            onResponse = -1
            obolDeleted = False
            ostrCreatedBy = String.Empty
            odtCreatedOn = CDate("01/01/0001")
            ostrModifiedBy = String.Empty
            odtModifiedOn = CDate("01/01/0001")
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int64
            Get
                Return nInsRespID
            End Get
            Set(ByVal Value As Int64)
                nInsRespID = Value
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
        Public Property SOC() As Boolean
            Get
                Return bolSOC
            End Get
            Set(ByVal Value As Boolean)
                bolSOC = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Response() As Int64
            Get
                Return nResponse
            End Get
            Set(ByVal Value As Int64)
                nResponse = Value
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
                RaiseEvent evtInspectionResponsesInfoChanged(bolIsDirty)
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
