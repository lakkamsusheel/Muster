'-------------------------------------------------------------------------------
' MUSTER.Info.InspectionChecklistMasterInfo
'   Provides the container to persist MUSTER InspectionChecklistMaster state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      MNR         6/10/05     Original class definition
'
' Function          Description
' New()             Instantiates an empty InspectionChecklistMasterInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated InspectionChecklistMasterInfo object
' New(dr)           Instantiates a populated InspectionChecklistMasterInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
' NOTE: This file to be used as InspectionChecklistMaster to build other objects.
'       Replace keyword "InspectionChecklistMaster" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class InspectionChecklistMasterInfo
#Region "Public Events"
        Public Event evtInspectionChecklistMasterInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nQuestionID As Int64
        Private nPosition As Int64
        Private strCheckListItemNumber As String
        Private strSOC As String
        Private bolHeader As Boolean
        Private strHeaderQuestionText As String
        'Private strResponseTable As String
        Private bolAppliesToTank As Boolean
        Private bolAppliesToPipe As Boolean
        Private bolAppliesToPipeTerm As Boolean
        Private nCitation As Int64
        Private strDiscrepText As String
        Private strWhenVisible As String
        Private bolCCAT As Boolean
        Private bolDeleted As Boolean
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime
        Private bolShow As Boolean
        Private strForeColor As String
        Private strBackColor As String

        Private onQuestionID As Int64
        Private onPosition As Int64
        Private ostrCheckListItemNumber As String
        Private ostrSOC As String
        Private obolHeader As Boolean
        Private ostrHeaderQuestionText As String
        'Private ostrResponseTable As String
        Private obolAppliesToTank As Boolean
        Private obolAppliesToPipe As Boolean
        Private obolAppliesToPipeTerm As Boolean
        Private onCitation As Int64
        Private ostrDiscrepText As String
        Private ostrWhenVisible As String
        Private obolCCAT As Boolean
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As DateTime
        Private obolShow As Boolean
        Private ostrForeColor As String
        Private ostrBackColor As String

        Private alTank As ArrayList
        Private alPipe As ArrayList
        Private alPipeTerm As ArrayList
        Private bolIsDirty As Boolean
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub
        Sub New(ByVal id As Int64, _
        ByVal position As Int64, _
        ByVal CheckListNum As String, _
        ByVal soc As String, _
        ByVal header As Boolean, _
        ByVal headertext As String, _
        ByVal appliesToTank As Boolean, _
        ByVal appliesToPipe As Boolean, _
        ByVal appliesToPipeTerm As Boolean, _
        ByVal citation As Int64, _
        ByVal discrepText As String, _
        ByVal whenvisible As String, _
        ByVal ccat As Boolean, _
        ByVal deleted As Boolean, _
        ByVal createdBy As String, _
        ByVal createdOn As Date, _
        ByVal modifiedBy As String, _
        ByVal modifiedon As Date, _
        ByVal foreColor As String, _
        ByVal backColor As String)
            onQuestionID = id
            onPosition = position
            ostrCheckListItemNumber = CheckListNum
            ostrSOC = soc
            obolHeader = header
            ostrHeaderQuestionText = headertext
            'ostrResponseTable = responsetable
            obolAppliesToTank = appliesToTank
            obolAppliesToPipe = appliesToPipe
            obolAppliesToPipeTerm = appliesToPipeTerm
            onCitation = citation
            ostrDiscrepText = discrepText
            ostrWhenVisible = whenvisible
            obolCCAT = ccat
            obolDeleted = deleted
            ostrCreatedBy = createdBy
            odtCreatedOn = createdOn
            ostrModifiedBy = modifiedBy
            odtModifiedOn = modifiedon
            ostrForeColor = foreColor
            ostrBackColor = backColor
            InitArrayList()
            Me.Reset()
        End Sub
        Sub New(ByVal drInspectionChecklistMaster As DataRow)
            Try
                onQuestionID = drInspectionChecklistMaster.Item("QUESTION_ID")
                onPosition = drInspectionChecklistMaster.Item("CHKLST_POSITION")
                ostrCheckListItemNumber = drInspectionChecklistMaster.Item("CHKLST_ITEM_NUMBER")
                ostrSOC = IIf(drInspectionChecklistMaster.Item("SOC") Is System.DBNull.Value, String.Empty, drInspectionChecklistMaster.Item("SOC"))
                obolHeader = drInspectionChecklistMaster.Item("HEADER")
                ostrHeaderQuestionText = IIf(drInspectionChecklistMaster.Item("HEADER_QUESTION_TEXT") Is System.DBNull.Value, String.Empty, drInspectionChecklistMaster.Item("HEADER_QUESTION_TEXT"))
                'ostrResponseTable = IIf(drInspectionChecklistMaster.Item("RESPONSE_TABLE") Is System.DBNull.Value, String.Empty, drInspectionChecklistMaster.Item("RESPONSE_TABLE"))
                obolAppliesToTank = drInspectionChecklistMaster.Item("APPLIES_TO_TANK")
                obolAppliesToPipe = drInspectionChecklistMaster.Item("APPLIES_TO_PIPE")
                obolAppliesToPipeTerm = drInspectionChecklistMaster.Item("APPLIES_TO_PIPETERM")
                onCitation = drInspectionChecklistMaster.Item("CITATION")
                ostrDiscrepText = IIf(drInspectionChecklistMaster.Item("DISCREP_TEXT") Is System.DBNull.Value, String.Empty, drInspectionChecklistMaster.Item("DISCREP_TEXT"))
                ostrWhenVisible = IIf(drInspectionChecklistMaster.Item("WHEN_VISIBLE") Is System.DBNull.Value, String.Empty, drInspectionChecklistMaster.Item("WHEN_VISIBLE"))
                obolCCAT = drInspectionChecklistMaster.Item("CCAT")
                obolDeleted = drInspectionChecklistMaster.Item("DELETED")
                ostrCreatedBy = IIf(drInspectionChecklistMaster.Item("CREATED_BY") Is System.DBNull.Value, String.Empty, drInspectionChecklistMaster.Item("CREATED_BY"))
                odtCreatedOn = IIf(drInspectionChecklistMaster.Item("DATE_CREATED") Is System.DBNull.Value, CDate("01/01/0001"), drInspectionChecklistMaster.Item("DATE_CREATED"))
                ostrModifiedBy = IIf(drInspectionChecklistMaster.Item("LAST_EDITED_BY") Is System.DBNull.Value, String.Empty, drInspectionChecklistMaster.Item("LAST_EDITED_BY"))
                odtModifiedOn = IIf(drInspectionChecklistMaster.Item("DATE_LAST_EDITED") Is System.DBNull.Value, CDate("01/01/0001"), drInspectionChecklistMaster.Item("DATE_LAST_EDITED"))
                ostrForeColor = IIf(drInspectionChecklistMaster.Item("FORE_COLOR") Is System.DBNull.Value, "BLACK", IIf(drInspectionChecklistMaster.Item("FORE_COLOR") = String.Empty, "BLACK", drInspectionChecklistMaster.Item("FORE_COLOR")))
                ostrBackColor = IIf(drInspectionChecklistMaster.Item("BACK_COLOR") Is System.DBNull.Value, "WHITE", IIf(drInspectionChecklistMaster.Item("BACK_COLOR") = String.Empty, "WHITE", drInspectionChecklistMaster.Item("BACK_COLOR")))
                InitArrayList()
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nQuestionID >= 0 Then
                nQuestionID = onQuestionID
            End If
            nPosition = onPosition
            strCheckListItemNumber = ostrCheckListItemNumber
            strSOC = ostrSOC
            bolHeader = obolHeader
            strHeaderQuestionText = ostrHeaderQuestionText
            'strResponseTable = ostrResponseTable
            bolAppliesToTank = obolAppliesToTank
            bolAppliesToPipe = obolAppliesToPipe
            bolAppliesToPipeTerm = obolAppliesToPipeTerm
            nCitation = onCitation
            strDiscrepText = ostrDiscrepText
            strWhenVisible = ostrWhenVisible
            bolCCAT = obolCCAT
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolShow = obolShow
            strForeColor = ostrForeColor
            strBackColor = ostrBackColor
            bolIsDirty = False
            RaiseEvent evtInspectionChecklistMasterInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onQuestionID = nQuestionID
            onPosition = nPosition
            ostrCheckListItemNumber = strCheckListItemNumber
            ostrSOC = strSOC
            obolHeader = bolHeader
            ostrHeaderQuestionText = strHeaderQuestionText
            'ostrResponseTable = strResponseTable
            obolAppliesToTank = bolAppliesToTank
            obolAppliesToPipe = bolAppliesToPipe
            obolAppliesToPipeTerm = bolAppliesToPipeTerm
            onCitation = nCitation
            ostrDiscrepText = strDiscrepText
            ostrWhenVisible = strWhenVisible
            obolCCAT = bolCCAT
            obolDeleted = bolDeleted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            obolShow = bolShow
            ostrForeColor = strForeColor
            ostrBackColor = strBackColor
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty

            bolIsDirty = (onPosition <> nPosition) Or _
                        (ostrCheckListItemNumber <> ostrCheckListItemNumber) Or _
                        (ostrSOC <> strSOC) Or _
                        (obolHeader <> bolHeader) Or _
                        (ostrHeaderQuestionText <> strHeaderQuestionText) Or _
                        (obolAppliesToTank <> obolAppliesToTank) Or _
                        (obolAppliesToPipe <> obolAppliesToPipe) Or _
                        (obolAppliesToPipeTerm <> obolAppliesToPipeTerm) Or _
                        (onCitation <> nCitation) Or _
                        (ostrDiscrepText <> strDiscrepText) Or _
                        (ostrWhenVisible <> strWhenVisible) Or _
                        (obolCCAT <> bolCCAT) Or _
                        (bolDeleted <> obolDeleted) Or _
                        (strCreatedBy <> ostrCreatedBy) Or _
                        (dtCreatedOn <> odtCreatedOn) Or _
                        (strModifiedBy <> ostrModifiedBy) Or _
                        (dtModifiedOn <> odtModifiedOn) Or _
                        (bolShow <> obolShow)
            '(strForeColor <> ostrForeColor) Or _
            '(strBackColor <> ostrBackColor)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent evtInspectionChecklistMasterInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onQuestionID = 0
            nPosition = 0
            ostrCheckListItemNumber = String.Empty
            ostrSOC = String.Empty
            obolHeader = False
            ostrHeaderQuestionText = String.Empty
            'ostrResponseTable = String.Empty
            obolAppliesToTank = False
            obolAppliesToPipe = False
            obolAppliesToPipeTerm = False
            onCitation = 0
            ostrDiscrepText = String.Empty
            ostrWhenVisible = String.Empty
            obolCCAT = False
            obolDeleted = False
            ostrCreatedBy = String.Empty
            odtCreatedOn = CDate("01/01/0001")
            ostrModifiedBy = String.Empty
            odtModifiedOn = CDate("01/01/0001")
            obolShow = False
            ostrForeColor = "BLACK"
            ostrBackColor = "WHITE"
            InitArrayList()
            Me.Reset()
        End Sub
        Private Sub InitArrayList()
            alTank = New ArrayList
            alPipe = New ArrayList
            alPipeTerm = New ArrayList
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int64
            Get
                Return nQuestionID
            End Get
            Set(ByVal Value As Int64)
                nQuestionID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Position() As Int64
            Get
                Return nPosition
            End Get
            Set(ByVal Value As Int64)
                nPosition = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CheckListItemNumber() As String
            Get
                Return strCheckListItemNumber
            End Get
            Set(ByVal Value As String)
                strCheckListItemNumber = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property SOC() As String
            Get
                Return strSOC
            End Get
            Set(ByVal Value As String)
                strSOC = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Header() As Boolean
            Get
                Return bolHeader
            End Get
            Set(ByVal Value As Boolean)
                bolHeader = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property HeaderQuestionText() As String
            Get
                Return strHeaderQuestionText
            End Get
            Set(ByVal Value As String)
                strHeaderQuestionText = Value
                Me.CheckDirty()
            End Set
        End Property
        'Public Property ResponseTable() As String
        '    Get
        '        Return strResponseTable
        '    End Get
        '    Set(ByVal Value As String)
        '        strResponseTable = Value
        '        Me.CheckDirty()
        '    End Set
        'End Property
        Public ReadOnly Property AppliesToTank() As Boolean
            Get
                Return bolAppliesToTank
            End Get
        End Property
        Public ReadOnly Property AppliesToPipe() As Boolean
            Get
                Return bolAppliesToPipe
            End Get
        End Property
        Public ReadOnly Property AppliesToPipeTerm() As Boolean
            Get
                Return bolAppliesToPipeTerm
            End Get
        End Property
        Public Property Citation() As Int64
            Get
                Return nCitation
            End Get
            Set(ByVal Value As Int64)
                nCitation = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property DiscrepText() As String
            Get
                Return strDiscrepText
            End Get
            Set(ByVal Value As String)
                strDiscrepText = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property WhenVisible() As String
            Get
                Return strWhenVisible
            End Get
            Set(ByVal Value As String)
                strWhenVisible = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CCAT() As Boolean
            Get
                Return bolCCAT
            End Get
            Set(ByVal Value As Boolean)
                bolCCAT = Value
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
        Public Property Show() As Boolean
            Get
                Return bolShow
            End Get
            Set(ByVal Value As Boolean)
                bolShow = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ForeColor() As String
            Get
                Return strForeColor
            End Get
            Set(ByVal Value As String)
                strForeColor = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property BackColor() As String
            Get
                Return strBackColor
            End Get
            Set(ByVal Value As String)
                strBackColor = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TankArrayList() As ArrayList
            Get
                Return alTank
            End Get
            Set(ByVal Value As ArrayList)
                alTank = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PipeArrayList() As ArrayList
            Get
                Return alPipe
            End Get
            Set(ByVal Value As ArrayList)
                alPipe = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PipeTermArrayList() As ArrayList
            Get
                Return alPipeTerm
            End Get
            Set(ByVal Value As ArrayList)
                alPipeTerm = Value
                Me.CheckDirty()
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
