'-------------------------------------------------------------------------------
' MUSTER.Info.InspectionSOCInfo
'   Provides the container to persist MUSTER InspectionSOC state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MNR        06/15/05    Original class definition
'
' Function          Description
' New()             Instantiates an empty InspectionSOCInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated InspectionSOCInfo object
' New(dr)           Instantiates a populated InspectionSOCInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
' NOTE: This file to be used as InspectionSOC to build other objects.
'       Replace keyword "InspectionSOC" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class InspectionSOCInfo
#Region "Public Events"
        Public Event evtInspectionSOCInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nInsTosID As Int64
        Private nInspectionID As Int64
        Private nLeakPrevention As Int64
        Private strLeakPreventionCitation As String
        Private strLeakPreventionLineNumbers As String
        Private nLeakDetection As Int64
        Private strLeakDetectionCitation As String
        Private strLeakDetectionLineNumbers As String
        Private nLeakPreventionDetection As Int64
        Private bolDeleted As Boolean
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime
        Private bolCAEOverride As Boolean

        Private onInsTosID As Int64
        Private onInspectionID As Int64
        Private onLeakPrevention As Int64
        Private ostrLeakPreventionCitation As String
        Private ostrLeakPreventionLineNumbers As String
        Private onLeakDetection As Int64
        Private ostrLeakDetectionCitation As String
        Private ostrLeakDetectionLineNumbers As String
        Private onLeakPreventionDetection As Int64
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As DateTime
        Private obolCAEOverride As Boolean

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
        ByVal FSOC_LK_PREVENT As Int64, _
        ByVal FSOC_LK_PRE_CITATION As String, _
        ByVal FSOC_LK_PRE_LINENUMBERS As String, _
        ByVal FAC_SOC_LK_DETECTION As Int64, _
        ByVal FSOC_LK_DET_CITATION As String, _
        ByVal FSOC_LK_DET_LINENUMBERS As String, _
        ByVal FSOC_LK_PRE_LK_DET As Int64, _
        ByVal deleted As Boolean, _
        ByVal createdBy As String, _
        ByVal createdOn As Date, _
        ByVal modifiedBy As String, _
        ByVal modifiedOn As Date, _
        ByVal caeOverride As Boolean)
            onInsTosID = id
            onInspectionID = inspectionID
            onLeakPrevention = FSOC_LK_PREVENT
            ostrLeakPreventionCitation = FSOC_LK_PRE_CITATION
            ostrLeakPreventionLineNumbers = FSOC_LK_PRE_LINENUMBERS
            onLeakDetection = FAC_SOC_LK_DETECTION
            ostrLeakDetectionCitation = FSOC_LK_DET_CITATION
            ostrLeakDetectionLineNumbers = FSOC_LK_DET_LINENUMBERS
            onLeakPreventionDetection = FSOC_LK_PRE_LK_DET
            obolDeleted = deleted
            ostrCreatedBy = createdBy
            odtCreatedOn = createdOn
            ostrModifiedBy = modifiedBy
            odtModifiedOn = modifiedOn
            obolCAEOverride = caeOverride
            Me.Reset()
        End Sub
        Sub New(ByVal drInspectionSOC As DataRow)
            Try
                onInsTosID = drInspectionSOC.Item("INS_TOS_ID")
                onInspectionID = IIf(drInspectionSOC.Item("INSPECTION_ID") Is DBNull.Value, 0, drInspectionSOC.Item("INSPECTION_ID"))
                onLeakPrevention = IIf(drInspectionSOC.Item("FSOC_LK_PREVENT") Is DBNull.Value, -1, drInspectionSOC.Item("FSOC_LK_PREVENT"))
                ostrLeakPreventionCitation = IIf(drInspectionSOC.Item("FSOC_LK_PRE_CITATION") Is DBNull.Value, String.Empty, drInspectionSOC.Item("FSOC_LK_PRE_CITATION"))
                ostrLeakPreventionLineNumbers = IIf(drInspectionSOC.Item("FSOC_LK_PRE_LINE_NUMBERS") Is DBNull.Value, String.Empty, drInspectionSOC.Item("FSOC_LK_PRE_LINE_NUMBERS"))
                onLeakDetection = IIf(drInspectionSOC.Item("FAC_SOC_LK_DETECTION") Is DBNull.Value, -1, drInspectionSOC.Item("FAC_SOC_LK_DETECTION"))
                ostrLeakDetectionCitation = IIf(drInspectionSOC.Item("FSOC_LK_DET_CITATION") Is DBNull.Value, String.Empty, drInspectionSOC.Item("FSOC_LK_DET_CITATION"))
                ostrLeakDetectionLineNumbers = IIf(drInspectionSOC.Item("FSOC_LK_DET_LINENUMBERS") Is DBNull.Value, String.Empty, drInspectionSOC.Item("FSOC_LK_DET_LINENUMBERS"))
                onLeakPreventionDetection = IIf(drInspectionSOC.Item("FSOC_LK_PRE_LK_DET") Is DBNull.Value, -1, drInspectionSOC.Item("FSOC_LK_PRE_LK_DET"))
                obolDeleted = IIf(drInspectionSOC.Item("DELETED") Is DBNull.Value, False, drInspectionSOC.Item("DELETED"))
                ostrCreatedBy = IIf(drInspectionSOC.Item("CREATED_BY") Is DBNull.Value, String.Empty, drInspectionSOC.Item("CREATED_BY"))
                odtCreatedOn = IIf(drInspectionSOC.Item("DATE_CREATED") Is DBNull.Value, CDate("01/01/0001"), drInspectionSOC.Item("DATE_CREATED"))
                ostrModifiedBy = IIf(drInspectionSOC.Item("LAST_EDITED_BY") Is DBNull.Value, String.Empty, drInspectionSOC.Item("LAST_EDITED_BY"))
                odtModifiedOn = IIf(drInspectionSOC.Item("DATE_LAST_EDITED") Is DBNull.Value, CDate("01/01/0001"), drInspectionSOC.Item("DATE_LAST_EDITED"))
                obolCAEOverride = IIf(drInspectionSOC.Item("CAE_OVERRIDE") Is DBNull.Value, False, drInspectionSOC.Item("CAE_OVERRIDE"))
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nInsTosID >= 0 Then
                nInsTosID = onInsTosID
            End If
            nInspectionID = onInspectionID
            nLeakPrevention = onLeakPrevention
            strLeakPreventionCitation = ostrLeakPreventionCitation
            strLeakPreventionLineNumbers = ostrLeakPreventionLineNumbers
            nLeakDetection = onLeakDetection
            strLeakDetectionCitation = ostrLeakDetectionCitation
            strLeakDetectionLineNumbers = ostrLeakDetectionLineNumbers
            nLeakPreventionDetection = onLeakPreventionDetection
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolCAEOverride = obolCAEOverride
            bolIsDirty = False
            RaiseEvent evtInspectionSOCInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onInsTosID = nInsTosID
            onInspectionID = nInspectionID
            onLeakPrevention = nLeakPrevention
            ostrLeakPreventionCitation = strLeakPreventionCitation
            ostrLeakPreventionLineNumbers = strLeakPreventionLineNumbers
            onLeakDetection = nLeakDetection
            ostrLeakDetectionCitation = strLeakDetectionCitation
            ostrLeakDetectionLineNumbers = strLeakDetectionLineNumbers
            onLeakPreventionDetection = nLeakPreventionDetection
            obolDeleted = bolDeleted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            obolCAEOverride = bolCAEOverride
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty

            bolIsDirty = (nInspectionID <> onInspectionID) Or _
            (nLeakPrevention <> onLeakPrevention) Or _
            (strLeakPreventionCitation <> ostrLeakPreventionCitation) Or _
            (strLeakPreventionLineNumbers <> ostrLeakPreventionLineNumbers) Or _
            (nLeakDetection <> onLeakDetection) Or _
            (strLeakDetectionCitation <> ostrLeakDetectionCitation) Or _
            (strLeakDetectionLineNumbers <> ostrLeakDetectionLineNumbers) Or _
            (nLeakPreventionDetection <> onLeakPreventionDetection) Or _
            (bolDeleted <> obolDeleted) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn) Or _
            (bolCAEOverride <> obolCAEOverride)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent evtInspectionSOCInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onInsTosID = 0
            onInspectionID = 0
            onLeakPrevention = -1
            ostrLeakPreventionCitation = String.Empty
            ostrLeakPreventionLineNumbers = String.Empty
            onLeakDetection = -1
            ostrLeakDetectionCitation = String.Empty
            ostrLeakDetectionLineNumbers = String.Empty
            onLeakPreventionDetection = -1
            obolDeleted = False
            ostrCreatedBy = String.Empty
            odtCreatedOn = System.DateTime.Now
            ostrModifiedBy = String.Empty
            odtModifiedOn = System.DateTime.Now
            obolCAEOverride = False
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int64
            Get
                Return nInsTosID
            End Get
            Set(ByVal Value As Int64)
                nInsTosID = Value
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
        Public Property LeakPrevention() As Int64
            Get
                Return nLeakPrevention
            End Get
            Set(ByVal Value As Int64)
                nLeakPrevention = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LeakPreventionCitation() As String
            Get
                Return strLeakPreventionCitation
            End Get
            Set(ByVal Value As String)
                strLeakPreventionCitation = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LeakPreventionLineNumbers() As String
            Get
                Return strLeakPreventionLineNumbers
            End Get
            Set(ByVal Value As String)
                strLeakPreventionLineNumbers = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LeakDetection() As Int64
            Get
                Return nLeakDetection
            End Get
            Set(ByVal Value As Int64)
                nLeakDetection = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LeakDetectionCitation() As String
            Get
                Return strLeakDetectionCitation
            End Get
            Set(ByVal Value As String)
                strLeakDetectionCitation = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LeakDetectionLineNumbers() As String
            Get
                Return strLeakDetectionLineNumbers
            End Get
            Set(ByVal Value As String)
                strLeakDetectionLineNumbers = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LeakPreventionDetection() As Int64
            Get
                Return nLeakPreventionDetection
            End Get
            Set(ByVal Value As Int64)
                nLeakPreventionDetection = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CAEOverride() As Boolean
            Get
                Return bolCAEOverride
            End Get

            Set(ByVal value As Boolean)
                bolCAEOverride = value
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
                RaiseEvent evtInspectionSOCInfoChanged(bolIsDirty)
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
