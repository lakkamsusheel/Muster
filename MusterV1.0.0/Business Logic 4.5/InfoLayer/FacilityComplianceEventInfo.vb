'-------------------------------------------------------------------------------
' MUSTER.Info.FacilityComplianceEventInfo
'   Provides the container to persist MUSTER FacilityComplianceEvent state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       KUMAR      06/30/05    Original class definition
'
'
' Function          Description
' New()             Instantiates an empty FacilityComplianceEventInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated FacilityComplianceEventInfo object
' New(dr)           Instantiates a populated FacilityComplianceEventInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
' NOTE: This file to be used as FacilityComplianceEvent to build other objects.
'       Replace keyword "FacilityComplianceEvent" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class FacilityComplianceEventInfo
#Region "Public Events"
        Public Event FCEInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"

        Private nFCEID As Int32
        Private nInspectionID As Int32
        Private nOwnerID As Int32
        Private nFacilityID As Int32
        Private dtFCEDate As DateTime
        Private strSource As String
        Private dtDueDate As DateTime
        Private dtReceivedDate As DateTime
        'Private strOwnerName As String
        'Private strFacilityName As String
        'Private strInspectorName As String
        'Private dtInspectedOn As DateTime
        'Private nCitations As Integer
        Private strCreatedBy As String
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String
        Private dtModifiedOn As DateTime
        Private bolDeleted As Boolean
        Private bolOCEGenerated As Boolean
        Private nOCEID As Int32

        Private onFCEID As Int32
        Private onInspectionID As Int32
        Private onOwnerID As Int32
        Private onFacilityID As Int32
        Private odtFCEDate As DateTime
        Private ostrSource As String
        Private odtDueDate As DateTime
        Private odtReceivedDate As DateTime
        'Private ostrOwnerName As String
        'Private ostrFacilityName As String
        'Private ostrInspectorName As String
        'Private odtInspectedOn As DateTime
        'Private onCitations As Integer
        Private ostrCreatedBy As String
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String
        Private odtModifiedOn As DateTime
        Private obolDeleted As Boolean
        Private obolOCEGenerated As Boolean
        Private onOCEID As Int32

        'Private bolSelected1 As Boolean
        'Private bolSelected2 As Boolean
        'Private obolSelected1 As Boolean
        'Private obolSelected2 As Boolean
        'Private bolShowDeleted As Boolean = False

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub
        Sub New(ByVal FCEID As Int32, _
           ByVal InspectionID As Int32, _
           ByVal OwnerID As Int32, _
           ByVal FacilityID As Int32, _
           ByVal FCEDate As Date, _
           ByVal Source As String, _
           ByVal DueDate As Date, _
           ByVal ReceivedDate As Date, _
           ByVal CreatedBy As String, _
           ByVal CreatedOn As Date, _
           ByVal ModifiedBy As String, _
           ByVal ModifiedOn As Date, _
           ByVal Deleted As Boolean, _
           ByVal OCEGenerated As Boolean, _
           ByVal oceID As Int32)
            onFCEID = FCEID
            onInspectionID = InspectionID
            onOwnerID = OwnerID
            onFacilityID = FacilityID
            odtFCEDate = FCEDate
            ostrSource = Source
            odtDueDate = DueDate
            odtReceivedDate = ReceivedDate
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = ModifiedOn
            obolDeleted = Deleted
            obolOCEGenerated = OCEGenerated
            onOCEID = oceID
            Me.Reset()
        End Sub
        'Sub New(ByVal FCEID As Int32, _
        '   ByVal InspectionID As Int32, _
        '   ByVal OwnerID As Int32, _
        '   ByVal FacilityID As Int32, _
        '   ByVal FCEDate As Date, _
        '   ByVal Source As String, _
        '   ByVal DueDate As Date, _
        '   ByVal ReceivedDate As Date, _
        '   ByVal OwnerName As String, _
        '   ByVal FacilityName As String, _
        '   ByVal InspectorName As String, _
        '   ByVal InspectedOn As DateTime, _
        '   ByVal CItations As Integer, _
        '   ByVal CreatedBy As String, _
        '   ByVal CreatedOn As Date, _
        '   ByVal ModifiedBy As String, _
        '   ByVal ModifiedOn As Date, _
        '   ByVal Deleted As Boolean, _
        '   ByVal OCEGenerated As Boolean)
        '    onFCEID = FCEID
        '    onInspectionID = InspectionID
        '    onOwnerID = OwnerID
        '    onFacilityID = FacilityID
        '    odtFCEDate = FCEDate
        '    ostrSource = Source
        '    odtDueDate = DueDate
        '    odtReceivedDate = ReceivedDate
        '    ostrOwnerName = OwnerName
        '    ostrFacilityName = FacilityName
        '    ostrInspectorName = InspectorName
        '    odtInspectedOn = InspectedOn
        '    onCitations = CItations
        '    ostrCreatedBy = CreatedBy
        '    odtCreatedOn = CreatedOn
        '    ostrModifiedBy = ModifiedBy
        '    odtModifiedOn = ModifiedOn
        '    obolDeleted = Deleted
        '    obolOCEGenerated = OCEGenerated
        '    Me.Reset()
        'End Sub
        Sub New(ByVal dr As DataRow)
            Try
                onFCEID = dr.Item("FCE_ID")
                onInspectionID = dr.Item("INSPECTION_ID")
                onOwnerID = dr.Item("OWNER_ID")
                onFacilityID = dr.Item("FACILITY_ID")
                odtFCEDate = dr.Item("FCE_DATE")
                ostrSource = IIf(dr.Item("SOURCE") Is DBNull.Value, String.Empty, dr.Item("SOURCE"))
                odtDueDate = IIf(dr.Item("DUE_DATE") Is DBNull.Value, CDate("01/01/0001"), dr.Item("DUE_DATE"))
                odtReceivedDate = IIf(dr.Item("RECEIVED_DATE") Is DBNull.Value, CDate("01/01/0001"), dr.Item("RECEIVED_DATE"))
                'ostrOwnerName = OwnerName
                'ostrFacilityName = FacilityName
                'ostrInspectorName = InspectorName
                'odtInspectedOn = InspectedOn
                'onCitations = Citations
                ostrCreatedBy = IIf(dr.Item("CREATED_BY") Is DBNull.Value, String.Empty, dr.Item("CREATED_BY"))
                odtCreatedOn = IIf(dr.Item("DATE_CREATED") Is DBNull.Value, CDate("01/01/0001"), dr.Item("DATE_CREATED"))
                ostrModifiedBy = IIf(dr.Item("LAST_EDITED_BY") Is DBNull.Value, String.Empty, dr.Item("LAST_EDITED_BY"))
                odtModifiedOn = IIf(dr.Item("DATE_LAST_EDITED") Is DBNull.Value, CDate("01/01/0001"), dr.Item("DATE_LAST_EDITED"))
                obolDeleted = IIf(dr.Item("DELETED") Is DBNull.Value, False, dr.Item("DELETED"))
                obolOCEGenerated = dr.Item("OCE_GENERATED")
                onOCEID = IIf(dr.Item("OCE_ID") Is DBNull.Value, 0, dr.Item("OCE_ID"))
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nFCEID >= 0 Then
                nFCEID = onFCEID
            End If
            nInspectionID = onInspectionID
            nOwnerID = onOwnerID
            nFacilityID = onFacilityID
            dtFCEDate = odtFCEDate
            strSource = ostrSource
            dtDueDate = odtDueDate
            dtReceivedDate = odtReceivedDate
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolOCEGenerated = obolOCEGenerated
            nOCEID = onOCEID
            bolIsDirty = False
            'strOwnerName = ostrOwnerName
            'strFacilityName = ostrFacilityName
            'strInspectorName = ostrInspectorName
            'dtInspectedOn = odtInspectedOn
            'nCitations = onCitations
            'bolSelected1 = obolSelected1
            'bolSelected2 = obolSelected2
            RaiseEvent FCEInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onFCEID = nFCEID
            onInspectionID = nInspectionID
            onOwnerID = nOwnerID
            onFacilityID = nFacilityID
            odtFCEDate = dtFCEDate
            ostrSource = strSource
            odtDueDate = dtDueDate
            odtReceivedDate = dtReceivedDate
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            obolDeleted = bolDeleted
            obolOCEGenerated = bolOCEGenerated
            onOCEID = nOCEID
            bolIsDirty = False
            'ostrOwnerName = strOwnerName
            'ostrFacilityName = strFacilityName
            'ostrInspectorName = strInspectorName
            'odtInspectedOn = dtInspectedOn
            'onCitations = nCitations
            'obolSelected1 = bolSelected1
            'obolSelected2 = bolSelected2
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty

            bolIsDirty = (nInspectionID <> onInspectionID) Or _
            (nOwnerID <> onOwnerID) Or _
            (nFacilityID <> onFacilityID) Or _
            (dtFCEDate <> odtFCEDate) Or _
            (strSource <> ostrSource) Or _
            (dtDueDate <> odtDueDate) Or _
            (dtReceivedDate <> odtReceivedDate) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn) Or _
            (bolDeleted <> obolDeleted) Or _
            (bolOCEGenerated <> obolOCEGenerated) Or _
            (nOCEID <> onOCEID)

            '(strOwnerName <> ostrOwnerName) Or _
            '(strFacilityName <> ostrFacilityName) Or _
            '(strInspectorName <> ostrInspectorName) Or _
            '(dtInspectedOn <> odtInspectedOn) Or _
            '(nCitations <> onCitations) Or _
            '(bolSelected1 <> obolSelected1) Or _
            '(bolSelected2 <> obolSelected2) Or _
            If obolIsDirty <> bolIsDirty Then
                RaiseEvent FCEInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onFCEID = 0
            onInspectionID = 0
            onOwnerID = 0
            onFacilityID = 0
            odtFCEDate = CDate("01/01/0001")
            ostrSource = 0
            odtDueDate = CDate("01/01/0001")
            odtReceivedDate = CDate("01/01/0001")
            ostrCreatedBy = String.Empty
            odtCreatedOn = CDate("01/01/0001")
            ostrModifiedBy = String.Empty
            odtModifiedOn = CDate("01/01/0001")
            obolDeleted = False
            obolOCEGenerated = False
            onOCEID = 0
            bolIsDirty = False
            'ostrOwnerName = String.Empty
            'ostrFacilityName = String.Empty
            'ostrInspectorName = String.Empty
            'odtInspectedOn = CDate("01/01/0001")
            'onCitations = 0
            'obolSelected1 = False
            'obolSelected2 = False
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int32
            Get
                Return nFCEID
            End Get
            Set(ByVal Value As Int32)
                nFCEID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property InspectionID() As Int32
            Get
                Return nInspectionID
            End Get
            Set(ByVal Value As Int32)
                nInspectionID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property OwnerID() As Int32
            Get
                Return nOwnerID
            End Get
            Set(ByVal Value As Int32)
                nOwnerID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FacilityID() As Int32
            Get
                Return nFacilityID
            End Get
            Set(ByVal Value As Int32)
                nFacilityID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FCEDate() As DateTime
            Get
                Return dtFCEDate
            End Get
            Set(ByVal Value As DateTime)
                dtFCEDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Source() As String
            Get
                Return strSource
            End Get
            Set(ByVal Value As String)
                strSource = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property DueDate() As DateTime
            Get
                Return dtDueDate
            End Get
            Set(ByVal Value As DateTime)
                dtDueDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ReceivedDate() As DateTime
            Get
                Return dtReceivedDate
            End Get
            Set(ByVal Value As DateTime)
                dtReceivedDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return strCreatedBy
            End Get
            Set(ByVal Value As String)
                strCreatedBy = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CreatedOn() As DateTime
            Get
                Return dtCreatedOn
            End Get
            Set(ByVal Value As DateTime)
                dtCreatedOn = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ModifiedOn() As DateTime
            Get
                Return dtModifiedOn
            End Get
            Set(ByVal Value As DateTime)
                dtModifiedOn = Value
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
        Public Property OCEGenerated() As Boolean
            Get
                Return bolOCEGenerated
            End Get
            Set(ByVal Value As Boolean)
                bolOCEGenerated = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property OCEID() As Int32
            Get
                Return nOCEID
            End Get
            Set(ByVal Value As Int32)
                nOCEID = Value
                Me.CheckDirty()
            End Set
        End Property
        'Public Property selected1() As Boolean
        '    Get
        '        Return bolSelected1
        '    End Get
        '    Set(ByVal Value As Boolean)
        '        bolSelected1 = Value
        '    End Set
        'End Property
        'Public Property selected2() As Boolean
        '    Get
        '        Return bolSelected2
        '    End Get
        '    Set(ByVal Value As Boolean)
        '        bolSelected2 = Value
        '    End Set
        'End Property
        'Public Property OwnerName() As String
        '    Get
        '        Return strOwnerName
        '    End Get
        '    Set(ByVal Value As String)
        '        strOwnerName = Value
        '        Me.CheckDirty()
        '    End Set
        'End Property
        'Public Property FacilityName() As String
        '    Get
        '        Return strFacilityName
        '    End Get
        '    Set(ByVal Value As String)
        '        strFacilityName = Value
        '        Me.CheckDirty()
        '    End Set
        'End Property
        'Public Property InspectorName() As String
        '    Get
        '        Return strInspectorName
        '    End Get
        '    Set(ByVal Value As String)
        '        strInspectorName = Value
        '        Me.CheckDirty()
        '    End Set
        'End Property
        'Public Property InspectedOn() As DateTime
        '    Get
        '        Return dtInspectedOn
        '    End Get
        '    Set(ByVal Value As DateTime)
        '        dtInspectedOn = Value
        '        Me.CheckDirty()
        '    End Set
        'End Property
        'Public Property Citations() As Integer
        '    Get
        '        Return nCitations
        '    End Get
        '    Set(ByVal Value As Integer)
        '        nCitations = Value
        '        Me.CheckDirty()
        '    End Set
        'End Property

        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get
            Set(ByVal value As Boolean)
                bolIsDirty = value
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
