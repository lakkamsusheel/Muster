'-------------------------------------------------------------------------------
' MUSTER.Info.ReportInfo
'   Provides the container to persist MUSTER Report state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC2      11/23/04    Original class definition.
'  1.1        AN        12/30/04    Added Try catch and Exception Handling/Logging
'  1.2        AN         1/05/05    Added Events for Report Data Changed (Save/Cancel)
'  1.3        AN         1/21/05    Added  Raise Event for Reset
'  1.4        AB        02/22/05    Added AgeThreshold and IsAgedData Attributes
'  1.5        JC        07/26/05    WTF???  Just added the Archive operation - what gives?
'
' Function          Description
' New()             Instantiates an empty ReportInfo object.
' New(ID, Name, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
'                   Instantiates a populated ReportInfo object.
' New(dr)           Instantiates a populated ReportInfo object taking member state
'                       from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                       last saved to the repository.
'
'TODO - Complete Attributes List
'
'Attribute          Description
' ID                The unique identifier associated with the Report in the repository.
' Name              The name of the Entity.
' Description
' Module
' Path
' Deleted
' IsDirty           Indicates if the Entity state has been altered since it was
'                       last loaded from or saved to the repository.
'-------------------------------------------------------------------------------

Namespace MUSTER.Info

    <Serializable()> _
      Public Class ReportInfo

#Region "Private member variables"

        Private nReportID As Int64                  'The internal ID number associated with the user group
        Private strName As String                   'The name of the user group
        Private strModule As String                 'The module the report is associated with
        Private strDescription As String            'The description of the report
        Private strReportLoc As String              'The location that the report resides in
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime = DateTime.Now.ToShortDateString
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime = DateTime.Now.ToShortDateString
        Private bolDeleted As Boolean
        Private bolActive As Boolean
        'Private colAssociatedGroups As UserGroups   'The list of user groups associated with the form
        Private dtReportNames As DataTable         'Lists the report names for the client
        Private bolFavoriteReport As Boolean = False

        Private onReportID As Int64
        Private ostrName As String
        Private ostrDescription As String
        Private ostrModule As String
        Private ostrReportLoc As String
        Private obolDeleted As Boolean
        Private oBolActive As Boolean
        Private bolShowDeleted As Boolean = False
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As DateTime = DateTime.Now.ToShortDateString
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As DateTime = DateTime.Now.ToShortDateString
        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions

        Private colReportGroupRelation As MUSTER.Info.ReportGroupRelationsCollection
#End Region
#Region "Public Events"
        Public Event ReportChanged(ByVal bolValue As Boolean)
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Init()
            InitCollection()
            dtDataAge = Now()
        End Sub
        Sub New(ByVal ReportID As Integer, _
            ByVal ReportName As String, _
            ByVal ReportModule As String, _
            ByVal ReportDescription As String, _
            ByVal ReportLoc As String, _
            ByVal Deleted As Boolean, _
            ByVal CreatedBy As String, _
            ByVal CreatedOn As Date, _
            ByVal ModifiedBy As String, _
            ByVal LastEdited As Date, _
            ByVal Active As Boolean)
            onReportID = ReportID
            ostrName = ReportName
            ostrDescription = ReportDescription
            ostrModule = ReportModule
            ostrReportLoc = ReportLoc
            obolDeleted = Deleted
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = LastEdited
            oBolActive = Active
            dtDataAge = Now()
            InitCollection()
            Me.Reset()
        End Sub
        Sub New(ByVal drReport As DataRow)
            Try
                onReportID = drReport.Item("REPORT_ID")
                ostrName = drReport.Item("REPORT_NAME")
                ostrDescription = drReport.Item("REPORT_DESC")
                ostrModule = drReport.Item("REPORT_MODULE")
                ostrReportLoc = drReport.Item("REPORT_LOC")
                obolDeleted = Not drReport.Item("ACTIVE")
                ostrCreatedBy = drReport.Item("CREATED_BY")
                odtCreatedOn = drReport.Item("DATE_CREATED")
                ostrModifiedBy = drReport.Item("LAST_EDITED_BY")
                odtModifiedOn = drReport.Item("DATE_LAST_EDITED")
                oBolActive = drReport.Item("ACTIVE")
                dtDataAge = Now()
                InitCollection()
                Me.Reset()
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            nReportID = onReportID
            strName = ostrName
            strDescription = ostrDescription
            strModule = ostrModule
            strReportLoc = ostrReportLoc
            bolDeleted = obolDeleted
            bolActive = oBolActive

            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn

            bolIsDirty = False
            ResetGroupReportRelationCollection()
            RaiseEvent ReportChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onReportID = nReportID
            ostrName = strName
            ostrDescription = strDescription
            ostrModule = strModule
            ostrReportLoc = strReportLoc
            obolDeleted = bolDeleted
            oBolActive = bolActive
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn

            bolIsDirty = False
            RaiseEvent ReportChanged(bolIsDirty)

        End Sub
        Public Sub ResetGroupReportRelationCollection()
            For Each groupReportRelInfo As MUSTER.Info.ReportGroupRelationInfo In colReportGroupRelation.Values
                groupReportRelInfo.Reset()
            Next
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim bolOldValue As Boolean = bolIsDirty

            bolIsDirty = (nReportID <> onReportID) Or _
                         (strName <> ostrName) Or _
                         (strDescription <> ostrDescription) Or _
                         (strModule <> ostrModule) Or _
                         (strReportLoc <> ostrReportLoc) Or _
                         (bolDeleted <> obolDeleted) Or _
                         (bolActive <> oBolActive)

            If Not bolIsDirty Then
                For Each groupReportRelInfo As MUSTER.Info.ReportGroupRelationInfo In colReportGroupRelation.Values
                    If groupReportRelInfo.IsDirty Then
                        bolIsDirty = True
                        Exit For
                    End If
                Next
            End If

            If bolOldValue <> bolIsDirty Then
                RaiseEvent ReportChanged(bolIsDirty)
            End If

        End Sub
        Private Sub Init()
            onReportID = 0
            ostrName = String.Empty
            ostrDescription = String.Empty
            ostrModule = String.Empty
            ostrReportLoc = String.Empty
            obolDeleted = False
            oBolActive = True
            dtCreatedOn = System.DateTime.Now
            dtModifiedOn = System.DateTime.Now
            strCreatedBy = String.Empty
            strModifiedBy = String.Empty
            InitCollection()
            Me.Reset()

        End Sub
        Private Sub InitCollection()
            colReportGroupRelation = New MUSTER.Info.ReportGroupRelationsCollection
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return nReportID
            End Get

            Set(ByVal value As Integer)
                nReportID = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property

        Public Property Name() As String
            Get
                Return strName
            End Get

            Set(ByVal value As String)
                strName = value
                Me.CheckDirty()
            End Set
        End Property

        Public ReadOnly Property FileName() As String
            Get
                Return strReportLoc.Substring(strReportLoc.LastIndexOf("\") + 1, strReportLoc.Length - (strReportLoc.LastIndexOf("\") + 1))
            End Get
        End Property

        Public Property Description() As String
            Get
                Return strDescription
            End Get

            Set(ByVal value As String)
                strDescription = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property [Module]() As String
            Get
                Return strModule
            End Get

            Set(ByVal value As String)
                strModule = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property Path() As String
            Get
                Return strReportLoc
            End Get

            Set(ByVal value As String)
                strReportLoc = value
                Me.CheckDirty()
            End Set
        End Property

        'Public Property ReportParams() As Muster.BusinessLogic.pReportParams
        '    Get
        '        Return oReportParams.ReportParams
        '    End Get

        '    Set(ByVal value As Muster.BusinessLogic.pReportParams)
        '        oReportParams = value
        '    End Set
        'End Property

        'Public Property ReportParamName() As String
        '    Get
        '        Return oReportParams.ParamName
        '    End Get

        '    Set(ByVal value As String)
        '        oReportParams.ParamName = value
        '    End Set
        'End Property

        'Public Property ReportParamDescription() As String
        '    Get
        '        Return oReportParams.ParamDescription
        '    End Get

        '    Set(ByVal value As String)
        '        oReportParams.ParamDescription = value
        '    End Set
        'End Property

        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get

            Set(ByVal value As Boolean)
                bolDeleted = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property Active() As Boolean
            Get
                Return bolActive
            End Get

            Set(ByVal value As Boolean)
                bolActive = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property IsDirty() As Boolean
            Get
                If bolIsDirty Then
                    Return bolIsDirty
                Else
                    For Each groupReportRegInfo As MUSTER.Info.ReportGroupRelationInfo In colReportGroupRelation.Values
                        If groupReportRegInfo.IsDirty Then
                            Return True
                        End If
                    Next
                End If
                Return False
            End Get

            Set(ByVal value As Boolean)
                bolIsDirty = value
            End Set
        End Property

        Public Property AgeThreshold() As Int16
            Get
                Return nAgeThreshold
            End Get

            Set(ByVal value As Int16)
                nAgeThreshold = Int16.Parse(value)
            End Set
        End Property

        Public Property ReportGroupRelationCollection() As MUSTER.Info.ReportGroupRelationsCollection
            Get
                Return colReportGroupRelation
            End Get
            Set(ByVal Value As MUSTER.Info.ReportGroupRelationsCollection)
                colReportGroupRelation = Value
            End Set
        End Property

        Public ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get
        End Property

#Region "iAccessors"
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
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class

End Namespace

