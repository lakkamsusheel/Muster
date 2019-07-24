'-------------------------------------------------------------------------------
' MUSTER.Info.CitationPenaltyInfo
'   Provides the container to persist MUSTER CitationPenalty state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      MKK/RAF     06/27/2005  Initial Development
'
' Function          Description
' New()             Instantiates an empty CitationPenaltyInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated CitationPenaltyInfo object
' New(dr)           Instantiates a populated CitationPenaltyInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
' NOTE: This file to be used as CitationPenalty to build other objects.
'       Replace keyword "CitationPenalty" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class CitationPenaltyInfo
#Region "Public Events"
        Public Event CitationPenaltyInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nID As Int64
        Private strStateCitation As String
        Private strFederalCitation As String
        Private strSection As String
        Private strDescription As String
        Private strCategory As String
        Private nSmall As Integer
        Private nMedium As Integer
        Private nLarge As Integer
        Private strCorrectiveAction As String
        Private strEPA As String
        Private bolDeleted As Boolean
        Private strCreatedBy As String
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String
        Private dtModifiedOn As DateTime

        Private onID As Int64
        Private ostrStateCitation As String
        Private ostrFederalCitation As String
        Private ostrSection As String
        Private ostrDescription As String
        Private ostrCategory As String
        Private onSmall As Integer
        Private onMedium As Integer
        Private onLarge As Integer
        Private ostrCorrectiveAction As String
        Private ostrEPA As String
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String
        Private odtModifiedOn As DateTime

        Private bolShowDeleted As Boolean = False

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub
        Sub New(ByVal ID As Int64, _
         ByVal StateCitation As String, _
            ByVal FederalCitation As String, _
            ByVal Section As String, _
            ByVal Description As String, _
            ByVal Category As String, _
            ByVal Small As Integer, _
            ByVal Medium As Integer, _
            ByVal Large As Integer, _
            ByVal CorrectiveAction As String, _
            ByVal EPA As String, _
            ByVal CreatedBy As String, _
            ByVal CreatedOn As Date, _
            ByVal ModifiedBy As String, _
            ByVal LastEdited As Date, _
            ByVal Deleted As Boolean)
            onID = ID
            ostrStateCitation = StateCitation
            ostrFederalCitation = FederalCitation
            ostrSection = Section
            ostrDescription = Description
            ostrCategory = Category
            onSmall = Small
            onLarge = Large
            ostrCorrectiveAction = CorrectiveAction
            ostrEPA = EPA
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = LastEdited
            obolDeleted = Deleted
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            nID = onID
            strStateCitation = ostrStateCitation
            strFederalCitation = ostrFederalCitation
            strSection = ostrSection
            strDescription = ostrDescription
            strCategory = ostrCategory
            nSmall = onSmall
            nMedium = onMedium
            nLarge = onLarge
            strCorrectiveAction = ostrCorrectiveAction
            strEPA = ostrEPA
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
            RaiseEvent CitationPenaltyInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onID = nID
            ostrStateCitation = strStateCitation
            ostrFederalCitation = strFederalCitation
            ostrSection = strSection
            ostrDescription = strDescription
            ostrCategory = strCategory
            onSmall = nSmall
            onMedium = nMedium
            onLarge = nLarge
            ostrCorrectiveAction = strCorrectiveAction
            ostrEPA = strEPA
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

            bolIsDirty = (nID <> onID) Or _
            (strStateCitation <> ostrStateCitation) Or _
            (strFederalCitation <> ostrFederalCitation) Or _
            (strSection <> ostrSection) Or _
            (strDescription <> ostrDescription) Or _
            (strCategory <> ostrCategory) Or _
            (nSmall <> onSmall) Or _
            (nMedium <> onMedium) Or _
            (nLarge <> onLarge) Or _
            (strCorrectiveAction <> ostrCorrectiveAction) Or _
            (strEPA <> ostrEPA) Or _
            (bolDeleted <> obolDeleted) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent CitationPenaltyInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onID = 0
            strStateCitation = String.Empty
            strFederalCitation = String.Empty
            strSection = String.Empty
            strDescription = String.Empty
            strCategory = String.Empty
            nSmall = 0
            nMedium = 0
            nLarge = 0
            strCorrectiveAction = String.Empty
            strEPA = String.Empty
            bolDeleted = False
            strCreatedBy = String.Empty
            dtCreatedOn = System.DateTime.Now
            strModifiedBy = String.Empty
            dtModifiedOn = System.DateTime.Now
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int64
            Get
                Return nID
            End Get
            Set(ByVal Value As Int64)
                nID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property StateCitation() As String
            Get
                Return strStateCitation
            End Get
            Set(ByVal Value As String)
                strStateCitation = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Small() As Integer
            Get
                Return nSmall
            End Get
            Set(ByVal Value As Integer)
                nSmall = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Section() As String
            Get
                Return strSection
            End Get
            Set(ByVal Value As String)
                strSection = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Medium() As Integer
            Get
                Return nMedium
            End Get
            Set(ByVal Value As Integer)
                nMedium = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Large() As Integer
            Get
                Return nLarge
            End Get
            Set(ByVal Value As Integer)
                nLarge = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FederalCitation() As String
            Get
                Return strFederalCitation
            End Get
            Set(ByVal Value As String)
                strFederalCitation = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property EPA() As String
            Get
                Return strEPA
            End Get
            Set(ByVal Value As String)
                strEPA = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Description() As String
            Get
                Return strDescription
            End Get
            Set(ByVal Value As String)
                strDescription = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property DATE_LAST_EDITED() As Date
            Get
                Return dtModifiedOn
            End Get
            Set(ByVal Value As Date)
                dtModifiedOn = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LAST_EDITED_BY() As String
            Get
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property DATE_CREATED() As Date
            Get
                Return dtCreatedOn
            End Get
            Set(ByVal Value As Date)
                dtCreatedOn = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CREATED_BY() As String
            Get
                Return strCreatedBy
            End Get
            Set(ByVal Value As String)
                strCreatedBy = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CorrectiveAction() As String
            Get
                Return strCorrectiveAction
            End Get
            Set(ByVal Value As String)
                strCorrectiveAction = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Category() As String
            Get
                Return strCategory
            End Get
            Set(ByVal Value As String)
                strCategory = Value
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
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace
