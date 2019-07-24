'-------------------------------------------------------------------------------
' MUSTER.Info.CompanyLicenseeInfo
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      MKK/RAF     05/26/2005  Original class definition
'  1.1        MR        06/04/2005  Added Functions and New Attributes
'
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class ManagerFacRelationInfo
#Region "Public Events"
        Public Event MgrFacRelationInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nID As Int64
        Private nManagerID As Integer
        Private nFacilityID As Integer
        Private nRelationID As Integer
        Private bolDeleted As Boolean
        Private strRelationDesc = String.Empty

        Private onID As Int64
        Private onManagerID As Integer

        Private onFacilityID As Integer
        Private onRelationID As Integer
        Private obolDeleted As Boolean
        Private ostrRelationDesc As String


        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions



#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub
        Sub New(ByVal ID As Integer, _
        ByVal ManagerID As Integer, _
        ByVal FacilityID As Integer, _
        ByVal RelationID As Integer, _
        ByVal RelationDesc As String, _
        ByVal Deleted As Boolean)
            onID = ID
            onManagerID = ManagerID
            onFacilityID = FacilityID
            onRelationID = RelationID

            ostrRelationDesc = RelationDesc


            Me.Reset()

        End Sub
#End Region

#Region "Exposed Operations"
        Public Sub Reset()
            If nID >= 0 Then
                nID = onID
            End If

            nManagerID = onManagerID

            nFacilityID = onFacilityID
            nRelationID = onRelationID

            strRelationDesc = ostrRelationDesc
            bolIsDirty = False
            RaiseEvent MgrFacRelationInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onID = nID
            onManagerID = nManagerID
            onFacilityID = nFacilityID
            onRelationID = nRelationID
            ostrRelationDesc = strRelationDesc

            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            '(nID <> onID) Or _
            '(nLicenseeID <> onLicenseeID) Or _
            Dim obolIsDirty As Boolean = bolIsDirty

            bolIsDirty = (nFacilityID <> onFacilityID) Or _
            (nRelationID <> onRelationID) Or _
            (strRelationDesc <> ostrRelationDesc) Or _
            (bolDeleted <> obolDeleted)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent MgrFacRelationInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            nID = 0
            nManagerID = 0
            nFacilityID = 0
            nRelationID = 0
            bolDeleted = False
            strRelationDesc = String.Empty

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
        Public Property ManagerID() As Int64
            Get
                Return nManagerID
            End Get
            Set(ByVal Value As Int64)
                nManagerID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FacilityID() As Integer
            Get
                Return nFacilityID
            End Get
            Set(ByVal Value As Integer)
                nFacilityID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property RelationID() As Integer
            Get
                Return nRelationID
            End Get
            Set(ByVal Value As Integer)
                nRelationID = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property RelationDesc() As String
            Get
                Return strRelationDesc
            End Get
            Set(ByVal Value As String)
                strRelationDesc = Value
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
