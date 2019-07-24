'-------------------------------------------------------------------------------
' MUSTER.Info.CompanyLicenseeInfo
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      MKK/RAF     05/26/2005  Original class definition
'
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class CompanyLicenseeInfo
#Region "Public Events"
        Public Event CompanyLicenseeInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nID As Int64
        Private nCompanyID As Integer
        Private nLicenseeID As Integer
        Private nCompanyAddressID As Integer
        Private bolDeleted As Boolean
        Private strCreatedBy As String
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String
        Private dtModifiedOn As DateTime

        Private onID As Int64
        Private onCompanyID As Integer
        Private onLicenseeID As Integer
        Private onCompanyAddressID As Integer
        Private ostrCreatedBy As String
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String
        Private odtModifiedOn As DateTime
        Private obolDeleted As Boolean

        Private bolShowDeleted As Boolean = False

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub

        Sub New(ByVal ID As Integer, _
     ByVal CompanyID As Integer, _
     ByVal LicenseeID As Integer, _
     ByVal CompanyAddressID As Integer, _
     ByVal Deleted As Boolean, _
     ByVal CreatedBy As String, _
     ByVal CreatedOn As Date, _
     ByVal ModifiedBy As String, _
     ByVal LastEdited As Date)
            onID = ID
            onCompanyID = CompanyID
            onLicenseeID = LicenseeID
            onCompanyAddressID = CompanyAddressID
            obolDeleted = Deleted
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = LastEdited
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nID >= 0 Then
                nID = onID
            End If
            nCompanyID = onCompanyID
            nLicenseeID = onLicenseeID
            nCompanyAddressID = onCompanyAddressID
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
            RaiseEvent CompanyLicenseeInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onID = nID
            obolDeleted = bolDeleted
            onCompanyID = nCompanyID
            onLicenseeID = nLicenseeID
            onCompanyAddressID = nCompanyAddressID
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        '********************************************************
        '
        ' Update eval in CheckDirty with other current and prior
        '    state variables.
        '
        '********************************************************
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty

            bolIsDirty = (nID <> onID) Or _
            (nCompanyID <> onCompanyID) Or _
            (nLicenseeID <> onLicenseeID) Or _
            (nCompanyAddressID <> onCompanyAddressID) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn) Or _
            (bolDeleted <> obolDeleted)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent CompanyLicenseeInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            nID = 0
            nCompanyID = 0
            nCompanyAddressID = 0
            nLicenseeID = 0
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
        Public Property CompanyID() As Integer
            Get
                Return nCompanyID
            End Get
            Set(ByVal Value As Integer)
                nCompanyID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LicenseeID() As Integer
            Get
                Return nLicenseeID
            End Get
            Set(ByVal Value As Integer)
                nLicenseeID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ComLicAddressID() As Integer
            Get
                Return nCompanyAddressID
            End Get
            Set(ByVal Value As Integer)
                nCompanyAddressID = Value
                Me.CheckDirty()
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
        Public Property CreatedOn() As DateTime
            Get
                Return dtCreatedOn
            End Get
            Set(ByVal Value As DateTime)
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
        Public Property ModifiedOn() As DateTime
            Get
                Return dtModifiedOn
            End Get
            Set(ByVal Value As DateTime)
                dtModifiedOn = Value
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
