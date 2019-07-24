Namespace MUSTER.Info
    <Serializable()> _
    Public Class AdvancedSearchInfo
#Region "Private member variables"
        Private strBrand As String
        Private strContact As String
        Private StrFacilityAddress As String
        Private strFacilityCity As String
        Private strFacilityCounty As String
        Private strFacilityName As String
        Private strLicensedCompany As String
        Private strLicensedContractors As String
        Private strOwnerAddress As String
        Private strOwnerCity As String
        Private nOwnerId As Integer
        Private strOwnerName As String
        Private strProjectMgr As String
        Private strContactName As String

#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub
        Sub New(ByVal Brand As String, _
                ByVal Contact As String, _
                ByVal FacilityAddress As String, _
                ByVal FacilityCity As String, _
                ByVal FacilityCounty As String, _
                ByVal FacilityName As String, _
                ByVal LicensedCompany As String, _
                ByVal LicensedContractors As String, _
                ByVal OwnerAddress As String, _
                ByVal OwnerCity As String, _
                ByVal OwnerId As Integer, _
                ByVal OwnerName As String, _
                ByVal ProjectMgr As String, _
                ByVal ContactName As String)

            strBrand = Brand
            strContact = Contact
            StrFacilityAddress = FacilityAddress
            strFacilityCity = FacilityCity
            strFacilityCounty = FacilityCounty
            strFacilityName = FacilityName
            strLicensedCompany = LicensedCompany
            strLicensedContractors = LicensedContractors
            strOwnerAddress = OwnerAddress
            strOwnerCity = OwnerCity
            nOwnerId = OwnerId
            strOwnerName = OwnerName
            strProjectMgr = ProjectMgr
            strContactName = ContactName

        End Sub
#End Region
#Region "Private Operations"
        Private Sub Init()
            strBrand = ""
            strContact = ""
            StrFacilityAddress = ""
            strFacilityCity = ""
            strFacilityCounty = ""
            strFacilityName = ""
            strLicensedCompany = ""
            strLicensedContractors = ""
            strOwnerAddress = ""
            strOwnerCity = ""
            nOwnerId = 0
            strOwnerName = ""
            strProjectMgr = ""
            strContactName = ""
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property Brand() As String
            Get
                Return strBrand
            End Get
            Set(ByVal Value As String)
                strBrand = Value
            End Set
        End Property
        Public Property Contact() As String
            Get
                Return strContact
            End Get
            Set(ByVal Value As String)
                strContact = Value
            End Set
        End Property
        Public Property FacilityAddress() As String
            Get
                Return StrFacilityAddress
            End Get

            Set(ByVal Value As String)
                StrFacilityAddress = Value
            End Set
        End Property
        Public Property FacilityCity() As String
            Get
                Return strFacilityCity
            End Get

            Set(ByVal Value As String)
                strFacilityCity = Value
            End Set
        End Property
        Public Property FacilityCounty() As String
            Get
                Return strFacilityCounty
            End Get

            Set(ByVal Value As String)
                strFacilityCounty = Value
            End Set
        End Property
        Public Property FacilityName() As String
            Get
                Return strFacilityName
            End Get

            Set(ByVal Value As String)
                strFacilityName = Value
            End Set
        End Property
        Public Property LicensedCompany() As String
            Get
                Return strLicensedCompany
            End Get

            Set(ByVal Value As String)
                strLicensedCompany = Value
            End Set
        End Property
        Public Property LicensedContractors() As String
            Get
                Return strLicensedContractors
            End Get

            Set(ByVal Value As String)
                strLicensedContractors = Value
            End Set
        End Property
        Public Property OwnerAddress() As String
            Get
                Return strOwnerAddress
            End Get

            Set(ByVal Value As String)
                strOwnerAddress = Value
            End Set
        End Property
        Public Property OwnerCity() As String
            Get
                Return strOwnerCity
            End Get

            Set(ByVal Value As String)
                strOwnerCity = Value
            End Set
        End Property
        Public Property OwnerId() As Integer
            Get
                Return nOwnerId
            End Get
            Set(ByVal Value As Integer)
                nOwnerId = Value
            End Set
        End Property
        Public Property OwnerName() As String
            Get
                Return strOwnerName
            End Get

            Set(ByVal Value As String)
                strOwnerName = Value
            End Set
        End Property
        Public Property ProjectMgr() As String
            Get
                Return strProjectMgr
            End Get

            Set(ByVal Value As String)
                strProjectMgr = Value
            End Set
        End Property
        Public Property ContactName() As String
            Get
                Return strContactName
            End Get

            Set(ByVal Value As String)
                strContactName = Value
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