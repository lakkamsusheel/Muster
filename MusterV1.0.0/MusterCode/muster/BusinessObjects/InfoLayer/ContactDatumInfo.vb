'-------------------------------------------------------------------------------
' MUSTER.Info.ContactDatumInfo
'   Provides the container to persist MUSTER Template state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date            Description
'  1.0       KKM        03/28/2005      Original class definition
'  1.1       MR         04/29/05        Added Ext1 and Ext2 Attributes


Namespace MUSTER.Info
    <Serializable()> _
Public Class ContactDatumInfo
#Region "Public Events"
        Public Event ContactDatumInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nContactID As Integer
        Private strModifiedBy As String = String.Empty
        Private strCreatedBy As String = String.Empty
        Private bolIsDirty As Boolean
        Private bolIsAddressDirty As Boolean
        Private bolIsOthersDirty As Boolean
        Private bolDeleted As Boolean
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private dtDataAge As Date
        Private strCity As String
        Private strState As String
        Private strZipCode As String
        Private bolIsPerson As Boolean
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private strTitle As String
        Private strPrefix As String
        Private strFirstName As String
        Private strMiddleName As String
        Private strLastName As String
        Private strSuffix As String
        Private strAddressLine1 As String
        Private strAddressLine2 As String
        Private strPhone1 As String
        Private strPhone2 As String
        Private strExt1 As String
        Private strExt2 As String
        Private strFax As String
        Private strCell As String
        Private strPublicEmail As String
        Private strPrivateEmail As String
        Private strFipsCode As String
        Private nOrgEntityCode As Integer
        Private strCompanyName As String
        Private nOrgCode As Integer
        Private strVendorNumber As String

        Private onContactID As Integer
        Private ostrModifiedBy As String
        Private ostrCreatedBy As String
        Private obolDeleted As Boolean
        Private odtCreatedOn As Date
        Private odtDataAge As Date
        Private ostrCity As String
        Private ostrState As String
        Private ostrZipCode As String
        Private odtModifiedOn As Date
        Private ostrTitle As String
        Private ostrPrefix As String
        Private ostrFirstName As String
        Private ostrMiddleName As String
        Private ostrLastName As String
        Private ostrSuffix As String
        Private ostrAddressLine1 As String
        Private ostrAddressLine2 As String
        Private ostrPhone1 As String
        Private ostrPhone2 As String
        Private ostrExt1 As String
        Private ostrExt2 As String
        Private ostrFax As String
        Private ostrCell As String
        Private ostrPublicEmail As String
        Private ostrPrivateEmail As String
        Private ostrFipsCode As String
        Private onOrgEntityCode As Integer
        Private ostrCompanyName As String
        Private onOrgCode As Integer
        Private ostrVendorNumber As String

        Private nAgeThreshold As Int16 = 5
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
            dtDataAge = Now()
        End Sub

        Sub New(ByVal ContactID As Integer, _
            ByVal IsPerson As Boolean, _
            ByVal Org_Entity_Code As Integer, _
            ByVal Company_Name As String, _
            ByVal title As String, _
            ByVal prefix As String, _
            ByVal firstName As String, _
            ByVal middleName As String, _
            ByVal lastName As String, _
            ByVal suffix As String, _
            ByVal Address1 As String, _
            ByVal Address2 As String, _
            ByVal city As String, _
            ByVal state As String, _
            ByVal zipcode As String, _
            ByVal fipsCode As String, _
            ByVal phone1 As String, _
            ByVal phone2 As String, _
            ByVal Ext1 As String, _
            ByVal Ext2 As String, _
            ByVal Fax As String, _
            ByVal Cell As String, _
            ByVal Email1 As String, _
            ByVal Email2 As String, _
            ByVal strVendorNo As String, _
            ByVal CreatedBy As String, _
            ByVal dateCreated As Date, _
            ByVal LAST_EDITED_BY As String, _
            ByVal DATE_LAST_EDITED As Date, _
            ByVal deleted As Boolean)
            onContactID = ContactID
            bolIsPerson = IsPerson
            onOrgEntityCode = Org_Entity_Code
            ostrCompanyName = Company_Name

            ostrCity = city
            ostrState = state
            ostrZipCode = zipcode
            obolDeleted = deleted

            ostrTitle = title
            ostrPrefix = prefix
            ostrFirstName = firstName
            ostrMiddleName = middleName
            ostrLastName = lastName
            ostrSuffix = suffix
            ostrAddressLine1 = Address1
            ostrAddressLine2 = Address2
            ostrPhone1 = phone1
            ostrPhone2 = phone2
            ostrExt1 = Ext1
            ostrExt2 = Ext2
            ostrFax = Fax
            ostrCell = Cell
            ostrPublicEmail = Email1
            ostrPrivateEmail = Email2
            ostrFipsCode = fipsCode
            onOrgCode = Org_Entity_Code

            odtCreatedOn = dateCreated
            odtModifiedOn = DATE_LAST_EDITED
            ostrCreatedBy = CreatedBy
            ostrModifiedBy = LAST_EDITED_BY
            ostrVendorNumber = strVendorNo
            dtDataAge = Now()
            Me.Reset()
        End Sub
        Sub New(ByVal drContactManagement As DataRow)
            Try
                dtDataAge = Now()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nContactID >= 0 Then
                nContactID = onContactID
            End If

            strModifiedBy = ostrModifiedBy
            strCreatedBy = ostrCreatedBy
            nOrgCode = onOrgCode
            strCity = ostrCity
            strState = ostrState
            strZipCode = ostrZipCode
            bolDeleted = obolDeleted
            dtCreatedOn = odtCreatedOn
            dtModifiedOn = odtModifiedOn

            strTitle = ostrTitle
            strPrefix = ostrPrefix
            strFirstName = ostrFirstName
            strMiddleName = ostrMiddleName
            strLastName = ostrLastName
            strSuffix = ostrSuffix
            strCompanyName = ostrCompanyName
            strAddressLine1 = ostrAddressLine1
            strAddressLine2 = ostrAddressLine2
            strPhone1 = ostrPhone1
            strPhone2 = ostrPhone2
            strExt1 = ostrExt1
            strExt2 = ostrExt2
            strFax = ostrFax
            strCell = ostrCell
            strPublicEmail = ostrPublicEmail
            strPrivateEmail = ostrPrivateEmail
            strFipsCode = ostrFipsCode
            strVendorNumber = ostrVendorNumber
            bolIsDirty = False
            bolIsAddressDirty = False
            bolIsOthersDirty = False
            RaiseEvent ContactDatumInfoChanged(bolIsDirty)

        End Sub
        Public Sub Archive()
            onContactID = nContactID
            ostrModifiedBy = strModifiedBy
            ostrCreatedBy = strCreatedBy
            ostrCity = strCity
            ostrState = strState
            ostrZipCode = strZipCode
            obolDeleted = bolDeleted
            odtCreatedOn = dtCreatedOn
            odtModifiedOn = dtModifiedOn

            ostrTitle = strTitle
            ostrPrefix = strPrefix
            ostrFirstName = strFirstName
            ostrMiddleName = strMiddleName
            ostrLastName = strLastName
            ostrSuffix = strSuffix
            ostrAddressLine1 = strAddressLine1
            ostrAddressLine2 = strAddressLine2
            ostrPhone1 = strPhone1
            ostrPhone2 = strPhone2
            ostrExt1 = strExt1
            ostrExt2 = strExt2
            ostrFax = strFax
            ostrCell = strCell
            ostrPublicEmail = strPublicEmail
            ostrPrivateEmail = strPrivateEmail
            ostrFipsCode = strFipsCode
            ostrVendorNumber = strVendorNumber
            bolIsDirty = False
            bolIsAddressDirty = False
            bolIsOthersDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty
            bolIsDirty = (nContactID <> onContactID) Or _
            (nOrgEntityCode <> onOrgEntityCode) Or _
            (strCompanyName <> ostrCompanyName) Or _
            (strTitle <> ostrTitle) Or _
            (strPrefix <> ostrPrefix) Or _
            (strFirstName <> ostrFirstName) Or _
            (strMiddleName <> ostrMiddleName) Or _
            (strLastName <> ostrLastName) Or _
            (strSuffix <> ostrSuffix) Or _
            (strAddressLine1 <> ostrAddressLine1) Or _
            (strAddressLine2 <> ostrAddressLine2) Or _
            (ostrCity <> strCity) Or _
            (ostrState <> strState) Or _
            (ostrZipCode <> strZipCode) Or _
            (ostrFipsCode <> strFipsCode) Or _
            (ostrPhone1 <> strPhone1) Or _
            (ostrPhone2 <> strPhone2) Or _
            (ostrExt1 <> strExt1) Or _
            (ostrExt2 <> strExt2) Or _
            (ostrFax <> strFax) Or _
            (ostrCell <> strCell) Or _
            (ostrPublicEmail <> strPublicEmail) Or _
            (ostrPrivateEmail <> strPrivateEmail) Or _
            (obolDeleted <> bolDeleted) Or _
            (ostrModifiedBy <> strModifiedBy) Or _
            (ostrCreatedBy <> strCreatedBy) Or _
            (odtCreatedOn <> dtCreatedOn) Or _
            (odtModifiedOn <> dtModifiedOn) Or _
            (ostrVendorNumber <> strVendorNumber)
            If bolIsDirty <> obolIsDirty Then
                RaiseEvent ContactDatumInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub CheckAddressDirty()
            bolIsAddressDirty = (strAddressLine1 <> ostrAddressLine1) Or _
            (strAddressLine2 <> ostrAddressLine2) Or _
            (ostrCity <> strCity) Or _
            (ostrState <> strState) Or _
            (ostrZipCode <> strZipCode) Or _
            (ostrFipsCode <> strFipsCode)
        End Sub
        Private Sub CheckOthersDirty()
            bolIsOthersDirty = (nContactID <> onContactID) Or _
            (nOrgEntityCode <> onOrgEntityCode) Or _
            (strCompanyName <> ostrCompanyName) Or _
            (strTitle <> ostrTitle) Or _
            (strPrefix <> ostrPrefix) Or _
            (strFirstName <> ostrFirstName) Or _
            (strMiddleName <> ostrMiddleName) Or _
            (strLastName <> ostrLastName) Or _
            (strSuffix <> ostrSuffix) Or _
            (ostrPhone1 <> strPhone1) Or _
            (ostrPhone2 <> strPhone2) Or _
            (ostrExt1 <> strExt1) Or _
            (ostrExt2 <> strExt2) Or _
            (ostrFax <> strFax) Or _
            (ostrCell <> strCell) Or _
            (ostrPublicEmail <> strPublicEmail) Or _
            (ostrPrivateEmail <> strPrivateEmail) Or _
            (obolDeleted <> bolDeleted) Or _
            (ostrModifiedBy <> strModifiedBy) Or _
            (ostrCreatedBy <> strCreatedBy) Or _
            (odtCreatedOn <> dtCreatedOn) Or _
            (odtModifiedOn <> dtModifiedOn) Or _
            (ostrVendorNumber <> strVendorNumber)
        End Sub

        Private Sub Init()
            ostrModifiedBy = String.Empty
            ostrCreatedBy = String.Empty
            bolDeleted = False
            odtCreatedOn = DateTime.Now.ToShortDateString()
            onContactID = 0
            ostrCity = String.Empty
            ostrState = String.Empty
            ostrZipCode = String.Empty
            bolIsPerson = True
            odtModifiedOn = DateTime.Now.ToShortDateString()
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property companyName() As String
            Get
                Return strCompanyName
            End Get
            Set(ByVal Value As String)
                strCompanyName = Value
            End Set
        End Property
        Public Property orgCode() As Integer
            Get
                Return nOrgCode
            End Get
            Set(ByVal Value As Integer)
                nOrgCode = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get
            Set(ByVal Value As Boolean)
                bolIsDirty = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IsAddressDirty() As Boolean
            Get
                Return bolIsAddressDirty
            End Get
            Set(ByVal Value As Boolean)
                bolIsAddressDirty = Value
                Me.CheckAddressDirty()
            End Set
        End Property
        Public Property IsOthersDirty() As Boolean
            Get
                Return bolIsOthersDirty
            End Get
            Set(ByVal Value As Boolean)
                bolIsOthersDirty = Value
                Me.CheckOthersDirty()
            End Set
        End Property
        Public Property deleted() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolDeleted = Value
                Me.CheckDirty()
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
        Public Property ID() As Integer
            Get
                Return nContactID
            End Get
            Set(ByVal Value As Integer)
                nContactID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get
            Set(ByVal Value As Boolean)

            End Set
        End Property

        Public Property City() As String
            Get
                Return strCity
            End Get
            Set(ByVal Value As String)
                strCity = Value
                Me.CheckAddressDirty()
                Me.CheckDirty()
            End Set
        End Property
        Public Property State() As String
            Get
                Return strState
            End Get
            Set(ByVal Value As String)
                strState = Value
                Me.CheckAddressDirty()
                Me.CheckDirty()
            End Set
        End Property
        Public Property ZipCode() As String
            Get
                Return strZipCode
            End Get
            Set(ByVal Value As String)
                strZipCode = Value
                Me.CheckAddressDirty()
                Me.CheckDirty()
            End Set
        End Property
        Public Property IsPerson() As Boolean
            Get
                Return bolIsPerson
            End Get
            Set(ByVal Value As Boolean)
                bolIsPerson = Value
                Me.CheckDirty()
                Me.CheckOthersDirty()
            End Set
        End Property
        Public Property Title() As String
            Get
                Return strTitle
            End Get
            Set(ByVal Value As String)
                strTitle = Value
                Me.CheckOthersDirty()
            End Set
        End Property
        Public Property Prefix() As String
            Get
                Return strPrefix
            End Get
            Set(ByVal Value As String)
                strPrefix = Value
                Me.CheckOthersDirty()
            End Set
        End Property
        Public Property FirstName() As String
            Get
                Return strFirstName
            End Get
            Set(ByVal Value As String)
                strFirstName = Value
                Me.CheckOthersDirty()
            End Set
        End Property
        Public Property MiddleName() As String
            Get
                Return strMiddleName
            End Get
            Set(ByVal Value As String)
                strMiddleName = Value
                Me.CheckOthersDirty()
            End Set
        End Property
        Public Property LastName() As String
            Get
                Return strLastName

            End Get
            Set(ByVal Value As String)
                strLastName = Value
                Me.CheckOthersDirty()
            End Set
        End Property
        Public Property suffix() As String
            Get
                Return strSuffix
            End Get
            Set(ByVal Value As String)
                strSuffix = Value
                Me.CheckOthersDirty()
            End Set
        End Property
        Public ReadOnly Property fullName() As String
            Get
                Return strTitle + " " + strPrefix + " " + strFirstName + " " + strMiddleName + " " + strLastName + " " + strSuffix
            End Get
        End Property
        Public ReadOnly Property SortName() As String
            Get
                Return strLastName + ", " + strFirstName + ", " + strMiddleName
            End Get
        End Property
        Public ReadOnly Property PlainName() As String
            Get
                Return strFirstName + " " + strMiddleName + " " + strLastName + " " + strSuffix
            End Get
        End Property
        Public Property AddressLine1() As String
            Get
                Return strAddressLine1
            End Get
            Set(ByVal Value As String)
                strAddressLine1 = Value
                Me.CheckAddressDirty()
                Me.CheckDirty()
            End Set
        End Property
        Public Property AddressLine2() As String
            Get
                Return strAddressLine2
            End Get
            Set(ByVal Value As String)
                strAddressLine2 = Value
                Me.CheckAddressDirty()
                Me.CheckDirty()
            End Set
        End Property
        Public Property Phone1() As String
            Get
                Return strPhone1
            End Get
            Set(ByVal Value As String)
                strPhone1 = Value
                Me.CheckOthersDirty()
            End Set
        End Property
        Public Property Phone2() As String
            Get
                Return strPhone2
            End Get
            Set(ByVal Value As String)
                strPhone2 = Value
                Me.CheckOthersDirty()
            End Set
        End Property
        Public Property Ext1() As String
            Get
                Return strExt1
            End Get
            Set(ByVal Value As String)
                strExt1 = Value
                Me.CheckOthersDirty()
            End Set
        End Property
        Public Property Ext2() As String
            Get
                Return strExt2
            End Get
            Set(ByVal Value As String)
                strExt2 = Value
                Me.CheckOthersDirty()
            End Set
        End Property
        Public Property Fax() As String
            Get
                Return strFax
            End Get
            Set(ByVal Value As String)
                strFax = Value
                Me.CheckOthersDirty()
            End Set
        End Property
        Public Property Cell() As String
            Get
                Return strCell
            End Get
            Set(ByVal Value As String)
                strCell = Value
                Me.CheckOthersDirty()
            End Set
        End Property
        Public Property publicEmail() As String
            Get
                Return strPublicEmail
            End Get
            Set(ByVal Value As String)
                strPublicEmail = Value
                Me.CheckOthersDirty()
            End Set
        End Property
        Public Property privateEmail() As String
            Get
                Return strPrivateEmail
            End Get
            Set(ByVal Value As String)
                strPrivateEmail = Value
                Me.CheckOthersDirty()
            End Set
        End Property
        Public Property FipsCode() As String
            Get
                Return strFipsCode
            End Get
            Set(ByVal Value As String)
                strFipsCode = Value
            End Set
        End Property
        Public Property VendorNumber() As String
            Get
                Return strVendorNumber
            End Get
            Set(ByVal Value As String)
                strVendorNumber = Value
                Me.CheckOthersDirty()
            End Set
        End Property
#Region "iAccessors"
        Public Property modifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
            Set(ByVal Value As Date)
                dtModifiedOn = Value
            End Set
        End Property

        Public Property modifiedBy() As String
            Get
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
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

#End Region

#End Region

#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace
