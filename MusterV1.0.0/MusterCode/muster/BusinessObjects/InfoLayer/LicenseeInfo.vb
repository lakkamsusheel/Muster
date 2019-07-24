'-------------------------------------------------------------------------------
' MUSTER.Info.LicenseeInfo
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MKK/RAF    05/24/05  Original class definition
'
'
'-------------------------------------------------------------------------------
Option Strict On
Option Explicit On 
Namespace MUSTER.Info
    <Serializable()> _
    Public Class LicenseeInfo
#Region "Public Events"
        Public Event LicenseeInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nID As Int64
        Private strTitle As String
        Private strFirstName As String
        Private strMiddleName As String
        Private strLastName As String
        Private strSuffix As String
        Private strLicenseNumberPrefix As String
        Private nLicenseNumber As String
        Private strEmailAddress As String
        Private nAssociatedCompanyID As Integer
        Private strHireStatus As String
        Private bolEmployeeLetter As Boolean
        Private nStatus As Integer
        Private nCMStatus As Integer
        Private strStatusDesc As String
        Private strCMStatusDesc As String
        Private bolOverrideExpire As Boolean
        Private nCertType As Integer
        Private nCMCertType As Integer
        Private strCertTypeDesc As String
        Private strCMCertTypeDesc As String
        Private dtAppRecvdDate As DateTime
        Private dtOriginalIssuedDate As DateTime
        Private dtIssuedDate As DateTime
        Private dtExtensionDeadlineDate As DateTime
        Private dtLicenseExpirationDate As String
        Private dtExceptionGrantedDate As DateTime
        Private bolComplianceManager As Boolean
        Private bolIsLicensee As Boolean
        Private dtInitCertDate As String
        Private nInitCertBy As Integer
        Private strInitCertByDesc As String
        Private dtRetrainDate1 As String
        Private dtRetrainDate2 As String
        Private dtRetrainDate3 As String
        Private dtRevokeDate As String
        Private dtRetrainReqDate As String
        Private bolDeleted As Boolean
        Private strCreatedBy As String
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String
        Private dtModifiedOn As DateTime

        Private onID As Int64
        Private ostrTitle As String
        Private ostrFirstName As String
        Private ostrMiddleName As String
        Private ostrLastName As String
        Private ostrSuffix As String
        Private ostrLicenseNumberPrefix As String
        Private onLicenseNumber As String
        Private ostrEmailAddress As String
        Private onAssociatedCompanyID As Integer
        Private ostrHireStatus As String
        Private obolEmployeeLetter As Boolean
        Private onStatus As Integer
        Private onCMStatus As Integer
        Private ostrStatusDesc As String
        Private ostrCMStatusDesc As String
        Private obolOverrideExpire As Boolean
        Private onCertType As Integer
        Private onCMCertType As Integer
        Private ostrCertTypeDesc As String
        Private ostrCMCertTypeDesc As String
        Private odtAppRecvdDate As DateTime
        Private odtOriginalIssuedDate As DateTime
        Private odtIssuedDate As DateTime
        Private odtExtensionDeadlineDate As DateTime
        Private obolComplianceManager As Boolean
        Private obolIsLicensee As Boolean
        Private odtInitCertDate As String
        Private onInitCertBy As Integer
        Private ostrInitCertByDesc As String
        Private odtRetrainDate1 As String
        Private odtRetrainDate2 As String
        Private odtRetrainDate3 As String
        Private odtRevokeDate As String
        Private odtRetrainReqDate As String
        Private odtLicenseExpirationDate As String
        Private odtExceptionGrantedDate As DateTime
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
            ByVal Title As String, _
            ByVal FirstName As String, _
            ByVal MiddleName As String, _
            ByVal LastName As String, _
            ByVal Suffix As String, _
            ByVal LicenseNumberPrefix As String, _
            ByVal LicenseNumber As Integer, _
            ByVal EmailAddress As String, _
            ByVal AssociatedCompanyID As Integer, _
            ByVal HireStatus As String, _
            ByVal EmployeeLetter As Boolean, _
            ByVal status As Integer, _
            ByVal CMstatus As Integer, _
            ByVal OverrideExpire As Boolean, _
            ByVal CertType As Integer, _
            ByVal CMCertType As Integer, _
            ByVal AppRecvdDate As DateTime, _
            ByVal OriginalIssuedDate As DateTime, _
            ByVal IssuedDate As DateTime, _
            ByVal LicenseExpirationDate As String, _
            ByVal ExceptionGrantedDate As DateTime, _
            ByVal CreatedBy As String, _
            ByVal CreatedOn As DateTime, _
            ByVal ModifiedBy As String, _
            ByVal ModifiedOn As DateTime, _
            ByVal Deleted As Boolean, _
            ByVal statusDesc As String, _
            ByVal CMstatusDesc As String, _
            ByVal certTypeDesc As String, _
            ByVal CMcertTypeDesc As String, _
            ByVal ExtensionDeadlineDate As DateTime, _
            ByVal ComplianceManager As Boolean, _
            ByVal IsLicensee As Boolean, _
            ByVal InitCertDate As String, _
            ByVal InitCertBy As Integer, _
            ByVal InitCertByDesc As String, _
            ByVal RetrainDate1 As String, _
            ByVal RetrainDate2 As String, _
            ByVal RetrainDate3 As String, _
            ByVal RevokeDate As String, _
            ByVal RetrainReqDate As String)
            onID = ID
            ostrTitle = Title
            ostrFirstName = FirstName
            ostrMiddleName = MiddleName
            ostrLastName = LastName
            ostrSuffix = Suffix
            ostrLicenseNumberPrefix = LicenseNumberPrefix
            If LicenseNumber.ToString.Length = 1 Then
                onLicenseNumber = "000" + LicenseNumber.ToString
            ElseIf LicenseNumber.ToString.Length = 2 Then
                onLicenseNumber = "00" + LicenseNumber.ToString
            ElseIf LicenseNumber.ToString.Length = 3 Then
                onLicenseNumber = "0" + LicenseNumber.ToString
            Else
                onLicenseNumber = LicenseNumber.ToString
            End If

            ostrEmailAddress = EmailAddress
            onAssociatedCompanyID = AssociatedCompanyID
            ostrHireStatus = HireStatus
            obolEmployeeLetter = EmployeeLetter
            onStatus = status
            ostrStatusDesc = statusDesc
            onCMStatus = CMstatus
            ostrCMStatusDesc = CMstatusDesc
            obolOverrideExpire = OverrideExpire
            onCertType = CertType
            onCMCertType = CMCertType
            ostrCertTypeDesc = certTypeDesc
            ostrCMCertTypeDesc = CMcertTypeDesc
            odtAppRecvdDate = AppRecvdDate
            odtOriginalIssuedDate = OriginalIssuedDate
            odtIssuedDate = IssuedDate
            odtExtensionDeadlineDate = ExtensionDeadlineDate
            obolComplianceManager = ComplianceManager
            obolIsLicensee = IsLicensee
            odtInitCertDate = InitCertDate
            onInitCertBy = InitCertBy
            ostrInitCertByDesc = InitCertByDesc
            odtRetrainDate1 = RetrainDate1
            odtRetrainDate2 = RetrainDate2
            odtRetrainDate3 = RetrainDate3
            odtRevokeDate = RevokeDate
            odtRetrainReqDate = RetrainReqDate
            odtLicenseExpirationDate = LicenseExpirationDate
            odtExceptionGrantedDate = ExceptionGrantedDate
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = ModifiedOn
            Me.Reset()
        End Sub
        'Sub New(ByVal drTemplate As DataRow)
        '    Try
        '        'onID = drTemplate.Item("ID")
        '        ''********************************************************
        '        ''
        '        '' Other private member variables for prior state here
        '        ''
        '        ''********************************************************
        '        'obolDeleted = drTemplate.Item("DELETED")
        '        'ostrCreatedBy = drTemplate.Item("CREATED_BY")
        '        'odtCreatedOn = drTemplate.Item("DATE_CREATED")
        '        'ostrModifiedBy = drTemplate.Item("LAST_EDITED_BY")
        '        'odtModifiedOn = drTemplate.Item("DATE_LAST_EDITED")
        '        Me.Reset()
        '    Catch ex As Exception
        '        MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nID >= 0 Then
                nID = onID
            End If
            strTitle = ostrTitle
            strFirstName = ostrFirstName
            strMiddleName = ostrMiddleName
            strLastName = ostrLastName
            strSuffix = ostrSuffix
            strLicenseNumberPrefix = ostrLicenseNumberPrefix
            nLicenseNumber = onLicenseNumber
            strEmailAddress = ostrEmailAddress
            nAssociatedCompanyID = onAssociatedCompanyID
            strHireStatus = ostrHireStatus
            bolEmployeeLetter = obolEmployeeLetter
            nStatus = onStatus
            nCMStatus = onCMStatus
            strStatusDesc = ostrStatusDesc
            strCMStatusDesc = ostrCMStatusDesc
            bolOverrideExpire = obolOverrideExpire
            nCertType = onCertType
            nCMCertType = onCMCertType
            strCertTypeDesc = ostrCertTypeDesc
            strCMCertTypeDesc = ostrCMCertTypeDesc
            dtAppRecvdDate = odtAppRecvdDate
            dtOriginalIssuedDate = odtOriginalIssuedDate
            dtIssuedDate = odtIssuedDate
            dtExtensionDeadlineDate = odtExtensionDeadlineDate
            bolComplianceManager = obolComplianceManager
            bolIsLicensee = obolIsLicensee
            dtInitCertDate = odtInitCertDate
            nInitCertBy = onInitCertBy
            strInitCertByDesc = ostrInitCertByDesc
            dtRetrainDate1 = odtRetrainDate1
            dtRetrainDate2 = odtRetrainDate2
            dtRetrainDate3 = odtRetrainDate3
            dtRevokeDate = odtRevokeDate
            dtRetrainReqDate = odtRetrainReqDate
            dtLicenseExpirationDate = odtLicenseExpirationDate
            dtExceptionGrantedDate = odtExceptionGrantedDate
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
            RaiseEvent LicenseeInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onID = nID
            ostrTitle = strTitle
            ostrFirstName = strFirstName
            ostrMiddleName = strMiddleName
            ostrLastName = strLastName
            ostrSuffix = strSuffix
            ostrLicenseNumberPrefix = strLicenseNumberPrefix
            onLicenseNumber = nLicenseNumber
            ostrEmailAddress = strEmailAddress
            onAssociatedCompanyID = nAssociatedCompanyID
            ostrHireStatus = strHireStatus
            obolEmployeeLetter = bolEmployeeLetter
            onStatus = nStatus
            ostrStatusDesc = strStatusDesc
            onCMStatus = nCMStatus
            ostrCMStatusDesc = strCMStatusDesc
            obolOverrideExpire = bolOverrideExpire
            onCertType = nCertType
            onCMCertType = nCMCertType
            ostrCertTypeDesc = strCertTypeDesc
            ostrCMCertTypeDesc = strCMCertTypeDesc
            odtAppRecvdDate = dtAppRecvdDate
            odtOriginalIssuedDate = dtOriginalIssuedDate
            odtIssuedDate = dtIssuedDate
            odtExtensionDeadlineDate = dtExtensionDeadlineDate
            obolComplianceManager = bolComplianceManager
            obolIsLicensee = bolIsLicensee
            odtInitCertDate = dtInitCertDate
            onInitCertBy = nInitCertBy
            odtRetrainDate1 = dtRetrainDate1
            odtRetrainDate2 = dtRetrainDate2
            odtRetrainDate3 = dtRetrainDate3
            odtRevokeDate = dtRevokeDate
            odtRetrainReqDate = dtRetrainReqDate
            odtLicenseExpirationDate = dtLicenseExpirationDate
            odtExceptionGrantedDate = dtExceptionGrantedDate
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

            bolIsDirty = (ostrTitle <> strTitle) Or _
            (ostrFirstName <> strFirstName) Or _
            (ostrMiddleName <> strMiddleName) Or _
            (ostrLastName <> strLastName) Or _
            (ostrSuffix <> strSuffix) Or _
            (ostrLicenseNumberPrefix <> strLicenseNumberPrefix) Or _
            (onLicenseNumber <> nLicenseNumber) Or _
            (ostrEmailAddress <> strEmailAddress) Or _
            (onAssociatedCompanyID <> nAssociatedCompanyID) Or _
            (ostrHireStatus <> strHireStatus) Or _
            (obolEmployeeLetter <> bolEmployeeLetter) Or _
            (onStatus <> nStatus) Or _
            (onCMStatus <> nCMStatus) Or _
            (obolOverrideExpire <> bolOverrideExpire) Or _
            (onCertType <> nCertType) Or _
            (onCMCertType <> nCMCertType) Or _
            (odtAppRecvdDate <> dtAppRecvdDate) Or _
            (odtOriginalIssuedDate <> dtOriginalIssuedDate) Or _
            (odtIssuedDate <> dtIssuedDate) Or _
            (odtExtensionDeadlineDate <> dtExtensionDeadlineDate) Or _
            (obolComplianceManager <> bolComplianceManager) Or _
            (obolIsLicensee <> bolIsLicensee) Or _
            (odtInitCertDate <> dtInitCertDate) Or _
            (onInitCertBy <> nInitCertBy) Or _
            (odtRetrainDate1 <> dtRetrainDate1) Or _
            (odtRetrainDate2 <> dtRetrainDate2) Or _
            (odtRetrainDate3 <> dtRetrainDate3) Or _
            (odtRevokeDate <> dtRevokeDate) Or _
            (odtRetrainReqDate <> dtRetrainReqDate) Or _
            (odtLicenseExpirationDate <> dtLicenseExpirationDate) Or _
            (odtExceptionGrantedDate <> dtExceptionGrantedDate) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent LicenseeInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onID = 0
            ostrTitle = String.Empty
            ostrFirstName = String.Empty
            ostrMiddleName = String.Empty
            ostrLastName = String.Empty
            ostrSuffix = String.Empty
            ostrLicenseNumberPrefix = String.Empty
            onLicenseNumber = String.Empty
            ostrEmailAddress = String.Empty
            onAssociatedCompanyID = 0
            ostrHireStatus = String.Empty
            obolEmployeeLetter = False
            onStatus = 0
            onCMStatus = 0
            ostrStatusDesc = String.Empty
            ostrCMStatusDesc = String.Empty
            obolOverrideExpire = False
            onCertType = 0
            onCMCertType = 0
            ostrCertTypeDesc = String.Empty
            ostrCMCertTypeDesc = String.Empty
            odtAppRecvdDate = CDate("01/01/0001")
            odtOriginalIssuedDate = CDate("01/01/0001")
            odtIssuedDate = CDate("01/01/0001")
            odtExtensionDeadlineDate = CDate("01/01/0001")
            obolComplianceManager = False
            obolIsLicensee = True
            odtInitCertDate = String.Empty
            onInitCertBy = 0
            ostrInitCertByDesc = String.Empty
            odtRetrainDate1 = String.Empty
            odtRetrainDate2 = String.Empty
            odtRetrainDate3 = String.Empty
            odtRevokeDate = String.Empty
            odtRetrainReqDate = String.Empty
            odtLicenseExpirationDate = String.Empty
            odtExceptionGrantedDate = CDate("01/01/0001")
            ostrCreatedBy = String.Empty
            odtCreatedOn = CDate("01/01/0001")
            ostrModifiedBy = String.Empty
            odtModifiedOn = CDate("01/01/0001")
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property APP_RECVD_DATE() As DateTime
            Get
                Return dtAppRecvdDate
            End Get
            Set(ByVal Value As DateTime)
                dtAppRecvdDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ASSOCATED_COMPANY_ID() As Integer
            Get
                Return nAssociatedCompanyID
            End Get
            Set(ByVal Value As Integer)
                nAssociatedCompanyID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CertType() As Integer
            Get
                Return nCertType
            End Get
            Set(ByVal Value As Integer)
                nCertType = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CMCertType() As Integer
            Get
                Return nCMCertType
            End Get
            Set(ByVal Value As Integer)
                nCMCertType = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CertTypeDesc() As String
            Get
                Return strCertTypeDesc
            End Get
            Set(ByVal Value As String)
                strCertTypeDesc = Value
            End Set
        End Property
        Public Property CMCertTypeDesc() As String
            Get
                Return strCMCertTypeDesc
            End Get
            Set(ByVal Value As String)
                strCMCertTypeDesc = Value
            End Set
        End Property
        Public Property DELETED() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolDeleted = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property EMAIL_ADDRESS() As String
            Get
                Return strEmailAddress
            End Get
            Set(ByVal Value As String)
                strEmailAddress = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property EMPLOYEE_LETTER() As Boolean
            Get
                Return bolEmployeeLetter
            End Get
            Set(ByVal Value As Boolean)
                bolEmployeeLetter = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property EXCEPT_GRANT_DATE() As DateTime
            Get
                Return dtExceptionGrantedDate
            End Get
            Set(ByVal Value As DateTime)
                dtExceptionGrantedDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FIRST_NAME() As String
            Get
                Return strFirstName
            End Get
            Set(ByVal Value As String)
                strFirstName = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property HIRE_STATUS() As String
            Get
                Return strHireStatus
            End Get
            Set(ByVal Value As String)
                strHireStatus = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ID() As Int64
            Get
                Return nID
            End Get
            Set(ByVal Value As Int64)
                nID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get
            Set(ByVal Value As Boolean)
                bolIsDirty = Value
            End Set
        End Property
        Public Property ISSUED_DATE() As Date
            Get
                Return dtIssuedDate
            End Get
            Set(ByVal Value As Date)
                dtIssuedDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property EXTENSION_DEADLINE_DATE() As Date
            Get
                Return dtExtensionDeadlineDate
            End Get
            Set(ByVal Value As Date)
                dtExtensionDeadlineDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property COMPLIANCEMANAGER() As Boolean
            Get
                Return bolComplianceManager
            End Get
            Set(ByVal Value As Boolean)
                bolComplianceManager = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ISLICENSEE() As Boolean
            Get
                Return bolIsLicensee
            End Get
            Set(ByVal Value As Boolean)
                bolIsLicensee = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property INITCERTDATE() As String
            Get
                Return dtInitCertDate
            End Get
            Set(ByVal Value As String)
                dtInitCertDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property INITCERTBY() As Integer
            Get
                Return nInitCertBy
            End Get
            Set(ByVal Value As Integer)
                nInitCertBy = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property INITCERTBYDESC() As String
            Get
                Return strInitCertByDesc
            End Get
            Set(ByVal Value As String)
                strInitCertByDesc = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property RETRAINDATE1() As String
            Get
                Return dtRetrainDate1
            End Get
            Set(ByVal Value As String)
                dtRetrainDate1 = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property RETRAINDATE2() As String
            Get
                Return dtRetrainDate2
            End Get
            Set(ByVal Value As String)
                dtRetrainDate2 = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property RETRAINDATE3() As String
            Get
                Return dtRetrainDate3
            End Get
            Set(ByVal Value As String)
                dtRetrainDate3 = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property REVOKEDATE() As String
            Get
                Return dtRevokeDate
            End Get
            Set(ByVal Value As String)
                dtRevokeDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property RETRAINREQDATE() As String
            Get
                Return dtRetrainReqDate
            End Get
            Set(ByVal Value As String)
                dtRetrainReqDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LAST_NAME() As String
            Get
                Return strLastName
            End Get
            Set(ByVal Value As String)
                strLastName = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LICENSE_EXPIRE_DATE() As String
            Get
                Return dtLicenseExpirationDate
            End Get
            Set(ByVal Value As String)
                dtLicenseExpirationDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property MIDDLE_NAME() As String
            Get
                Return strMiddleName
            End Get
            Set(ByVal Value As String)
                strMiddleName = Value
                Me.CheckDirty()
            End Set
        End Property
        Public ReadOnly Property FULLNAME() As String
            Get
                Return strTitle + " " + strFirstName + " " + strMiddleName + " " + strLastName + " " + strSuffix
            End Get
        End Property
        Public Property ORIGIN_ISSUED_DATE() As Date
            Get
                Return dtOriginalIssuedDate
            End Get
            Set(ByVal Value As Date)
                dtOriginalIssuedDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property OVERRIDE_EXPIRE() As Boolean
            Get
                Return bolOverrideExpire
            End Get
            Set(ByVal Value As Boolean)
                bolOverrideExpire = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property STATUS() As Integer
            Get
                Return nStatus
            End Get
            Set(ByVal Value As Integer)
                nStatus = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CMSTATUS() As Integer
            Get
                Return nCMStatus
            End Get
            Set(ByVal Value As Integer)
                nCMStatus = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property StatusDesc() As String
            Get
                Return strStatusDesc
            End Get
            Set(ByVal Value As String)
                strStatusDesc = Value
            End Set
        End Property
        Public Property CMStatusDesc() As String
            Get
                Return strCMStatusDesc
            End Get
            Set(ByVal Value As String)
                strCMStatusDesc = Value
            End Set
        End Property
        Public Property SUFFIX() As String
            Get
                Return strSuffix
            End Get
            Set(ByVal Value As String)
                strSuffix = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TITLE() As String
            Get
                Return strTitle
            End Get
            Set(ByVal Value As String)
                strTitle = Value
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
        Public Property DATE_LAST_EDITED() As Date
            Get
                Return dtModifiedOn
            End Get
            Set(ByVal Value As Date)
                dtModifiedOn = Value
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
        Public Property LICENSEE_NUMBER_PREFIX() As String
            Get
                Return strLicenseNumberPrefix
            End Get
            Set(ByVal Value As String)
                strLicenseNumberPrefix = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LICENSEE_NUMBER() As String
            Get
                Return nLicenseNumber
            End Get
            Set(ByVal Value As String)
                nLicenseNumber = Value
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

