'-------------------------------------------------------------------------------
' MUSTER.Info.CompanyInfo
' '
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MKK        05/24/05  Original class definition
'
' Function          Description
' New()             Instantiates an empty CompanyInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated CompanyInfo object
' New(dr)           Instantiates a populated CompanyInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class CompanyInfo
#Region "Public Events"
        Public Event CompanyInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nID As Int64
        Private strCertResp As String
        Private strCompanyName As String
        Private dtFinRespEndDate As Date
        Private strEmailAddress As String
        Private strProEngin As String
        Private nProEnginID As Integer
        Private strProEnginNumber As String
        Private strProEnginEmail As String

        Private dtProEnginAppApprovalDate As Date
        Private dtProEnginLiabilityExpiration As Date
        Private strProGeolo As String
        Private nProGeoloAddID As Integer
        Private strProGeoloNumber As String
        Private strProGeoloEmail As String

        Private bolCTIAC As Boolean
        Private bolCTC As Boolean
        Private bolPTTT As Boolean
        Private bolLDC As Boolean
        Private bolUSTSE As Boolean
        Private bolWL As Boolean
        Private bolTST As Boolean
        Private bolErac As Boolean
        Private bolIrac As Boolean
        Private bolEC As Boolean
        Private bolED As Boolean
        Private bolTL As Boolean
        Private bolCE As Boolean
        Private bolCM As Boolean
        Private bolActive As Boolean
        Private bolDeleted As Boolean
        Private strCreatedBy As String
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String
        Private dtModifiedOn As DateTime

        Private onID As Int64
        Private ostrCertResp As String
        Private ostrCompanyName As String
        Private odtFinRespEndDate As Date
        Private ostrEmailAddress As String
        Private ostrProEngin As String
        Private onProEnginID As Integer
        Private ostrProEnginNumber As String
        Private ostrProEnginEmail As String

        Private odtProEnginAppApprovalDate As Date
        Private odtProEnginLiabilityExpiration As Date
        Private ostrProGeolo As String
        Private onProGeoloAddID As Integer
        Private ostrProGeoloNumber As String
        Private ostrProGeoloEmail As String

        Private obolCTIAC As Boolean
        Private obolCTC As Boolean
        Private obolPTTT As Boolean
        Private obolLDC As Boolean
        Private obolUSTSE As Boolean
        Private obolWL As Boolean
        Private obolTST As Boolean
        Private obolErac As Boolean
        Private obolIrac As Boolean
        Private obolEC As Boolean
        Private obolED As Boolean
        Private obolTL As Boolean
        Private obolCE As Boolean
        Private obolCM As Boolean
        Private obolActive As Boolean
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
        '********************************************************
        '
        ' Update signature to reflect other attributes persisted
        '   by the object.
        '
        '********************************************************
        Sub New(ByVal ID As Int64, _
            ByVal CertRespon As String, _
            ByVal CompanyName As String, _
            ByVal FinRespEndDate As Date, _
            ByVal EmailAddress As String, _
            ByVal ProEngin As String, _
            ByVal ProEnginID As Integer, _
            ByVal ProEnginNumber As String, _
            ByVal ProEnginAppApprovalDate As Date, _
            ByVal ProEnginLiabilityExpiration As Date, _
            ByVal ProGeolo As String, _
            ByVal ProGeoloAddID As Integer, _
            ByVal ProGeoloNumber As String, _
            ByVal CTIAC As Boolean, _
            ByVal CTC As Boolean, _
            ByVal PTTT As Boolean, _
            ByVal LDC As Boolean, _
            ByVal USTSE As Boolean, _
            ByVal WL As Boolean, _
            ByVal TST As Boolean, _
            ByVal Erac As Boolean, _
            ByVal Irac As Boolean, _
            ByVal EC As Boolean, _
            ByVal ED As Boolean, _
            ByVal TL As Boolean, _
            ByVal CE As Boolean, _
            ByVal CM As Boolean, _
            ByVal Active As Boolean, _
            ByVal Deleted As Boolean, _
            ByVal CreatedBy As String, _
            ByVal CreatedOn As Date, _
            ByVal ModifiedBy As String, _
            ByVal LastEdited As Date, _
            ByVal ProEnginEmail As String, _
            ByVal ProGeoloEmail As String)
            onID = ID
            ostrCertResp = CertRespon
            ostrCompanyName = CompanyName
            odtFinRespEndDate = FinRespEndDate
            ostrEmailAddress = EmailAddress
            ostrProEngin = ProEngin
            onProEnginID = ProEnginID
            ostrProEnginNumber = ProEnginNumber
            ostrProEnginEmail = ProEnginEmail

            odtProEnginAppApprovalDate = ProEnginAppApprovalDate
            odtProEnginLiabilityExpiration = ProEnginLiabilityExpiration
            ostrProGeolo = ProGeolo
            onProGeoloAddID = ProGeoloAddID
            ostrProGeoloNumber = ProGeoloNumber
            ostrProGeoloEmail = ProGeoloEmail

            obolCTIAC = CTIAC
            obolCTC = CTC
            obolPTTT = PTTT
            obolLDC = LDC
            obolUSTSE = USTSE
            obolWL = WL
            obolTST = TST
            obolErac = Erac
            obolIrac = Irac
            obolEC = EC
            obolED = ED
            obolTL = TL
            obolCE = CE
            obolCM = CM
            obolActive = Active
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = LastEdited
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
            strCertResp = ostrCertResp
            strCompanyName = ostrCompanyName
            dtFinRespEndDate = odtFinRespEndDate
            strEmailAddress = ostrEmailAddress
            strProEngin = ostrProEngin
            nProEnginID = onProEnginID
            strProEnginNumber = ostrProEnginNumber
            strProEnginEmail = ostrProEnginEmail

            dtProEnginAppApprovalDate = odtProEnginAppApprovalDate
            dtProEnginLiabilityExpiration = odtProEnginLiabilityExpiration
            strProGeolo = ostrProGeolo
            nProGeoloAddID = onProGeoloAddID
            strProGeoloNumber = ostrProGeoloNumber
            strProGeoloEmail = ostrProGeoloEmail

            bolCTIAC = obolCTIAC
            bolCTC = obolCTC
            bolPTTT = obolPTTT
            bolLDC = obolLDC
            bolUSTSE = obolUSTSE
            bolWL = obolWL
            bolTST = obolTST
            bolErac = obolErac
            bolIrac = obolIrac
            bolEC = obolEC
            bolED = obolED
            bolTL = obolTL
            bolCE = obolCE
            bolCM = obolCM
            bolActive = obolActive
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
            RaiseEvent CompanyInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onID = nID
            ostrCertResp = strCertResp
            ostrCompanyName = strCompanyName
            odtFinRespEndDate = dtFinRespEndDate
            ostrEmailAddress = strEmailAddress
            ostrProEngin = strProEngin
            onProEnginID = nProEnginID
            ostrProEnginNumber = strProEnginNumber
            ostrProEnginEmail = strProEnginEmail

            odtProEnginAppApprovalDate = dtProEnginAppApprovalDate
            odtProEnginLiabilityExpiration = dtProEnginLiabilityExpiration
            ostrProGeolo = strProGeolo
            onProGeoloAddID = nProGeoloAddID
            ostrProGeoloNumber = strProGeoloNumber
            ostrProGeoloEmail = strProGeoloEmail

            obolCTIAC = bolCTIAC
            obolCTC = bolCTC
            obolPTTT = bolPTTT
            obolLDC = bolLDC
            obolUSTSE = bolUSTSE
            obolWL = bolWL
            obolTST = bolTST
            obolErac = bolErac
            obolIrac = bolIrac
            obolEC = bolEC
            obolED = bolED
            obolTL = bolTL
            obolCE = bolCE
            obolCM = bolCM
            obolActive = bolActive
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

            bolIsDirty = (strCertResp <> ostrCertResp) Or _
            (strCompanyName <> ostrCompanyName) Or _
            (dtFinRespEndDate <> odtFinRespEndDate) Or _
            (strEmailAddress <> ostrEmailAddress) Or _
            (strProEngin <> ostrProEngin) Or _
            (nProEnginID <> onProEnginID) Or _
            (strProEnginNumber <> ostrProEnginNumber) Or _
            (strProEnginEmail <> ostrProEnginEmail) Or _
            (dtProEnginAppApprovalDate <> odtProEnginAppApprovalDate) Or _
            (dtProEnginLiabilityExpiration <> odtProEnginLiabilityExpiration) Or _
            (strProGeolo <> ostrProGeolo) Or _
            (nProGeoloAddID <> onProGeoloAddID) Or _
            (strProGeoloNumber <> ostrProGeoloNumber) Or _
            (strProGeoloEmail <> ostrProGeoloEmail) Or _
            (bolCTIAC <> obolCTIAC) Or _
            (bolCTC <> obolCTC) Or _
            (bolPTTT <> obolPTTT) Or _
            (bolLDC <> obolLDC) Or _
            (bolUSTSE <> obolUSTSE) Or _
            (bolWL <> obolWL) Or _
            (bolTST <> obolTST) Or _
            (bolErac <> obolErac) Or _
            (bolIrac <> obolIrac) Or _
            (bolEC <> obolEC) Or _
            (bolED <> obolED) Or _
            (bolTL <> obolTL) Or _
            (bolCE <> obolCE) Or _
            (bolCM <> obolCM) Or _
            (bolActive <> obolActive) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn) Or bolDeleted

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent CompanyInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onID = 0
            ostrCertResp = 0
            ostrCompanyName = String.Empty
            odtFinRespEndDate = CDate("01/01/0001")
            ostrEmailAddress = String.Empty
            ostrProEngin = String.Empty
            onProEnginID = 0
            ostrProEnginNumber = String.Empty
            ostrProEnginEmail = String.Empty

            odtProEnginAppApprovalDate = CDate("01/01/0001")
            odtProEnginLiabilityExpiration = CDate("01/01/0001")
            ostrProGeolo = String.Empty
            onProGeoloAddID = 0
            ostrProGeoloNumber = String.Empty
            ostrProGeoloEmail = String.Empty

            obolCTIAC = False
            obolCTC = False
            obolPTTT = False
            obolLDC = False
            obolUSTSE = False
            obolWL = False
            obolTST = False
            obolErac = False
            obolIrac = False
            obolEC = False
            obolED = False
            obolTL = False
            obolCE = False
            obolCM = False
            obolActive = False
            bolDeleted = False
            ostrCreatedBy = String.Empty
            odtCreatedOn = CDate("01/01/0001")
            ostrModifiedBy = String.Empty
            odtModifiedOn = CDate("01/01/0001")
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
        Public Property ACTIVE() As Boolean
            Get
                Return bolActive
            End Get
            Set(ByVal Value As Boolean)
                bolActive = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CE() As Boolean
            Get
                Return bolCE
            End Get
            Set(ByVal Value As Boolean)
                bolCE = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CM() As Boolean
            Get
                Return bolCM
            End Get
            Set(ByVal Value As Boolean)
                bolCM = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CERT_RESPON() As String
            Get
                Return strCertResp
            End Get
            Set(ByVal Value As String)
                strCertResp = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property COMPANY_NAME() As String
            Get
                Return strCompanyName
            End Get
            Set(ByVal Value As String)
                strCompanyName = Value
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
        Public Property CTC() As Boolean
            Get
                Return bolCTC
            End Get
            Set(ByVal Value As Boolean)
                bolCTC = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CTIAC() As Boolean
            Get
                Return bolCTIAC
            End Get
            Set(ByVal Value As Boolean)
                bolCTIAC = Value
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
        Public Property DELETED() As Boolean
            Get
                Return (bolDeleted)
            End Get
            Set(ByVal Value As Boolean)
                bolDeleted = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property EC() As Boolean
            Get
                Return bolEC
            End Get
            Set(ByVal Value As Boolean)
                bolEC = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ED() As Boolean
            Get
                Return bolED
            End Get
            Set(ByVal Value As Boolean)
                bolED = Value
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
        Public Property ERAC() As Boolean
            Get
                Return bolErac
            End Get
            Set(ByVal Value As Boolean)
                bolErac = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FIN_RESP_END_DATE() As Date
            Get
                Return dtFinRespEndDate
            End Get
            Set(ByVal Value As Date)
                dtFinRespEndDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IRAC() As Boolean
            Get
                Return bolIrac
            End Get
            Set(ByVal Value As Boolean)
                bolIrac = Value
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
        Public Property LAST_EDITED_BY() As String
            Get
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LDC() As Boolean
            Get
                Return bolLDC
            End Get
            Set(ByVal Value As Boolean)
                bolLDC = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PRO_ENGIN() As String
            Get
                Return strProEngin
            End Get
            Set(ByVal Value As String)
                strProEngin = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PRO_ENGIN_ADD_ID() As Integer
            Get
                Return nProEnginID
            End Get
            Set(ByVal Value As Integer)
                nProEnginID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PRO_ENGIN_APP_APRV_DATE() As Date
            Get
                Return dtProEnginAppApprovalDate
            End Get
            Set(ByVal Value As Date)
                dtProEnginAppApprovalDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PRO_ENGIN_LIABIL_DATE() As Date
            Get
                Return dtProEnginLiabilityExpiration
            End Get
            Set(ByVal Value As Date)
                dtProEnginLiabilityExpiration = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PRO_ENGIN_NUMBER() As String
            Get
                Return strProEnginNumber
            End Get
            Set(ByVal Value As String)
                strProEnginNumber = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property PRO_ENGIN_EMAIL() As String
            Get
                Return strProEnginEmail
            End Get
            Set(ByVal Value As String)
                strProEnginEmail = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property PRO_GEOLO() As String
            Get
                Return strProGeolo
            End Get
            Set(ByVal Value As String)
                strProGeolo = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PRO_GEOLO_ADD_ID() As Integer
            Get
                Return nProGeoloAddID
            End Get
            Set(ByVal Value As Integer)
                nProGeoloAddID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PRO_GEOLO_NUMBER() As String
            Get
                Return strProGeoloNumber
            End Get
            Set(ByVal Value As String)
                strProGeoloNumber = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property PRO_GEOLO_EMAIL() As String
            Get
                Return strProGeoloEmail
            End Get
            Set(ByVal Value As String)
                strProGeoloEmail = Value
                Me.CheckDirty()
            End Set
        End Property


        Public Property PTTT() As Boolean
            Get
                Return bolPTTT
            End Get
            Set(ByVal Value As Boolean)
                bolPTTT = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TL() As Boolean
            Get
                Return bolTL
            End Get
            Set(ByVal Value As Boolean)
                bolTL = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TST() As Boolean
            Get
                Return bolTST
            End Get
            Set(ByVal Value As Boolean)
                bolTST = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property USTSE() As Boolean
            Get
                Return bolUSTSE
            End Get
            Set(ByVal Value As Boolean)
                bolUSTSE = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property WL() As Boolean
            Get
                Return bolWL
            End Get
            Set(ByVal value As Boolean)
                bolWL = value
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

